VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmStuffQualityList 
   Caption         =   "卫材质量管理"
   ClientHeight    =   8010
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11745
   Icon            =   "frmStuffQualityList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8010
   ScaleWidth      =   11745
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   7650
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffQualityList.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15637
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
   Begin VB.Frame fraFilter 
      Height          =   6885
      Left            =   45
      TabIndex        =   16
      Top             =   765
      Width           =   2505
      Begin VB.CheckBox chkVerify 
         Caption         =   "处理时间"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   2625
         Width           =   1455
      End
      Begin VB.CheckBox chkCheck 
         Caption         =   "登记时间"
         Height          =   180
         Left            =   135
         TabIndex        =   3
         Top             =   1320
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1965
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤"
         Height          =   300
         Left            =   1215
         TabIndex        =   15
         Top             =   6420
         Width           =   855
      End
      Begin VB.PictureBox picVerify 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   135
         ScaleHeight     =   735
         ScaleWidth      =   2295
         TabIndex        =   28
         Top             =   3240
         Width           =   2295
         Begin MSComCtl2.DTPicker dtpVerifyBegin 
            Height          =   315
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   114229251
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpVerifyEnd 
            Height          =   315
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   114229251
            CurrentDate     =   36263
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   29
            Top             =   120
            Width           =   180
         End
      End
      Begin VB.ComboBox cboVerifyDate 
         Enabled         =   0   'False
         Height          =   300
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2880
         Width           =   1965
      End
      Begin VB.PictureBox picCheck 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   135
         ScaleHeight     =   735
         ScaleWidth      =   2295
         TabIndex        =   26
         Top             =   1920
         Width           =   2295
         Begin MSComCtl2.DTPicker dtpCheckBegin 
            Height          =   315
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   114229251
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtpCheckEnd 
            Height          =   315
            Left            =   0
            TabIndex        =   6
            Top             =   360
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   114229251
            CurrentDate     =   36263
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   2040
            TabIndex        =   27
            Top             =   120
            Width           =   180
         End
      End
      Begin VB.ComboBox cboCheckDate 
         Height          =   300
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   1965
      End
      Begin VB.TextBox txtStuff 
         Height          =   300
         Left            =   120
         TabIndex        =   12
         Top             =   4800
         Width           =   1695
      End
      Begin VB.CommandButton CmdStuff 
         Caption         =   "…"
         Height          =   300
         Left            =   1800
         TabIndex        =   25
         Top             =   4770
         Width           =   255
      End
      Begin VB.TextBox txtProvider 
         Height          =   300
         Left            =   135
         TabIndex        =   11
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton cmdProvider 
         Caption         =   "…"
         Height          =   300
         Left            =   1800
         TabIndex        =   24
         Top             =   4170
         Width           =   255
      End
      Begin VB.TextBox txtCheck 
         Height          =   300
         Left            =   135
         TabIndex        =   13
         Top             =   5400
         Width           =   1965
      End
      Begin VB.TextBox txtVerify 
         Height          =   300
         Left            =   135
         TabIndex        =   14
         Top             =   6000
         Width           =   1965
      End
      Begin VB.OptionButton opt单位 
         Caption         =   "包装单位"
         Height          =   255
         Index           =   1
         Left            =   1245
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton opt单位 
         Caption         =   "散装单位"
         Height          =   255
         Index           =   0
         Left            =   135
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label lblStore 
         AutoSize        =   -1  'True
         Caption         =   "库房"
         Height          =   180
         Left            =   135
         TabIndex        =   34
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lblStuff 
         AutoSize        =   -1  'True
         Caption         =   "材料"
         Height          =   180
         Left            =   165
         TabIndex        =   33
         Top             =   4560
         Width           =   360
      End
      Begin VB.Label LblProvider 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "供药单位"
         Height          =   180
         Left            =   135
         TabIndex        =   32
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label lblCheck 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
         Height          =   180
         Left            =   135
         TabIndex        =   31
         Top             =   5160
         Width           =   540
      End
      Begin VB.Label lblVerify 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "处理人"
         Height          =   180
         Left            =   135
         TabIndex        =   30
         Top             =   5760
         Width           =   540
      End
   End
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   6405
      Left            =   2655
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6405
      ScaleWidth      =   45
      TabIndex        =   23
      Top             =   930
      Width           =   50
   End
   Begin VB.PictureBox pic水平分割 
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   50
      Left            =   2835
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8775
      TabIndex        =   22
      Top             =   3045
      Width           =   8775
   End
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   9495
      Top             =   375
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
            Picture         =   "frmStuffQualityList.frx":70E6
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":7302
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":751E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":7738
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":7952
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":7B6C
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":7D86
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":8480
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":869A
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":87F4
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":8A10
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":8C2C
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   10215
      Top             =   360
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
            Picture         =   "frmStuffQualityList.frx":8E46
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":9062
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":927E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":9498
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":96B4
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":98CE
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":9AE8
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":A1E2
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":A3FC
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":A556
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":A772
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffQualityList.frx":A98E
            Key             =   "View"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   3960
      Left            =   2835
      TabIndex        =   18
      Top             =   3300
      Width           =   8790
      _cx             =   15505
      _cy             =   6985
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
      Rows            =   1
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffQualityList.frx":ABA8
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
      Height          =   1800
      Left            =   2805
      TabIndex        =   19
      Top             =   1005
      Width           =   8850
      _cx             =   15610
      _cy             =   3175
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
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmStuffQualityList.frx":AD6D
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
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11745
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   6000
      MinHeight1      =   720
      Width1          =   6210
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   11625
         _ExtentX        =   20505
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
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
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
               Enabled         =   0   'False
               Caption         =   "修改"
               Key             =   "Modify"
               Description     =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
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
               Enabled         =   0   'False
               Caption         =   "处理"
               Key             =   "Verify"
               Description     =   "处理"
               Object.ToolTipText     =   "处理"
               Object.Tag             =   "处理"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助主题"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmStuffQualityList.frx":AE5D
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "预览(&V)"
      End
      Begin VB.Menu mnuFileSplit 
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
      End
      Begin VB.Menu mnuEditUpdate 
         Caption         =   "修改(&M)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "处理(&O)"
         Enabled         =   0   'False
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
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
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmStuffQualityList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mFMT As g_FmtString '小数位数的格式串
Private mintUnit As Integer '0-散装单位，1-包装单位

Private Sub cboCheckDate_Click()
    With cboCheckDate
        If .Text = "自定义" Then
            picCheck.Visible = True
        Else
            picCheck.Visible = False
        End If
    End With
End Sub

Private Sub cboVerifyDate_Click()
    With cboVerifyDate
        If .Text = "自定义" Then
            picVerify.Visible = True
        Else
            picVerify.Visible = False
        End If
    End With
End Sub

Private Sub SetCboDate()
    '往cbo控件中添加数据
    With cboCheckDate
        .AddItem "一周内"
        .AddItem "一月内"
        .AddItem "三月内"
        .AddItem "半年内"
        .AddItem "一年内"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    With cboVerifyDate
        .AddItem "一周内"
        .AddItem "一月内"
        .AddItem "三月内"
        .AddItem "半年内"
        .AddItem "一年内"
        .AddItem "自定义"
        .ListIndex = 0
    End With
    
    dtpCheckEnd = sys.Currentdate
    dtpVerifyEnd = dtpCheckEnd
    dtpCheckBegin = DateAdd("d", -7, dtpCheckEnd)
    dtpVerifyBegin = dtpCheckBegin
End Sub

Private Sub chkCheck_Click()
    If chkCheck.Value = 1 Then
        cboCheckDate.Enabled = True
    Else
        cboCheckDate.Enabled = False
    End If
End Sub

Private Sub chkVerify_Click()
    If chkVerify.Value = 1 Then
        cboVerifyDate.Enabled = True
    Else
        cboVerifyDate.Enabled = False
    End If
End Sub

Private Sub cmdFilter_Click()
    If chkCheck.Value = 0 And chkVerify.Value = 0 Then
        MsgBox "必须选择一个时间范围进行查询！", vbInformation, gstrSysName
        chkCheck.SetFocus
        Exit Sub
    End If
    
    Call GetList
End Sub

Private Sub cmdProvider_Click()
    Dim rsProvider As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHandle
    vRect = zlControl.GetControlRect(txtProvider.hwnd) '获取位置
    dblLeft = vRect.Left
    dblTop = vRect.Top - 700
    
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 供应商 " & _
              "Where (站点 = [1] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
              "  And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is null connect by prior ID =上级ID order by level,ID"
    Set rsProvider = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
                        True, dblLeft, dblTop, 1000, blnCancel, False, True, gstrNodeNo)
    If rsProvider Is Nothing Then
        Exit Sub
    Else
        txtProvider.Text = rsProvider!名称
        txtProvider.Tag = rsProvider!Id
    End If
    
    txtStuff.SetFocus
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdStuff_Click()
    '获取材料信息
    Dim RecReturn As Recordset
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 1, 0, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    If RecReturn.RecordCount = 0 Then Exit Sub
    txtStuff = "[" & RecReturn!编码 & "]" & RecReturn!名称
    txtStuff.Tag = RecReturn!材料ID
    
    txtCheck.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, "卫材质量管理"
End Sub

Private Sub opt单位_Click(Index As Integer)
    mintUnit = IIf(Index = 0, 0, 1)
End Sub

Private Sub txtCheck_GotFocus()
    zlControl.TxtSelAll txtCheck
End Sub

Private Sub txtCheck_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub txtProvider_GotFocus()
    zlControl.TxtSelAll txtProvider
End Sub

Private Sub txtStuff_Change()
    If Trim(txtStuff.Text) = "" Then txtStuff.Tag = 0
End Sub

Private Sub txtStuff_GotFocus()
    zlControl.TxtSelAll txtStuff
End Sub

Private Sub txtVerify_GotFocus()
    zlControl.TxtSelAll txtVerify
End Sub

Private Sub txtVerify_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call OS.PressKey(vbKeyTab)
End Sub

Private Sub txtStuff_KeyDown(KeyCode As Integer, Shift As Integer)
    '获取材料信息
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtStuff.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fraFilter.Left + txtStuff.Left + 100
    sngTop = Me.Top + fraFilter.Top + txtStuff.Top + txtStuff.Height + Me.Height - Me.ScaleHeight - 100
    If sngTop + 4300 > Screen.Height Then
        sngTop = sngTop - txtStuff.Height - 3680
    End If
    
    strKey = UCase(Trim(txtStuff.Text))
    If Mid(strKey, 1, 1) = "[" Then
        If InStr(2, strKey, "]") <> 0 Then
            strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
        Else
            strKey = Mid(strKey, 2)
        End If
    End If
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), strKey, sngLeft, sngTop, txtStuff.Width, txtStuff.Height)
    If RecReturn.RecordCount = 0 Then Exit Sub
    txtStuff = "[" & RecReturn!编码 & "]" & RecReturn!名称
    txtStuff.Tag = RecReturn!材料ID
    
    txtCheck.SetFocus
End Sub

Private Sub Form_Load()
    Dim strValue As String
    
    strValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Stuff\材料质量管理", "包装单位", 0)
    opt单位(Val(strValue)).Value = True
    
    Call CheckDepend
    Call SetCboDate
    Call SetVisible
    
    RestoreWinState Me, App.ProductName, "卫材质量管理"
End Sub

Private Sub SetVisible()
    '根据权限设置菜单/工具栏/表格
    If InStr(1, ";" & gstrPrivs & ";", ";质量登记;") = 0 Then
        mnuEditAdd.Visible = False
        tlbTool.Buttons("Add").Visible = False
        mnuEditUpdate.Visible = False
        tlbTool.Buttons("Modify").Visible = False
        mnuEditDelete.Visible = False
        tlbTool.Buttons("Delete").Visible = False
        tlbTool.Buttons(3).Visible = False
    End If
    If InStr(1, ";" & gstrPrivs & ";", ";质量审核;") = 0 Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
        tlbTool.Buttons(9).Visible = False
    End If
    If InStr(1, ";" & gstrPrivs & ";", ";查看成本价;") = 0 Then
        vsfDetail.ColHidden(vsfDetail.ColIndex("成本价")) = True
        vsfDetail.ColHidden(vsfDetail.ColIndex("成本金额")) = True
    End If
    If InStr(1, ";" & gstrPrivs & ";", ";单据打印;") = 0 Then
        mnuFilePrint.Visible = False
        tlbTool.Buttons("Print").Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With cbrTool
        .Width = Me.ScaleWidth
    End With
    
    With fraFilter
        If cbrTool.Bands(1).Visible = True Then
            .Move 65, cbrTool.Top + cbrTool.Height - 30, 2500, Me.ScaleHeight - cbrTool.Top - cbrTool.Height - IIf(staThis.Visible = True, staThis.Height - 30, 0)
        Else
            .Move 65, 30, 2500, Me.ScaleHeight - IIf(staThis.Visible = True, staThis.Height + 30, 30)
        End If
    End With
    
    With picFilter
        .Move fraFilter.Left + fraFilter.Width, fraFilter.Top, picFilter.Width, fraFilter.Height
    End With
    
    With vsfList
        .Move picFilter.Left + picFilter.Width, fraFilter.Top + 85, Me.ScaleWidth - picFilter.Width - picFilter.Left - 20, picFilter.Height / 3
    End With
    
    With pic水平分割
        .Move vsfList.Left, vsfList.Top + vsfList.Height, vsfList.Width, .Height
    End With
    
    With vsfDetail
        .Move picFilter.Left + picFilter.Width, vsfList.Top + vsfList.Height + pic水平分割.Height, Me.ScaleWidth - picFilter.Width - picFilter.Left - 20, Me.ScaleHeight - vsfList.Top - vsfList.Height - pic水平分割.Height - 360 + IIf(staThis.Visible = True, 0, staThis.Height - 30)
    End With
    
End Sub

Private Sub mnuEditAdd_Click()
    '新增
    If cboStock.ListIndex = -1 Then
        MsgBox "必须选择一个具有药库或者药房性质的部门！", vbInformation, gstrSysName
        Exit Sub
    End If
    frmStuffQualityCard.ShowMe 1, Me, cboStock.ItemData(cboStock.ListIndex), 0, mintUnit
    Call GetList
End Sub

Private Sub mnuEditDelete_Click()
    '删除
    Dim lng质量id As Long
    Dim strNo As String
    
    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("id")) <> "" Then
                lng质量id = Val(.TextMatrix(.Row, .ColIndex("id")))
                strNo = .TextMatrix(.Row, .ColIndex("NO"))
            Else
                lng质量id = 0
            End If
        Else
            lng质量id = 0
        End If
    End With
    
    If lng质量id <> 0 Then
        Call DeleteStuff(lng质量id, strNo)
        Call GetList
    End If
End Sub

Private Sub mnuEditUpdate_Click()
    '修改
    Dim lng质量id As Long
    
    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("id")) <> "" Then
                lng质量id = Val(.TextMatrix(.Row, .ColIndex("id")))
            Else
                lng质量id = 0
            End If
        Else
            lng质量id = 0
        End If
    End With
    
    If lng质量id <> 0 Then
        frmStuffQualityCard.ShowMe 2, Me, cboStock.ItemData(cboStock.ListIndex), lng质量id, mintUnit
        Call GetList
    End If
End Sub

Private Sub mnuEditVerify_Click()
    '处理
    Dim lng质量id As Long
    
    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("id")) <> "" Then
                lng质量id = Val(.TextMatrix(.Row, .ColIndex("id")))
            Else
                lng质量id = 0
            End If
        Else
            lng质量id = 0
        End If
    End With

    If lng质量id <> 0 Then
        frmStuffQualityCard.ShowMe 3, Me, cboStock.ItemData(cboStock.ListIndex), lng质量id, mintUnit
        Call GetList
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    '关于
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '帮助主题
    Call ShowHelp(App.ProductName, Me.hwnd, "材料质量管理")
End Sub

Private Sub mnuHelpWebForum_Click()
    '中联论坛
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewStatus_Click()
    '状态栏
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    '标准按钮
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    '文本标签
    Dim intCount As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
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

Private Sub mnufileexit_Click()
    '退出
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9Stuff\材料质量管理", "包装单位", IIf(opt单位(0).Value = True, 0, 1))
    Unload Me
End Sub

Private Sub picFilter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    With picFilter
        If .Left + x <= 1000 Then Exit Sub
        If .Left + x >= 5000 Then Exit Sub
        .Move .Left + x, .Top, .Width, .Height
    End With
    
    With fraFilter
        .Move .Left, .Top, .Width + x
    End With
    
    With vsfList
        .Move .Left + x, .Top, .Width - x
    End With
    
    With pic水平分割
        .Left = vsfList.Left
        .Width = vsfList.Width
    End With
    
    With vsfDetail
        .Left = vsfList.Left
        .Width = vsfList.Width
    End With
End Sub

Private Sub pic水平分割_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With pic水平分割
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > Me.ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With vsfList
        .Height = pic水平分割.Top - .Top
    End With
    
    With vsfDetail
        .Top = pic水平分割.Top + pic水平分割.Height
        .Height = Me.ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
    End With
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Add" '新增
            Call mnuEditAdd_Click
        Case "Modify" '修改
            Call mnuEditUpdate_Click
        Case "Delete" '删除
            Call mnuEditDelete_Click
        Case "Verify" '处理
            Call mnuEditVerify_Click
        Case "Help" '帮助
            Call mnuHelpTitle_Click
        Case "Exit" '退出
            Call mnufileexit_Click
        Case "Preview" '预览
            Call mnuFilePreView_Click
        Case "Print" '打印
            Call mnuFilePrint_Click
    End Select
End Sub

Private Sub mnuFilePrint_Click()
    '打印
    Dim lngCurRow As Long
    
    If ActiveControl Is vsfList Then
        lngCurRow = vsfList.Row
        vsfList.Redraw = flexRDNone
        subPrint 1
        vsfList.Row = lngCurRow
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    Else
        lngCurRow = vsfDetail.Row
        vsfDetail.Redraw = flexRDNone
        subPrint 1
        vsfDetail.Row = lngCurRow
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub mnuFilePreView_Click()
    '打印预览
    Dim lngCurRow As Long
    
    If ActiveControl Is vsfList Then
        lngCurRow = vsfList.Row
        vsfList.Redraw = flexRDNone
        subPrint 2
        vsfList.Row = lngCurRow
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    Else
        lngCurRow = vsfDetail.Row
        vsfDetail.Redraw = flexRDNone
        subPrint 2
        vsfDetail.Row = lngCurRow
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "材料质量管理"
    
    objRow.Add "部门：" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "打印人:" & UserInfo.用户名
    objRow.Add "打印日期:" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    If ActiveControl Is vsfList Then
        Set objPrint.Body = vsfList
    Else
        Set objPrint.Body = vsfDetail
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

Private Sub DeleteStuff(ByVal lng质量id As Long, ByVal strNo As String)
    '删除单据
    If MsgBox("将删除当前选中单据，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        gstrSQL = "Zl_材料质量主表_Delete(" & lng质量id & ")"
        Call zldatabase.ExecuteProcedure(gstrSQL, "材料质量管理")
        
        gstrSQL = "zl_材料其他出库_Delete('" & strNo & "')"
        Call zldatabase.ExecuteProcedure(gstrSQL, "材料质量管理")
        MsgBox "删除单据成功！", vbInformation, gstrSysName
    End If
End Sub

Private Sub CheckDepend()
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    
    '获取可操作的库房性质编码
    strStock = "VKW"
    
    '检查当前人员所属科室是否为“卫材库”、“制剂室”、“发料部门”
    gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
            & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
            & "Where (a.站点 = [3] Or a.站点 is Null) And c.工作性质 = b.名称 " _
            & "  AND Instr([2],b.编码,1) > 0 " _
            & "  AND a.id = c.部门id " _
            & "  AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" _
            & IIf(InStr(1, gstrPrivs, ";所有库房;") > 0, "", " and a.id in (Select 部门id from 部门人员 where 人员id =[1])")
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "材料质量管理", UserInfo.Id, strStock, gstrNodeNo)
    
    If rsDepend.EOF Then
        MsgBox "没有设置卫材库性质的部门或不具备相关的权限,请查看部门管理或找系统管事员授权！", vbInformation, gstrSysName
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Sub
    End If
    
    '装入库房数据
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!名称
            .ItemData(.NewIndex) = rsDepend!Id
            If rsDepend!Id = UserInfo.部门ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsDepend.Close
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtProvider_Change()
    If Trim(txtProvider.Text) = "" Then
        txtProvider.Tag = 0
    End If
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    Dim strProviderText As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    vRect = zlControl.GetControlRect(txtProvider.hwnd) '获取位置
    dblLeft = vRect.Left
    dblTop = vRect.Top
    
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(Trim(.Text))
        gstrSQL = "Select id,编码,名称,简码 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
                  "  And 末级=1 And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (简码 like [1] Or 编码 like [1] or 名称 like [1] )"

        Set rsProvider = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "供药单位", False, "", "", False, False, _
                        True, dblLeft, dblTop, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        If rsProvider Is Nothing Then
            MsgBox "未找到该供应商“" & Trim(.Text) & "”，请重新输入！", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        Else
            txtProvider.Text = rsProvider!名称
            txtProvider.Tag = rsProvider!Id
        End If
    End With
    
    txtStuff.SetFocus
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetList()
    '加载汇总数据
    Dim rsTemp As ADODB.Recordset
    Dim datTemp As Date
    Dim strCheckBegin As String
    Dim strCheckEnd As String
    Dim strVerifyBegin As String
    Dim strVerifyEnd As String
    Dim datCheckBegin As Date
    Dim datCheckEnd As Date
    Dim datVerifyBegin As Date
    Dim datVerifyEnd As Date
    Dim lng库房ID As Long
    Dim lng供药单位ID As Long
    Dim lng材料ID As Long
    Dim str登记人 As String
    Dim str处理人 As String
    
    On Error GoTo ErrHandle
    
    '小数格式化
    mintUnit = IIf(opt单位(0).Value = True, 0, 1)
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(2, g_售价)
    End With
    
    vsfList.Rows = 1
    '库房id
    gstrSQL = " and b.库房id=[1]"
    lng库房ID = cboStock.ItemData(cboStock.ListIndex)
    
    datTemp = zldatabase.Currentdate
    '登记日期
    If chkCheck.Value = 1 Then
        Select Case cboCheckDate.Text
        Case "一周内"
            strCheckBegin = Format(DateAdd("D", -7, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "一月内"
            strCheckBegin = Format(DateAdd("M", -1, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "三月内"
            strCheckBegin = Format(DateAdd("M", -3, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "半年内"
            strCheckBegin = Format(DateAdd("M", -6, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "一年内"
            strCheckBegin = Format(DateAdd("YYYY", -1, datTemp), "yyyy-mm-dd")
            strCheckEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        Case "自定义"
            strCheckBegin = Format(dtpCheckBegin.Value, "yyyy-mm-dd")
            strCheckEnd = Format(dtpCheckEnd.Value, "yyyy-mm-dd") & " 23:59:59"
            datCheckBegin = CDate(strCheckBegin)
            datCheckEnd = CDate(strCheckEnd)
        End Select
    End If
    
    '处理日期
    If chkVerify.Value = 1 Then
        Select Case cboVerifyDate.Text
        Case "一周内"
            strVerifyBegin = Format(DateAdd("D", -7, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "一月内"
            strVerifyBegin = Format(DateAdd("M", -1, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "三月内"
            strVerifyBegin = Format(DateAdd("M", -3, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "半年内"
            strVerifyBegin = Format(DateAdd("M", -6, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "一年内"
            strVerifyBegin = Format(DateAdd("YYYY", -1, datTemp), "yyyy-mm-dd")
            strVerifyEnd = Format(datTemp, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        Case "自定义"
            
            strVerifyBegin = Format(dtpVerifyBegin.Value, "yyyy-mm-dd")
            strVerifyEnd = Format(dtpVerifyEnd.Value, "yyyy-mm-dd") & " 23:59:59"
            datVerifyBegin = CDate(strVerifyBegin)
            datVerifyEnd = CDate(strVerifyEnd)
        End Select
    End If
    
    If chkCheck.Value = 1 And chkVerify.Value = 1 Then gstrSQL = gstrSQL & " and (a.登记日期 between [2] and [3] or a.处理日期 between [4] and [5]) "
    
    If chkCheck.Value = 1 And chkVerify.Value = 0 Then gstrSQL = gstrSQL & " and a.登记日期 between [2] and [3] and a.处理日期 is null "
    
    If chkCheck.Value = 0 And chkVerify.Value = 1 Then gstrSQL = gstrSQL & " and a.处理日期 between [4] and [5] "
    
    '供药单位id
    If Val(txtProvider.Tag) <> 0 Then
        gstrSQL = gstrSQL & " and b.供药单位id=[6]"
    End If
    lng供药单位ID = Val(txtProvider.Tag)
    
    '材料id
    If Val(txtStuff.Tag) <> 0 Then
        gstrSQL = gstrSQL & " and b.材料id=[7]"
    End If
    lng材料ID = Val(txtStuff.Tag)
    
    '登记人
    If Trim(txtCheck.Text) <> "" Then
        gstrSQL = gstrSQL & " and a.登记人 like [8]"
    End If
    str登记人 = Trim(txtCheck.Text)
    
    '处理人
    If Trim(txtVerify.Text) <> "" Then
        gstrSQL = gstrSQL & " and a.处理人 like [9]"
    End If
    str处理人 = Trim(txtVerify.Text)
    
    gstrSQL = "Select Distinct a.Id, a.No, a.登记人, a.登记日期, a.处理人, a.处理日期, a.备注 " & vbNewLine & _
            "From 材料质量主表 A, 材料质量记录 B " & vbNewLine & _
            "Where a.Id = b.质量id " & gstrSQL & " Order by a.No Desc"
            
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
    lng库房ID, _
    datCheckBegin, _
    datCheckEnd, _
    datVerifyBegin, _
    datVerifyEnd, _
    lng供药单位ID, _
    lng材料ID, _
    str登记人, _
    str处理人)
    
    With vsfList
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTemp!Id
            .TextMatrix(.Rows - 1, .ColIndex("NO")) = rsTemp!NO
            .TextMatrix(.Rows - 1, .ColIndex("登记人")) = IIf(IsNull(rsTemp!登记人), "", rsTemp!登记人)
            .TextMatrix(.Rows - 1, .ColIndex("登记日期")) = IIf(IsNull(rsTemp!登记日期), "", Format(rsTemp!登记日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.Rows - 1, .ColIndex("处理人")) = IIf(IsNull(rsTemp!处理人), "", rsTemp!处理人)
            .TextMatrix(.Rows - 1, .ColIndex("处理日期")) = IIf(IsNull(rsTemp!处理日期), "", Format(rsTemp!处理日期, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.Rows - 1, .ColIndex("备注")) = IIf(IsNull(rsTemp!备注), "", rsTemp!备注)
            rsTemp.MoveNext
        Loop
        
        If .Rows > 1 Then
            .Row = 1
            .SetFocus
            
            Call vsfList_EnterCell
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfList_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfList
        If Val(.TextMatrix(NewRow, .ColIndex("id"))) = 0 Then
            tlbTool.Buttons("Modify").Enabled = False
            tlbTool.Buttons("Delete").Enabled = False
            tlbTool.Buttons("Verify").Enabled = False
            mnuEditDelete.Enabled = False
            mnuEditUpdate.Enabled = False
            mnuEditVerify.Enabled = False
        Else
            If .TextMatrix(NewRow, .ColIndex("处理人")) <> "" Then
                tlbTool.Buttons("Delete").Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
                mnuEditDelete.Enabled = False
                mnuEditUpdate.Enabled = False
                mnuEditVerify.Enabled = False
            Else
                tlbTool.Buttons("Delete").Enabled = True
                tlbTool.Buttons("Verify").Enabled = True
                tlbTool.Buttons("Modify").Enabled = True
                mnuEditDelete.Enabled = True
                mnuEditUpdate.Enabled = True
                mnuEditVerify.Enabled = True
            End If
        End If
    End With
End Sub

Private Sub vsfList_DblClick()
    '查阅
    Dim lng质量id As Long

    With vsfList
        If .Row > 0 Then
            If .TextMatrix(.Row, .ColIndex("id")) <> "" Then
                lng质量id = Val(.TextMatrix(.Row, .ColIndex("id")))
            Else
                lng质量id = 0
            End If
        Else
            lng质量id = 0
        End If
    End With

    If lng质量id <> 0 Then
        frmStuffQualityCard.ShowMe 4, Me, cboStock.ItemData(cboStock.ListIndex), lng质量id, mintUnit
    End If
End Sub

Private Sub vsfList_EnterCell()
    Dim lng质量id As Long
    Dim rsTemp As ADODB.Recordset
    Dim str包装系数 As String
    Dim dblTemp As Double
    
    On Error GoTo ErrHandle
    With vsfDetail
        .Rows = 1
        
        If Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("id"))) <> 0 Then
            lng质量id = Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("id")))
            
            Select Case mintUnit
                Case 0
                    str包装系数 = "1"
                Case Else
                    str包装系数 = "c.换算系数"
            End Select
            
            gstrSQL = "Select b.编码, b.名称, b.规格, d.名称 As 供应商, a.批号, a.批次, a.产地, a.毁损原因, a.解决办法, " & IIf(mintUnit = 0, " b.计算单位", " c.包装单位") & " as 单位," & _
                    " a.毁损数量, a.成本价, a.零售价, c.换算系数 " & _
                    " From 材料质量记录 A, 收费项目目录 B, 材料特性 C, 供应商 D" & _
                    " Where a.材料id = b.Id And b.Id = c.材料id And a.供药单位id = d.Id(+) And a.质量id = [1]"
                    
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng质量id)
            
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                
                .TextMatrix(.Rows - 1, .ColIndex("材料信息")) = "[" & rsTemp!编码 & "]" & rsTemp!名称
                .TextMatrix(.Rows - 1, .ColIndex("规格")) = IIf(IsNull(rsTemp!规格), "", rsTemp!规格)
                .TextMatrix(.Rows - 1, .ColIndex("批次")) = IIf(IsNull(rsTemp!批次), "", rsTemp!批次)
                .TextMatrix(.Rows - 1, .ColIndex("供应商")) = IIf(IsNull(rsTemp!供应商), "", rsTemp!供应商)
                .TextMatrix(.Rows - 1, .ColIndex("产地")) = IIf(IsNull(rsTemp!产地), "", rsTemp!产地)
                .TextMatrix(.Rows - 1, .ColIndex("单位")) = IIf(IsNull(rsTemp!单位), "", rsTemp!单位)
                .TextMatrix(.Rows - 1, .ColIndex("毁损原因")) = IIf(IsNull(rsTemp!毁损原因), "", rsTemp!毁损原因)
                .TextMatrix(.Rows - 1, .ColIndex("解决办法")) = IIf(IsNull(rsTemp!解决办法), "", rsTemp!解决办法)
                
                str包装系数 = IIf(mintUnit = 0, 1, rsTemp!换算系数)
            
                If IsNull(rsTemp!毁损数量) = False Then dblTemp = Val(rsTemp!毁损数量) / Val(str包装系数)
                .TextMatrix(.Rows - 1, .ColIndex("毁损数量")) = Format(dblTemp, mFMT.FM_数量)
                
                If IsNull(rsTemp!成本价) = False Then dblTemp = Val(rsTemp!成本价) * Val(str包装系数)
                .TextMatrix(.Rows - 1, .ColIndex("成本价")) = Format(dblTemp, mFMT.FM_成本价)
                
                If IsNull(rsTemp!零售价) = False Then dblTemp = Val(rsTemp!零售价) * Val(str包装系数)
                .TextMatrix(.Rows - 1, .ColIndex("零售价")) = Format(dblTemp, mFMT.FM_零售价)
                
                dblTemp = Val(rsTemp!成本价) * Val(rsTemp!毁损数量)
                .TextMatrix(.Rows - 1, .ColIndex("成本金额")) = Format(dblTemp, mFMT.FM_金额)
                
                dblTemp = Val(rsTemp!零售价) * Val(rsTemp!毁损数量)
                .TextMatrix(.Rows - 1, .ColIndex("售价金额")) = Format(dblTemp, mFMT.FM_金额)
                
                rsTemp.MoveNext
            Loop
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
