VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "病人收费管理"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "frmManageCharge.frx":0000
   KeyPreview      =   -1  'True
   Picture         =   "frmManageCharge.frx":08CA
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picExtendInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   2355
      ScaleHeight     =   2505
      ScaleWidth      =   1755
      TabIndex        =   21
      Top             =   2445
      Width           =   1755
      Begin VSFlex8Ctl.VSFlexGrid vsfExtendInfo 
         Height          =   1275
         Left            =   240
         TabIndex        =   22
         Top             =   630
         Width           =   1290
         _cx             =   2275
         _cy             =   2249
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
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
   End
   Begin VB.PictureBox picTemp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   5025
      ScaleHeight     =   570
      ScaleWidth      =   810
      TabIndex        =   20
      Top             =   5235
      Width           =   810
   End
   Begin XtremeSuiteControls.TabControl tbSub 
      Height          =   1815
      Left            =   5580
      TabIndex        =   14
      Top             =   4035
      Width           =   4080
      _Version        =   589884
      _ExtentX        =   7197
      _ExtentY        =   3201
      _StockProps     =   64
   End
   Begin VB.PictureBox picSubInvoice 
      BorderStyle     =   0  'None
      Height          =   2220
      Left            =   6510
      ScaleHeight     =   2220
      ScaleWidth      =   2370
      TabIndex        =   17
      Top             =   2985
      Width           =   2370
      Begin VSFlex8Ctl.VSFlexGrid vsSubInvoice 
         Height          =   1305
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1635
         _cx             =   2884
         _cy             =   2302
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   1
         HighLight       =   2
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
   End
   Begin VB.PictureBox picSubBalance 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   4335
      ScaleHeight     =   2505
      ScaleWidth      =   1755
      TabIndex        =   15
      Top             =   2520
      Width           =   1755
      Begin VSFlex8Ctl.VSFlexGrid vsSubBalance 
         Height          =   1065
         Left            =   510
         TabIndex        =   16
         Top             =   480
         Width           =   1710
         _cx             =   3016
         _cy             =   1879
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
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
   End
   Begin VSFlex8Ctl.VSFlexGrid mshDetail 
      Height          =   1515
      Left            =   45
      TabIndex        =   13
      Top             =   4200
      Width           =   5325
      _cx             =   9393
      _cy             =   2672
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmManageCharge.frx":0A4C
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
      OutlineBar      =   4
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
   Begin MSComctlLib.ImageList imgGray 
      Left            =   8790
      Top             =   0
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
            Picture         =   "frmManageCharge.frx":0B3D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":0D57
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":0F71
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":118B
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":13A5
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":1B1F
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":1D39
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":1F53
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":216D
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":2387
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":25A1
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":27BB
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":C152
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   7845
      Top             =   60
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
            Picture         =   "frmManageCharge.frx":C84C
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":CA66
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":CC80
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":CE9A
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":D0B4
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":D82E
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":DA48
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":DC62
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":DE7C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":E096
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":E2B0
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":E4CA
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageCharge.frx":EBC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCons 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6000
      ScaleHeight     =   300
      ScaleWidth      =   7260
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   7260
      Begin VB.ComboBox cboDate 
         Height          =   300
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   0
         Width           =   1620
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   15
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   160038915
         CurrentDate     =   40777
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   285
         Left            =   4125
         TabIndex        =   9
         Top             =   15
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   160038915
         CurrentDate     =   40777
      End
      Begin VB.Label lblDateShow 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   2535
         TabIndex        =   12
         Top             =   45
         Width           =   90
      End
      Begin VB.Label lblSplit 
         Caption         =   "～"
         Height          =   210
         Left            =   3870
         TabIndex        =   10
         Top             =   45
         Width           =   330
      End
      Begin VB.Label lbl缺省 
         AutoSize        =   -1  'True
         Caption         =   "缺省显示"
         Height          =   180
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   5535
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1695
      ScaleWidth      =   45
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4170
      Width           =   45
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   15
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9675
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4140
      Width           =   9675
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9675
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "收费"
               Key             =   "Charge"
               Description     =   "收费"
               Object.ToolTipText     =   "进入收费窗口"
               Object.Tag             =   "收费"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退费"
               Key             =   "Del"
               Description     =   "退费"
               Object.ToolTipText     =   "对当前选中单据退费"
               Object.Tag             =   "退费"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "作废"
               Key             =   "Cancel"
               Object.ToolTipText     =   "作废异常单据"
               Object.Tag             =   "作废"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查阅"
               Key             =   "View"
               Description     =   "查阅"
               Object.ToolTipText     =   "查阅当前单据的内容"
               Object.Tag             =   "查阅"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "按设置条件重新筛选记录"
               Object.Tag             =   "过滤"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "定位"
               Key             =   "Go"
               Description     =   "定位"
               Object.ToolTipText     =   "定位到满足条件的记录上"
               Object.Tag             =   "定位"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "轧帐"
               Key             =   "轧帐"
               Object.ToolTipText     =   "收费轧帐"
               Object.Tag             =   "轧帐"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "扩展"
               Key             =   "Extra"
               Object.ToolTipText     =   "外挂扩展功能"
               Object.Tag             =   "扩展"
               ImageIndex      =   13
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExtraItem"
                     Object.Tag             =   "功能"
                     Text            =   "功能"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   5844
      Width           =   9672
      _ExtentX        =   17066
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageCharge.frx":F2BE
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8229
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
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmManageCharge.frx":FB52
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshFact 
      Height          =   1815
      Left            =   5580
      TabIndex        =   0
      Top             =   4035
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmManageCharge.frx":FCE4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   405
      Left            =   15
      TabIndex        =   19
      Top             =   795
      Width           =   9615
      _Version        =   589884
      _ExtentX        =   16960
      _ExtentY        =   714
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid mshList 
      Height          =   1515
      Left            =   120
      TabIndex        =   23
      Top             =   1500
      Width           =   5325
      _cx             =   9393
      _cy             =   2672
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmManageCharge.frx":FFFE
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
      OutlineBar      =   4
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "现金点钞(&E)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "收费轧帐(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Insure 
         Caption         =   "保险类别(&I)"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEdit_Charge 
         Caption         =   "门诊收费(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Simple 
         Caption         =   "简单收费(&S)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditReCharge 
         Caption         =   "重新收费(&R)"
      End
      Begin VB.Menu mnuEditCancelBill 
         Caption         =   "作废收费(&Z)"
      End
      Begin VB.Menu mnuEdit_Charge_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "调整时间(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_DelMulti 
         Caption         =   "结算退费(&U)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplitMzToZy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzToZyDel 
         Caption         =   "转住院费用退费(Q)"
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "查阅单据(&V)"
      End
      Begin VB.Menu mnuEdit_View_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "重打收费票据(&R)"
      End
      Begin VB.Menu mnuEditInvoicePrint 
         Caption         =   "按发票号重打票据(&F)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "补打收费票据(&B)"
      End
      Begin VB.Menu mnuEditMakeupPrn 
         Caption         =   "按病人补打票据(&M)"
      End
      Begin VB.Menu mnuEdit_PrintDel 
         Caption         =   "重打退费票据(&D)"
      End
      Begin VB.Menu mnuEdit_PrintList 
         Caption         =   "打印收费清单(&L)"
      End
      Begin VB.Menu mnuEdit_PrintProve 
         Caption         =   "打印收据证明(&O)"
      End
      Begin VB.Menu mnuEdit_Apply_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Apply 
         Caption         =   "退费申请(&P)"
      End
      Begin VB.Menu mnuEdit_UnApply 
         Caption         =   "取消申请(&D)"
      End
      Begin VB.Menu mnuEdit_Audit 
         Caption         =   "退费审核(&T)"
      End
      Begin VB.Menu mnuEdit_RefuseApply 
         Caption         =   "拒绝申请(&R)"
      End
      Begin VB.Menu mnuEdit_UnAudit 
         Caption         =   "取消审核(&D)"
      End
      Begin VB.Menu mnuEditSplitW 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWriteCard 
         Caption         =   "门诊信息写卡(&W)"
      End
      Begin VB.Menu mnuEdit_Extra 
         Caption         =   "扩展"
         Begin VB.Menu mnuEdit_ExtraItem 
            Caption         =   "功能"
            Index           =   0
         End
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
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "定位(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "刷新方式(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后不要刷新数据(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后提示是否刷新(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "操作后自动刷新数据(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
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
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuFeeDetial 
      Caption         =   "费用明细右键菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuFeeDetial_Print 
         Caption         =   "重打收费票据(&R)"
      End
      Begin VB.Menu mnuFeeDetial_Supplemental 
         Caption         =   "补打收费票据(&B)"
      End
      Begin VB.Menu mnuFeeDetial_PrintList 
         Caption         =   "打印收费清单(&L)"
      End
      Begin VB.Menu mnuFeeDetial_PrintProve 
         Caption         =   "打印收据证明(&O)"
      End
   End
End
Attribute VB_Name = "frmManageCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mrsList As ADODB.Recordset  '单据列表
Private mrsDetail As ADODB.Recordset
Private mrsTotal As ADODB.Recordset
Private mrsFact As ADODB.Recordset
Private mrsBalance As ADODB.Recordset
Private mbln立即销帐 As Boolean
Private mblnFirst As Boolean
Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    ChargeKind As String
    PayKind As String
    PayKindName As String
    PatientID As Long '病人ID
    PatientName As String '病人姓名
    PatientIdentity As String '标识号
    NOB As String
    NOE As String
    FactB As String
    FactE As String
    DeptID As Long
    Doctor As String
    Operator As String
    FeeItems As String
    ApplyName As String
    ApplyDateB As Date
    ApplyDateE As Date
    AuditName As String
    AuditDateB As Date
    AuditDateE As Date
    int门诊标志 As Integer  '1-门诊;2-住院;3-门诊和住院 33789
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String, mstrFilter2 As String, mstrInsure As String
Private mbln收费 As Boolean, mbln退费 As Boolean

Private mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNOMoved As Boolean '显明细时记录当前选择的单据是否在在线数据表中,以其它操作时无需再判断
Private mobjInExise As Object
Private mblnNotClick As Boolean
Private mstrWriteCardTypeIDs As String   '当前包含的所有卡类别ID
Private mblnPrinting As Boolean
'-----------------------------------------------------------------------------------
'结算卡相关
Private mstrPassWord As String
Private mstrPrivs_RollingCurtain As String  '收费轧帐管理权限
'消息相关对象变量
Private WithEvents mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private Sub cboDate_Click()
    lblSplit.Visible = cboDate.ListIndex = 5
    dtpStartDate.Visible = cboDate.ListIndex = 5
    dtpEndDate.Visible = cboDate.ListIndex = 5
    lblDateShow.Visible = cboDate.ListIndex <> 5
    If cboDate.Visible = False Then Exit Sub
    Call mnuViewReFlash_Click
End Sub
'-----------------------------------------------------------------------------------
 Private Sub cbr_Resize()
    Form_Resize
End Sub

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
    If Not mblnFirst Then Exit Sub
    mblnFirst = False
    Call CheckErrBill
End Sub

Private Sub mnuEdit_Adjust_Click()
    Dim strNo As String
    
    strNo = mshDetail.TextMatrix(mshDetail.Row, mshDetail.ColIndex("单据号"))

    If strNo = "" Then
        MsgBox "当前没有单据可以调整！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If

    '已经退过费(部分)的单据不允许调整
    If mshList.TextMatrix(mshList.Row, GetColNum("医保")) <> "√" Then
        If InStr(mstrPrivs, "允许非医保病人") = 0 Then
            MsgBox "你没有权限对非医保病人的单据进行调整时间操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If BillExistDelete(strNo, 1) Then
        MsgBox "该单据包含已退费内容,不允许调整！", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear

    If frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, EM_ED_调整, , , , , , mobjMsgModule, strNo) = True Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mblnNotClick = True
                mnuViewReFlash_Click
                mblnNotClick = False
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
    End If
End Sub

Private Sub mnuEdit_DelMulti_Click()
    Dim strNo As String, blnDel As Boolean
    With mshList
        '重大改进后，传入的是结帐ID
        strNo = .TextMatrix(.Row, GetColNum("结算序号"))
    End With
    If Val(strNo) = 0 Then Exit Sub
    If CheckBillExistReplenishData(0, Val(strNo)) = True Then
        MsgBox "选择的退费记录进行了医保补充结算，不允许进行退费操作！", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(strNo) < 0 Then
        blnDel = frmClinicDelAndView.ShowMe(Me, EM_MULTI_退费, mstrPrivs, Val(strNo))
    Else
        Call DelOldBill
        Exit Sub
    End If
    If blnDel Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mblnNotClick = True
                mnuViewReFlash_Click
                mblnNotClick = False
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
    End If
End Sub

Private Sub mnuEdit_PrintDel_Click()
    Call PrintDelBill
End Sub

Private Sub PrintDelBill()
'功能：当前退款记录重新打印一张票据
    Dim strNo As String, lngBalance As Long, blnMediCare As Boolean
    Dim intInsure As Integer, blnVirtualPrint As Boolean
    Dim lng结帐ID As Long, lng病人ID As Long, blnDel As Boolean
    Dim strUseType  As String, lngShareUseID As Long, intInvoiceFormat As Integer
    
    Err = 0: On Error GoTo errHandler
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    If strNo = "" Then
        MsgBox "当前没有单据可以重打退费票据！", vbInformation, gstrSysName
        Exit Sub
    End If
    lngBalance = Val(mshList.TextMatrix(mshList.Row, GetColNum("结算序号")))

    If CheckBillExistReplenishData(0, lngBalance) = True Then
        MsgBox "选择的记录进行了医保补充结算，不允许进行重打退费票据操作！", vbInformation, gstrSysName
        Exit Sub
    End If

    blnMediCare = mshList.TextMatrix(mshList.Row, GetColNum("医保")) = "√"
    blnDel = mshList.TextMatrix(mshList.Row, GetColNum("符号")) = "3"   '记录状态为2的，目前是禁用了打印菜单项的
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    If Not BillOperCheck(2, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("收费时间"))), "退费单据重打", strNo, , 1) Then Exit Sub
    
    If blnMediCare Then
        intInsure = ChargeExistInsure(strNo, lng病人ID, lng结帐ID, , blnDel)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
        End If
    End If
    
    lng病人ID = zlGet病人ID(strNo)
    strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
    lngShareUseID = zl_GetInvoiceShareID(mlngModul, strUseType)
    intInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, strUseType, , , True)
    
    '打印退费票据(红票)
    If PrintDelCharge(lngBalance, Me, 0, , , intInvoiceFormat, blnVirtualPrint, blnDel, lngShareUseID, strUseType) Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容，要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub mnuEdit_Simple_Click()
    On Error Resume Next
    Err.Clear
    frmSimpleCharge.mlngModul = mlngModul
    frmSimpleCharge.mstrPrivs = mstrPrivs
    frmSimpleCharge.mbytInState = 0
    frmSimpleCharge.Show 1, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mblnNotClick = True
                mnuViewReFlash_Click
                mblnNotClick = False
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
    End If
End Sub

Private Sub mnuEdit_PrintList_Click()
    Dim strNo As String, strNos As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    If strNo = "" Then
        MsgBox "当前没有单据可以打印清单！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    If Me.ActiveControl Is mshDetail Then
        '89761
        If mshDetail.IsSubtotal(mshDetail.Row) Then
            strNo = mshDetail.TextMatrix(mshDetail.Row + 1, 0)
        Else
            strNo = mshDetail.TextMatrix(mshDetail.Row, 0)
        End If
        strNos = "'" & strNo & "'"
    Else
        strNos = GetMultiNOs(strNo, , , True)  '可能是多单据收费中的一张
    End If
    
    If glngSys Like "8??" Then
        If ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me, "NO=" & strNos, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
        End If
    Else
        If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
        End If
    End If
End Sub

Private Sub mnuEdit_PrintProve_Click()
    Dim strNo As String, strNos As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    If strNo = "" Then
        MsgBox "当前没有单据可以打印证明！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    If Me.ActiveControl Is mshDetail Then
        '89761
        If mshDetail.IsSubtotal(mshDetail.Row) Then
            strNo = mshDetail.TextMatrix(mshDetail.Row + 1, 0)
        Else
            strNo = mshDetail.TextMatrix(mshDetail.Row, 0)
        End If
        strNos = "'" & strNo & "'"
    Else
        strNos = GetMultiNOs(strNo, , , True) '可能是多单据收费中的一张
    End If
    
    If glngSys Like "8??" Then
        If ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_2", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL8_BILL_1121_2", Me, "NO=" & strNos, 2)
        End If
    Else
        If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_2", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_2", Me, "NO=" & strNos, 2)
        End If
    End If
End Sub

Private Function ReChargeToErrBillBefore(ByVal strNo As String, ByVal lng结算序号 As Long, Optional blnDel As Boolean = False, _
    Optional bln退费异常 As Boolean = False, Optional ByVal strDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:10.34之前重新收取异常的单据费用
    '入参:
    '   blnDel True-作废单据,False-重新收费
    '   bln退费异常 是否退费异常单据
    '   strDate 收费时间
    '返回:收取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 15:41:08
    '说明：若blnDel=True And bln退费异常=True表示作废单据时产生的异常单据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivsReplenish As String
    
    On Error GoTo errHandle
    If bln退费异常 = False Then
        If zlIsCheckExiseSingularity(lng结算序号) Then
            MsgBox "该异常单据已经被作废，因此，不能再" & IIf(blnDel, "进行作废", "重新收费") & "，请刷新费用列表！", vbInformation, gstrSysName
            Exit Function
        End If
        If Not zlIsCheckExistErrBill(lng结算序号) Then
            MsgBox "该异常单据已经被重新收费，因此，不能再" & IIf(blnDel, "进行作废", "重新收费") & "，请刷新费用列表！", vbInformation, gstrSysName
            Exit Function
        End If
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInState = IIf(blnDel, 5, 4)
        frmCharge.mstrInNO = strNo
        frmCharge.mbln退费异常 = False
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        If blnDel Then
            '对作废收费异常记录产生的异常进行处理
            frmCharge.mlngModul = mlngModul
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 5
            frmCharge.mstrInNO = strNo
            frmCharge.mbln退费异常 = True
            Set frmCharge.mobjMsgModule = mobjMsgModule
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            If CheckBillExistReplenishData(0, lng结算序号) Then
                strPrivsReplenish = ";" & GetPrivFunc(glngSys, 1124) & ";"
                If InStr(strPrivsReplenish, ";结算退费;") > 0 Then
                    If MsgBox("选择的记录进行了医保补充结算且为异常补充结算退费记录，是否针对该记录进行再次结算退费？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        gblnOK = frmReplenishTheBalanceDel.zlShowMe(Me, 1124, strPrivsReplenish, EM_RBDTY_异常重退, Val(strNo), False, 0, False, strDate)
                    Else
                        Exit Function
                    End If
                Else
                    MsgBox "选择的记录进行了医保补充结算且为异常补充结算退费记录，你不具备操作该记录的权限，不允许进行退费操作！", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                gblnOK = frmMultiBills.ShowMe(Me, 2, mstrPrivs, strNo, strDate)
            End If
        End If
    End If
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mblnNotClick = True
                mnuViewReFlash_Click
                mblnNotClick = False
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If

        ReChargeToErrBillBefore = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mnuEditCancelBill_Click()
    Dim blnDel As Boolean, lng结算序号 As Long
    Dim strDelTime As String, bln退费异常 As Boolean
    Dim strNo As String
    
    '作废
    If tbPage.Selected.Index <> 2 Then Exit Sub
    With mshList
        strNo = .TextMatrix(.Row, GetColNum("首张单据"))
        lng结算序号 = Val(.TextMatrix(.Row, GetColNum("结算序号")))
        strDelTime = .TextMatrix(.Row, GetColNum("收费时间"))
    End With
    If lng结算序号 = 0 Then
        MsgBox "不存在需要作废的异常单据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If lng结算序号 < 0 Then
        Call ReChargeToErrBill(lng结算序号, True, False, strDelTime)
    Else
        'V10.34.0版本以前数据
        Call ReChargeToErrBillBefore(strNo, lng结算序号, True, False, strDelTime)
    End If
End Sub

Private Function IsCancelFee(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:是否作废异常单
    '编制:刘兴洪
    '日期:2012-03-01 01:04:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSql = "Select 1 From 门诊费用记录 where 记录性质=1 and NO=[1] and 记录状态=3 And RowNum=1 And nvl(费用状态,0)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    IsCancelFee = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mnuEditInvoicePrint_Click()
    '按发票号重打票据
    If frmFromInvoiceToPrint.zlRePrintBill(Me, mlngModul, mstrPrivs, 0) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
    End If
End Sub

Private Sub mnuEditMakeupPrn_Click()
    If frmMakeupPrintBill.zlRePrintBill(gfrmMain, mlngModul, mstrPrivs, 0) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
    End If
End Sub

Private Sub mnuEditMzToZyDel_Click()
    '功能:门诊转住院退费
    '问题:36076
    If InStr(1, mstrPrivs, ";转住院费用退费;") = 0 Or mbln立即销帐 Then Exit Sub
    
    If mobjInExise Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjInExise = CreateObject("zl9InExse.clsInExse")
        If Err <> 0 Then
            MsgBox "注意:" & "    住院费用部件创建失败,不能进行退费,请与系统管员联系!", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        Err = 0
    End If
    If mobjInExise Is Nothing Then Exit Sub
'    CallMzFeeTOZyFeeDel(ByVal frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, ByVal lngSys As Long, _
'    ByVal lngModule As Long, ByVal strPrivs As String,ByVal int性质 As Integer, optional lng病人ID as long =0) As Boolean
    If mobjInExise.CallMzFeeTOZyFeeDel(Me, gcnOracle, gstrDBUser, glngSys, mlngModul, mstrPrivs, 1) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
    End If
End Sub

Private Sub mnuEditReCharge_Click()
    Dim strNo As String, lng结算序号 As Long
    Dim bln退费异常 As Boolean, blnDel As Boolean
    Dim strDelTime As String
    
    Err = 0: On Error GoTo errHandler
    If tbPage.Selected.Index <> 2 And tbPage.Selected.Index <> 3 Then Exit Sub
    
    With mshList
        bln退费异常 = tbPage.Selected.Index = 3
        strNo = .TextMatrix(.Row, GetColNum("首张单据"))
        lng结算序号 = Val(.TextMatrix(.Row, GetColNum("结算序号")))
        strDelTime = .TextMatrix(.Row, GetColNum("收费时间"))
    End With
    If strNo = "" Then
        MsgBox "不存在重新收费或退费的异常单据！", vbInformation, gstrSysName

        Exit Sub
    End If
    
    If bln退费异常 Then
        '判断指定的单据是否异常的收费作废操作异常
        blnDel = zlIsErrChargeCancel(strNo)
    End If
    
    If lng结算序号 < 0 Then
        Call ReChargeToErrBill(lng结算序号, blnDel, bln退费异常, strDelTime)
    Else
        'V10.34.0版本以前数据
        Call ReChargeToErrBillBefore(strNo, lng结算序号, blnDel, bln退费异常, strDelTime)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditWriteCard_Click()
    Dim lngCardTypeID As Long, strExpend As String, lng病人ID As Long
    Dim lng结算序号 As Long, strNo As String, lng记录状态 As Long
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    '功能:将门诊信息写入卡中
    '问题:56615
    If InStr(mstrPrivs, ";门诊信息写卡;") = 0 Or mstrWriteCardTypeIDs = "" Then Exit Sub
    If gbln退费申请模式 And tbPage.Selected.Index = 1 Then Exit Sub
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    If strNo = "" Then
        MsgBox "当前没有单据可以重新写卡！", vbInformation, gstrSysName
        Exit Sub
    End If
    '是否查看退费单据
    lng记录状态 = Val(mshList.TextMatrix(mshList.Row, GetColNum("符号")))
    
    '1.费用未完全执行(执行状态=0,2)
    strSql = "Select  A.病人ID,B.结算序号" & _
        " From 门诊费用记录 A,病人预交记录 B " & vbNewLine & _
        " Where A.结帐ID=B.结帐ID and  Nvl(A.附加标志,0)<>9 And A.NO=[1] And A.记录性质=1  " & _
        "       And A.记录状态 =[2]  and Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, lng记录状态)
    If rsTemp.EOF Then Exit Sub
    
    lng病人ID = Val(NVL(rsTemp!病人ID))
    lng结算序号 = Val(NVL(rsTemp!结算序号))
    If lng病人ID = 0 Or lng结算序号 = 0 Then Exit Sub
    
    If InStr(1, mstrWriteCardTypeIDs, ",") = 0 Then lngCardTypeID = Val(mstrWriteCardTypeIDs)
    Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, lng病人ID, lng结算序号, strExpend)
End Sub
 

Private Sub mnuFeeDetial_Print_Click()
    Call mnuEdit_Print_Click
End Sub

Private Sub mnuFeeDetial_PrintList_Click()
    Call mnuEdit_PrintList_Click
End Sub

Private Sub mnuFeeDetial_PrintProve_Click()
    Call mnuEdit_PrintProve_Click
End Sub

Private Sub mnuFeeDetial_Supplemental_Click()
    Call mnuEdit_Print_Supplemental_Click
End Sub

Private Sub mnuFile_Insure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFileLocalSet_Click()
    Dim blnUnit As Boolean, blnErr As Boolean
    Dim intFrom As Integer
    Dim bytInvoice As Byte
    Dim blnErrPage As Boolean   '异常单据页面
    
    intFrom = gint病人来源
    blnUnit = gbln药房单位
    blnErr = gblnShowErr
    bytInvoice = gTy_Module_Para.byt票据分配规则
        
    With frmSetExpence
        .mlngModul = mlngModul
        .mstrPrivs = mstrPrivs
        .mbytInFun = 0
        .mblnSetDrugStore = False
        .Show 1, Me
    End With
    
 
    '更改了相关参数,重新刷新
    If gbln药房单位 <> blnUnit Or gblnShowErr <> blnErr Or intFrom <> gint病人来源 Or bytInvoice <> gTy_Module_Para.byt票据分配规则 Then
        '屏菜单:按发票号重打票据
        If Not tbPage.Selected Is Nothing Then
            blnErrPage = tbPage.Selected.Index = 2
        Else
            blnErrPage = False
        End If
        mnuEditInvoicePrint.Visible = gTy_Module_Para.byt票据分配规则 <> 0 And Not (InStr(mstrPrivs, ";重打票据;") = 0 Or InStr(mstrPrivs, "收据打印") = 0) And Not blnErrPage
        
        frmChargeGo.lbl标识号.Caption = "标识号"
        If gbln退费申请模式 And tbPage.Selected.Index = 1 Then
            frmChargeFilter.lbl标识号.Caption = "标识号"
        ElseIf gint病人来源 = 1 Then
            frmChargeFilter.opt病人(0).Value = True
        ElseIf gint病人来源 = 2 Then
            frmChargeFilter.opt病人(1).Value = True
        End If
        ShowBills IIf(gbln退费申请模式 And tbPage.Selected.Index = 1, mstrFilter2, mstrFilter)
    End If
End Sub

Private Sub mnuFileMoneyEnum_Click()
    Call frmMoneyEnum.ShowMe(Me)
End Sub

Private Sub mnuFileRollingCurtain_Click()
    Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String, strTmp As String, strColValue As String
    Dim strNos As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    If strNo <> "" Then
        With mshList
            If gbln退费申请模式 And tbPage.Selected.Index = 1 Then
                Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                        "NO=" & .TextMatrix(.Row, GetColNum("首张单据")))
            Else
                If Me.ActiveControl Is mshDetail Then
                    '89761
                    If mshDetail.IsSubtotal(mshDetail.Row) Then
                        strNo = mshDetail.TextMatrix(mshDetail.Row + 1, 0)
                    Else
                        strNo = mshDetail.TextMatrix(mshDetail.Row, 0)
                    End If
                    strNos = "'" & strNo & "'"
                Else
                    strNos = GetMultiNOs(strNo, , , True)  '可能是多单据收费中的一张
                End If
                
                strColValue = .TextMatrix(.Row, GetColNum("住院号")): strTmp = "住院号" '问题:33789
                strNos = Replace(strNos, "'", "")
                If strColValue = "" Then strColValue = .TextMatrix(.Row, GetColNum("门诊号")): strTmp = "门诊号"
                Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                        "NO=" & strNos, strTmp & "=" & strColValue, _
                        "开单人=" & .TextMatrix(.Row, GetColNum("开单人")))
            End If
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()

    If gbln退费申请模式 Then frmChargeFilter.mblnApply = tbPage.Selected.Index = 1
    frmChargeFilter.mstrPrivs = mstrPrivs
    frmChargeFilter.opt病人(IIf(gint病人来源 = 1, 0, 1)).Value = True
    frmChargeFilter.Show 1, Me
    
    If gblnOK Then
        With frmChargeFilter
            
            If gbln退费申请模式 And tbPage.Selected.Index = 1 Then
                mstrFilter2 = .mstrFilter
                SQLCondition.ApplyName = zlStr.NeedName(.cboApply.Text)
                SQLCondition.ApplyDateB = .dtpApplyB.Value
                SQLCondition.ApplyDateE = .dtpApplyE.Value
                SQLCondition.AuditName = zlStr.NeedName(.cboAudit.Text)
                SQLCondition.AuditDateB = .dtpAuditB.Value
                SQLCondition.AuditDateE = .dtpAuditE.Value
                SQLCondition.int门诊标志 = IIf(.opt病人(0).Value, 0, IIf(.opt病人(1).Value, 1, 2)) + 1
            Else
                mstrFilter = .mstrFilter
                mbln收费 = .chk收费.Value = 1
                mbln退费 = .chk退费.Value = 1
                
                '医保费用
                If .chk普通.Value = 1 And .chk医保.Value = 0 Then
                    mstrInsure = " And Nvl(t.险类,0) = 0"
                ElseIf .chk普通.Value = 0 And .chk医保.Value = 1 Then
                    mstrInsure = " And Nvl(t.险类,0) <> 0"
                Else
                    mstrInsure = ""
                End If
                
                SQLCondition.Default = False
                SQLCondition.DateB = .dtpBegin.Value
                SQLCondition.DateE = .dtpEnd.Value
                SQLCondition.int门诊标志 = IIf(.opt病人(0).Value, 0, IIf(.opt病人(1).Value, 1, 2)) + 1
                If .cbo费别.ListIndex > 0 Then SQLCondition.ChargeKind = zlStr.NeedName(.cbo费别.Text)
                If .cbo付款方式.ListIndex > 0 Then
                    SQLCondition.PayKind = zlStr.NeedCode(.cbo付款方式.Text)
                    SQLCondition.PayKindName = zlStr.NeedName(.cbo付款方式.Text)
                Else
                    SQLCondition.PayKind = ""
                    SQLCondition.PayKindName = ""
                End If
                
                SQLCondition.PatientName = gstrLike & UCase(.txt姓名.Text) & "%"
                SQLCondition.PatientIdentity = Val(.txt标识号.Text)
                SQLCondition.PatientID = .mlngPrePatient
                SQLCondition.NOB = .txtNOBegin.Text
                SQLCondition.NOE = .txtNoEnd.Text
                SQLCondition.FactB = .txtFactBegin.Text
                SQLCondition.FactE = .txtFactEnd.Text
                SQLCondition.DeptID = .cbo科室.ItemData(.cbo科室.ListIndex)
                SQLCondition.Doctor = .txt开单人.Text
                If .cbo操作员.ListIndex = -1 Then
                    SQLCondition.Operator = UserInfo.姓名
                Else
                    SQLCondition.Operator = zlStr.NeedName(.cbo操作员.Text)
                End If
                SQLCondition.FeeItems = .mstrFeeItems
            End If
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshDetail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim blnDo As Boolean '是否已打过票据
    
    If Button <> 2 Then Exit Sub '不是右键退出
    If tbPage.Selected Is Nothing Then Exit Sub
    If Not Me.ActiveControl Is mshDetail _
        Or Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 2 _
        Or Not tbPage.Selected.Index = 0 Then Exit Sub
    
    '按实际打印分配票号且按单据分别打印时，才可能选择某张单据进行补打
    If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 Then
        If mshDetail.IsSubtotal(mshDetail.Row) Then '汇总行
            If mshDetail.Cell(flexcpTextDisplay, mshDetail.Row + 1, mshDetail.ColIndex("发票号")) <> "" Then blnDo = True
        Else
            If mshDetail.Cell(flexcpTextDisplay, mshDetail.Row, mshDetail.ColIndex("发票号")) <> "" Then blnDo = True
        End If
        
        '弹出右键菜单
        If blnDo Then '重打
            Call SetPrintMenu(True)
        Else '补打
            Call SetPrintMenu(False)
        End If
    Else
        If mnuEdit_Print.Visible = False _
            Or Trim(mshList.TextMatrix(mshList.Row, GetColNum("首张发票"))) = "" Then
            Call SetPrintMenu(False, False)
            Exit Sub
        Else
            Call SetPrintMenu(True)
        End If
    End If
End Sub

Private Sub SetPrintMenu(Optional ByVal blnEnable As Boolean, Optional ByVal blnPrintVisible As Boolean = True)
    '功能：设置费用明细列表中的菜单
    If blnEnable Then
        mnuFeeDetial_Print.Visible = blnEnable And blnPrintVisible: mnuFeeDetial_Print.Enabled = blnEnable
        mnuEdit_Print_Supplemental.Enabled = Not blnEnable
        mnuFeeDetial_Supplemental.Visible = Not blnEnable And blnPrintVisible: mnuFeeDetial_Supplemental.Enabled = Not blnEnable
    Else '两个子菜单，必须得有一个可见，先得设置可见那个
        mnuEdit_Print_Supplemental.Enabled = Not blnEnable
        mnuFeeDetial_Supplemental.Visible = Not blnEnable And blnPrintVisible: mnuFeeDetial_Supplemental.Enabled = Not blnEnable
        mnuFeeDetial_Print.Visible = blnEnable And blnPrintVisible: mnuFeeDetial_Print.Enabled = blnEnable
    End If
    '弹出右键菜单
    PopupMenu mnuFeeDetial, 2
End Sub

Private Sub mshList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNo As String, strDate As String, blnDel As Boolean, blnAudit As Boolean
    Dim rsTmp As ADODB.Recordset, bytType As Byte
    Dim bln红票已打印 As Boolean
    
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    If NewRow < mshList.FixedRows Then Exit Sub
    
    With mshList
        If tbPage.Selected.Index <> 1 Then
            strNo = .TextMatrix(NewRow, GetColNum("结算序号"))
            blnDel = Val(.TextMatrix(NewRow, GetColNum("退费符号"))) < 0
            If strNo = "" Then Exit Sub
            bytType = IIf(CheckBalance(Val(strNo)), 2, 1)
            bln红票已打印 = Val(.TextMatrix(NewRow, GetColNum("红票已打印"))) = 1
        Else
            strNo = .TextMatrix(NewRow, GetColNum("首张单据"))
            If strNo = "" Then Exit Sub
        End If
    End With
    If mrsList Is Nothing Then Exit Sub
    If mrsList.State = 0 Then Exit Sub
    If mrsList.RecordCount = 0 Then Exit Sub
    
    Call SetMenuCaption
    mlngGo = NewRow
    mlngCurRow = NewRow: mlngTopRow = mshList.TopRow
    
    If gbln退费申请模式 Then
        If tbPage.Selected.Index = 1 Then
            blnAudit = mshList.TextMatrix(NewRow, GetColNum("审核状态")) = "通过" Or mshList.TextMatrix(NewRow, GetColNum("审核状态")) = "拒绝"
            
            mnuEdit_UnApply.Enabled = Not blnAudit
            mnuEdit_Audit.Enabled = Not blnAudit
            mnuEdit_RefuseApply.Visible = Not blnAudit And Val(mshList.TextMatrix(NewRow, GetColNum("结算序号"))) > 0
            mnuEdit_RefuseApply.Enabled = Not blnAudit And Val(mshList.TextMatrix(NewRow, GetColNum("结算序号"))) > 0
            mnuEdit_UnAudit.Enabled = mshList.TextMatrix(NewRow, GetColNum("审核状态")) = "通过"
            mnuEditWriteCard.Enabled = False
        Else
            strDate = mshList.TextMatrix(NewRow, GetColNum("登记时间"))
            blnDel = Val(mshList.TextMatrix(NewRow, GetColNum("符号"))) = 2
            
            mnuEdit_Apply.Enabled = Not blnDel
            
            mnuEdit_Adjust.Enabled = Not blnDel
            tbr.Buttons("Del").Enabled = Not blnDel
            mnuEdit_Print.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("首张发票"))) <> ""
            mnuFeeDetial_Print.Enabled = mnuEdit_Print.Enabled
            mnuEdit_Print_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("首张发票"))) = ""
            mnuFeeDetial_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("首张发票"))) = ""
            mnuEdit_PrintProve.Enabled = Not blnDel
            mnuFeeDetial_PrintProve.Enabled = Not blnDel
            mnuEdit_PrintList.Enabled = Not blnDel
            mnuFeeDetial_PrintList.Enabled = Not blnDel
            mnuEdit_PrintDel.Enabled = blnDel And IIf(bln红票已打印, InStr(mstrPrivs, ";重打票据;") > 0, InStr(mstrPrivs, ";补打票据;") > 0)
            mnuEdit_PrintDel.Caption = (IIf(bln红票已打印, "重打退费票据(&D)", "补打退费票据(&B)"))
            mnuEditWriteCard.Enabled = strNo <> ""
        End If
    Else
        strDate = mshList.TextMatrix(NewRow, GetColNum("登记时间"))
        blnDel = Val(mshList.TextMatrix(NewRow, GetColNum("退费符号"))) < 0
        
        mnuEdit_Adjust.Enabled = Not blnDel

        tbr.Buttons("Del").Enabled = Not blnDel
        
        mnuEdit_Print.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("首张发票"))) <> ""
        mnuFeeDetial_Print.Enabled = mnuEdit_Print.Enabled
        mnuEdit_Print_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("首张发票"))) = ""
        mnuFeeDetial_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("首张发票"))) = ""
        mnuEdit_PrintProve.Enabled = Not blnDel
        mnuFeeDetial_PrintProve.Enabled = Not blnDel
        mnuEdit_PrintList.Enabled = Not blnDel
        mnuFeeDetial_PrintList.Enabled = Not blnDel
        mnuEdit_PrintDel.Enabled = blnDel And IIf(bln红票已打印, InStr(mstrPrivs, ";重打票据;") > 0, InStr(mstrPrivs, ";补打票据;") > 0)
        mnuEdit_PrintDel.Caption = (IIf(bln红票已打印, "重打退费票据(&D)", "补打退费票据(&B)"))
        mnuEditWriteCard.Enabled = strNo <> ""
    End If
        
    mshList.ForeColorSel = mshList.CellForeColor
    If tbPage.Selected.Index = 1 Then
        Call ShowApplyDetail(strNo)
        Call ShowApplyFactList(strNo)
    Else
        Call ReadListData(bytType, Val(strNo), blnDel)
        Call ShowFactList(strNo)
        Call ShowInvoice(strNo)
        Call ShowBalanceList(strNo)
        Call ShowExtendInfo
    End If
End Sub

Private Sub mshList_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngCurrentRow As Long
    
    '触发AfterRowColChange事件
    lngCurrentRow = mshList.Row
    mshList.Row = -1
    mshList.Row = lngCurrentRow
End Sub

Private Sub mshList_DblClick()
    If mshList.Row <= 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub
Private Sub SetMenuCaption()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置菜单的Caption属性
    '编制:刘兴洪
    '日期:2011-09-04 11:40:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, blnDel As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Not tbPage.Selected.Index = 2 Then Exit Sub
    With mshList
        If .Row <= 0 Then Exit Sub
        strNo = .TextMatrix(.Row, GetColNum("首张单据"))
        blnDel = Val(.TextMatrix(.Row, GetColNum("符号"))) = 2
        If strNo = "" Then Exit Sub
    End With
    mnuEditCancelBill.Caption = IIf(blnDel, "重新退费(&Z)", "作废收费(&Z)")
    tbr.Buttons("Cancel").Caption = IIf(blnDel, "退费", "作废")
    tbr.Buttons("Cancel").ToolTipText = IIf(blnDel, "重新退费异常单据:" & strNo, "作废异常单据:" & strNo)
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
        mshFact.BackColorSel = &HE0E0E0
        vsfExtendInfo.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
        mshFact.BackColorSel = &HE0E0E0
        vsfExtendInfo.BackColorSel = &HE0E0E0
    ElseIf obj Is mshFact Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HE0E0E0
        mshFact.BackColorSel = &HC0C0C0
        vsfExtendInfo.BackColorSel = &HE0E0E0
    ElseIf obj Is vsSubBalance Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HE0E0E0
        mshFact.BackColorSel = &HE0E0E0
        vsSubBalance.BackColorSel = &HC0C0C0
        vsfExtendInfo.BackColorSel = &HE0E0E0
    ElseIf obj Is vsSubInvoice Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HE0E0E0
        mshFact.BackColorSel = &HE0E0E0
        vsSubInvoice.BackColorSel = &HC0C0C0
        vsfExtendInfo.BackColorSel = &HE0E0E0
    ElseIf obj Is vsfExtendInfo Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HE0E0E0
        mshFact.BackColorSel = &HE0E0E0
        vsSubInvoice.BackColorSel = &HE0E0E0
        vsfExtendInfo.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_DelMulti.Enabled And mnuEdit_DelMulti.Visible Then Call mnuEdit_DelMulti_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Me.ActiveControl Is mshDetail Then Exit Sub '使用费用明细列表的弹出菜单
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            '始终从当前行开始
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub DelOldBill()
    Dim strNo As String, strNos As String, strSql As String, i As Long
    Dim intInsure As Integer, blnHaveExe As Boolean, blnFlagPrint As Boolean
    Dim strTempNos As String, lngBalance As Long, blnOneCard As Boolean
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    lngBalance = mshList.TextMatrix(mshList.Row, GetColNum("结算序号"))
    If strNo = "" Then
        MsgBox "当前没有单据可以退费！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If CheckBalance(lngBalance) = False Then
        blnOneCard = GetOneCard.RecordCount > 0
        If frmMultiBills.ShowMe(gfrmMain, 1, mstrPrivs, strNo, "", , , blnOneCard) Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("当前操作已更改单据清单内容,要刷新吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        End If
        Exit Sub
    End If
    
    '权限检查
    If Not BillOperCheck(2, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("收费时间"))), "退费", strNo, , 1) Then Exit Sub
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    intInsure = ChargeExistInsure(strNo)
    
    If zlCheckIsInvoiceListPrinted(strNo, mblnNOMoved) Then
        '按打印明细进行打印时,根据收费次数进行多单据处理
        strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
    Else
        strTempNos = GetMultiNOs(strNo, , mblnNOMoved, True)
        strNos = GetMultiNOs(strNo, , mblnNOMoved, False)
        If InStr(strTempNos, ",") > 0 And InStr(strNos, ",") = 0 Then
            '肯定是按单据分别打印的
            '要多单据退费,就必须满足以下条件
            '1.医保多单据必须全退时,必须按结算序号进行退费
            '2.三方账户全退时,必须按结算序号进行退费
            If intInsure <> 0 Then
                If gclsInsure.GetCapability(support多单据收费必须全退, , intInsure) Then
                    strNos = strTempNos
                End If
            ElseIf zlIsExistsSquareCard(strTempNos, True) Then
                '检查一卡通结算部分是否存在全退的
                strNos = strTempNos
            End If
        End If
    End If
    
    If zlCheckIsMzToZY(strNo, 1) Then
        MsgBox "注意:" & vbCrLf & _
                      "    该单据已经被门诊费用转住院费用 " & vbCrLf & _
                      "    或已经审核了门诊费用转住院费用,不能再退费", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    '医保类型匹配判断(确定时会再重复判断一次,因为还要获取其它医保参数)
    If intInsure > 0 Then
        '保险退费权限检查
        If InStr(mstrPrivs, "保险收费") = 0 Then
            MsgBox "你没有权限对医保病人的单据退费！", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(strNos, ",") > 0 Then
            If gclsInsure.GetCapability(support多单据收费必须全退, , intInsure) Then
                MsgBox "当前医保不允许对其中一张单据退费！", vbInformation, gstrSysName
                Call mnuEdit_DelMulti_Click
                Exit Sub
            End If
        End If
    Else
        If InStr(mstrPrivs, "允许非医保病人") = 0 Then
            MsgBox "你没有权限对非医保病人进行退费操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
        
    If gblnMultiBalance And InStr(strNos, ",") > 0 Then
        If CheckSingleBalance(strNos) = False Then
            MsgBox "多张单据使用多种结算方式模式下不允许对其中一张单据退费！", vbInformation, gstrSysName
            Call mnuEdit_DelMulti_Click
            Exit Sub
        End If
    End If
    
    '刘兴洪:多单据退费，要检查是存在结算卡，存在结算卡的只能全退
    If UBound(Split(strNos, ",")) > 0 Then
        '多单据退费
        If zlIsExistsSquareCard(strNos) = True Then
            '调用多单据退费
            'If MsgBox("注意:" & vbCrLf & "    该张单据存在结算卡消费,不能对其中的一张单据退费,是否调用多单据退费?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call mnuEdit_DelMulti_Click
            Exit Sub
        End If
    End If
    
            
            
    '是否已执行
    i = BillCanDelete(strNo, 1, blnHaveExe, , blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '该单据不存在
                MsgBox "指定的单据不存在！", vbInformation, gstrSysName
            Case 2 '已经全部完全执行
                '不考虑退费自动退药
                MsgBox "该单据中的项目已经全部完全执行！", vbInformation, gstrSysName
            Case 3 '未完全执行部分剩余数量为0
                MsgBox "该单据中未完全执行部分项目剩余数量为零,没有可以退费的费用！", vbInformation, gstrSysName
        End Select
        Exit Sub
    ElseIf intInsure > 0 And blnHaveExe Then
        MsgBox "该医保收费单据中包含已经执行的项目,不能退费！", vbInformation, gstrSysName
        Exit Sub
    ElseIf intInsure = 0 And blnHaveExe Then
        If GetOneCardBalance(Val(mshList.TextMatrix(mshList.Row, GetColNum("结帐ID")))).RecordCount > 0 Then
            MsgBox "该单据由于存在已执行的项目,使用了一卡通结算,不能进行部分退费！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If blnHaveExe Then
        MsgBox "注意:该单据由于存在已执行的项目，当前将执行的是部分退费。", vbInformation, gstrSysName
    End If
    If blnFlagPrint Then
        If MsgBox("注意:检验医嘱的条码已打印，是否继续退费？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
            
    On Error GoTo errH
    If Not isSimple(strNo) Then
        On Error Resume Next    '开两个窗口操作作时，其中一个退出会unload窗体
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 0
        frmCharge.mbytInState = 3
        frmCharge.mstrInNO = strNo
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        If MsgBox("确实要将该单据退费吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If gblnBillPrint Then
            If gobjBillPrint.zlEraseBill("'" & strNo & "'", 0) = False Then Exit Sub
        End If
        
        '简单收费不支持医保,也不提供部分退费
        strSql = "zl_门诊简单收费_DELETE('" & strNo & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, "'" & strNo & "'")
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        
        gblnOK = True
    End If
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEdit_Apply_Click()
    Dim strNo As String, strNos As String, blnTogetherDo As Boolean, strBalance As String
    Dim arrTmp As Variant, i As Long, blnTrans As Boolean, strDate As String, strSql As String
    Dim strReason As String
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("结算序号"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))

    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_退费申请, mstrPrivs, Val(strBalance), , , mblnNOMoved) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "申请退费成功!"
        End If
    Else
        strNos = Replace(GetMultiNOs(strNo, , , True), "'", "")
        arrTmp = Split(strNos, ",")
        '多单据一次收费的历史数据必须一起申请和取消申请，以及拒绝申请，因为在管理窗口中只能选择首张单据，
        '如果按单张单据进行申请，有的单据就选择不到，无法进行申请
'        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNO)

'        If blnTogetherDo Then
        If UBound(arrTmp) > 0 Then
            If MsgBox("单据[" & strNos & "]必须一起申请退费，你确认要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        '71917,冉俊明,2014-4-17,在病人退费申请时增加退费申请原因
        If Not frmInputBox.InputBox(Me, "申请原因", "请输入申请原因：", 100, 2, True, False, strReason, False) Then Exit Sub
        
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                '71917,冉俊明,2014-4-17,在病人退费申请时增加退费申请原因
                strSql = "Zl_病人退费申请_Apply(0,'" & arrTmp(i) & "',1,'" & UserInfo.姓名 & "'," & strDate & ",'" & strReason & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "对单据[" & strNos & "]申请退费成功!"
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEdit_UnApply_Click()
    Dim strNo As String, strNos As String, blnTogetherDo As Boolean
    Dim arrTmp As Variant, i As Long, blnTrans As Boolean, strDate As String, strSql As String
    Dim strApplicant As String, strBalance As String
    
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("结算序号"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    strDate = mshList.TextMatrix(mshList.Row, GetColNum("申请时间"))
    strApplicant = mshList.TextMatrix(mshList.Row, GetColNum("申请人"))

    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_取消申请, mstrPrivs, Val(strBalance), , , mblnNOMoved, , strDate) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "取消退费申请成功!"
        End If
    Else
        If MsgBox("确实要取消单据[" & strNo & "]的退费申请吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        If InStr(1, mstrPrivs, "所有操作员") = 0 Then
            If mshList.TextMatrix(mshList.Row, GetColNum("申请人")) <> UserInfo.姓名 Then
                MsgBox "你没有权限取消他人的退费申请单！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
            
        strNos = Replace(GetMultiNOs(strNo, , , True), "'", "")
        If CheckBalance(, strNo) Then strNos = strNo
        arrTmp = Split(strNos, ",")
        '多单据一次收费的历史数据必须一起申请和取消申请，以及拒绝申请，因为在管理窗口中只能选择首张单据，
        '如果按单张单据进行申请，有的单据就选择不到，无法进行申请
'        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNO)

'        If blnTogetherDo Then
        If UBound(arrTmp) > 0 Then
            If MsgBox("单据[" & strNos & "]必须一起取消退费申请，你确认要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                '71917,冉俊明,2014-4-17,在病人退费申请时增加退费申请原因
                strSql = "Zl_病人退费申请_Apply(1,'" & arrTmp(i) & "',1,'" & strApplicant & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "对单据[" & strNos & "]取消退费申请成功!"
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEdit_UnAudit_Click()
    Dim strNo As String, strNos As String, blnTogetherDo As Boolean
    Dim arrTmp As Variant, i As Long, blnTrans As Boolean, strSql As String
    Dim strApplyDate As String, strReason As String, strDate As String, strBalance As String
    
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("结算序号"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    strApplyDate = mshList.TextMatrix(mshList.Row, GetColNum("申请时间"))
    
    If InStr(1, mstrPrivs, "所有操作员") = 0 Then
        If mshList.TextMatrix(mshList.Row, GetColNum("审核人")) <> "" And mshList.TextMatrix(mshList.Row, GetColNum("审核人")) <> UserInfo.姓名 Then
            MsgBox "你没有取消审核他人审核的退费申请单！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_取消审核, mstrPrivs, Val(strBalance), , , mblnNOMoved, , strApplyDate) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "取消审核成功!"
        End If
    Else
        If MsgBox("确实要取消单据[" & strNo & "]的审核吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    
        strNos = Replace(GetMultiNOs(strNo), "'", "")
        arrTmp = Split(strNos, ",")
        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNo)
        
        If blnTogetherDo Then
            If MsgBox("单据[" & strNos & "]必须一起取消审核，你确认要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            Else
                For i = 0 To UBound(arrTmp)
                    If BillExistDelete(arrTmp(i), 1) Then
                        MsgBox "单据[" & arrTmp(i) & "]已退费，不能取消审核。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next
            End If
        Else
            If BillExistDelete(strNo, 1) Then
                MsgBox "单据[" & strNo & "]已退费，不能取消审核。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                strSql = "Zl_病人退费申请_Audit('" & arrTmp(i) & "',1,To_Date('" & strApplyDate & "','YYYY-MM-DD HH24:MI:SS')," & _
                     "NULL,NULL,NULL,3)"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "已对单据[" & strNos & "]取消审核！"
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub mnuEdit_Audit_Click()
    Dim strNo As String, strNos As String, blnTogetherDo As Boolean
    Dim arrTmp As Variant, i As Long, blnTrans As Boolean, strDate As String, strSql As String
    Dim strApplyDate As String, strReason As String, strBalance As String
    Dim blnHaveExe As Boolean, strInfos As String
    
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("结算序号"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    strApplyDate = mshList.TextMatrix(mshList.Row, GetColNum("申请时间"))

    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_退费审核, mstrPrivs, Val(strBalance), , , mblnNOMoved, , strApplyDate) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "单据已审核！"
        End If
    Else
        If MsgBox("确实要将单据[" & strNo & "]审核吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        strNos = Replace(GetMultiNOs(strNo), "'", "")
        arrTmp = Split(strNos, ",")
        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNo)
        
        If blnTogetherDo Then
            If MsgBox("单据[" & strNos & "]必须一起审核，你确认要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        
        '检查是否存在已执行的项目
        For i = 0 To UBound(arrTmp)
            Call BillCanDelete(arrTmp(i), 1, blnHaveExe)
            If blnHaveExe Then
                strInfos = strInfos & "," & arrTmp(i)
            End If
        Next
        If strInfos <> "" Then
            strInfos = Mid(strInfos, 2)
            If MsgBox("单据[" & strInfos & "]中存在已执行的项目，你确认要继续吗？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
'        If Not frmInputBox.InputBox(Me, "审核原因", "请输入审核原因：", 100, 2, True, False, strReason, False) Then Exit Sub

        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                strSql = "Zl_病人退费申请_Audit('" & arrTmp(i) & "',1,To_Date('" & strApplyDate & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "'" & UserInfo.姓名 & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),'" & strReason & "',1)"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "单据[" & strNos & "]已审核！"
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEdit_RefuseApply_Click()
    Dim strNo As String, strNos As String, blnTogetherDo As Boolean
    Dim arrTmp As Variant, i As Long, blnTrans As Boolean, strSql As String
    Dim strApplyDate As String, strReason As String, strDate As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
    strApplyDate = mshList.TextMatrix(mshList.Row, GetColNum("申请时间"))
    
    If InStr(1, mstrPrivs, "所有操作员") = 0 Then
        If mshList.TextMatrix(mshList.Row, GetColNum("审核人")) <> "" And mshList.TextMatrix(mshList.Row, GetColNum("审核人")) <> UserInfo.姓名 Then
            MsgBox "你没有权限拒绝他人审核的退费申请单！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    strNos = Replace(GetMultiNOs(strNo, , , True), "'", "")
    arrTmp = Split(strNos, ",")
    '多单据一次收费的历史数据必须一起申请和取消申请，以及拒绝申请，因为在管理窗口中只能选择首张单据，
    '如果按单张单据进行申请，有的单据就选择不到，无法进行申请
'    If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNo)

'    If blnTogetherDo Then
    If UBound(arrTmp) > 0 Then
        If MsgBox("单据[" & strNos & "]必须一起拒绝申请，你确认要继续吗？", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
        Else
            For i = 0 To UBound(arrTmp)
                If BillExistDelete(arrTmp(i), 1) Then
                    MsgBox "单据[" & arrTmp(i) & "]已退费，不能拒绝申请。", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End If
    End If
    If Not frmInputBox.InputBox(Me, "拒绝原因", "请输入拒绝原因：", 100, 2, False, False, strReason, False) Then Exit Sub
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrTmp)
            strSql = "Zl_病人退费申请_Audit('" & arrTmp(i) & "',1,To_Date('" & strApplyDate & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                 UserInfo.姓名 & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),'" & strReason & "',2)"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Call mnuViewReFlash_Click
    stbThis.Panels(2).Text = "已对单据[" & strNos & "]拒绝申请！"
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckErrBill() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在异常的收费单据
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 15:27:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date, strErrWhere As String
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim blnDel As Boolean, strLast As String
    
    If InStr(mstrPrivs, ";门诊收费;") = 0 Then Exit Function
    
    On Error GoTo errHandle
    Select Case cboDate.ListIndex
       Case 0 '今日
           dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
       Case 1 '最近两天
           dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case 2 '最近三天
           dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case 3  '最近一周
           dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case 4  '本月
           dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-01") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case Else
           dtStartDate = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
       End Select
       lblDateShow.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
       lblDateShow.Caption = lblDateShow.Caption & "～" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
       
    '收费异常记录
    strSql = _
    " Select Count(distinct nvl(B.结算序号,B.结帐ID)) as 条数" & vbNewLine & _
    " From 门诊费用记录 A,病人预交记录 B" & _
    " Where A.结帐ID=B.结帐ID And Nvl(A.费用状态, 0) = 1 And A.记录性质 = 1  " & _
    "       And A.记录状态 = 1 And A.登记时间 Between [1] And [2] " & _
    "       And A.操作员姓名 = [3] " & vbNewLine & _
    "       And Not Exists (Select 1 From 门诊费用记录 Q Where a.No = Q.No And Mod(Q.记录性质, 10) = 1 And Q.记录状态 = 2) "
        

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
    CheckErrBill = rsTemp!条数 <> 0
    If rsTemp!条数 <> 0 Then
        tbPage.Item(2).Caption = "收费异常记录(" & rsTemp!条数 & ")"
        If tbPage.Selected.Index <> 2 Then
            If MsgBox("存在收费异常记录,是否处理收费异常记录?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tbPage.Item(2).Selected = True
                Call mnuEditReCharge_Click
                Exit Function
            End If
        End If
    Else
        tbPage.Item(2).Caption = "收费异常记录"
    End If
    
    '退费异常记录
    strSql = "" & _
        " Select /*+ Rule*/ Count(distinct nvl(B.结算序号,B.结帐ID)) as 条数 " & vbNewLine & _
        " From 门诊费用记录 A,病人预交记录 B" & vbNewLine & _
        " Where Nvl(A.费用状态, 0) = 1 And Mod(A.记录性质, 10) = 1 " & _
        "       And A.记录状态 = 2 And A.登记时间 Between [1] And [2] And A.操作员姓名 = [3] " & _
        "       And Exists (Select 1 From 病人预交记录 Q Where A.结帐id = Q.结帐id And Nvl(Q.校对标志, 0) <> 0)  " & _
        "       And A.结帐id =B.结帐id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
    
    CheckErrBill = rsTemp!条数 <> 0
    If rsTemp!条数 <> 0 Then
        tbPage.Item(3).Caption = "退费异常记录(" & rsTemp!条数 & ")"
        If tbPage.Selected.Index <> 3 Then
            If MsgBox("存在退费异常记录,是否处理退费异常记录?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tbPage.Item(3).Selected = True
                Call mnuEditReCharge_Click
                Exit Function
            End If
        End If
    Else
        tbPage.Item(3).Caption = "退费异常记录"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReChargeToErrBill(ByVal lng结算序号 As Long, Optional ByVal blnDel As Boolean = False, _
    Optional ByVal bln退费异常 As Boolean = False, Optional ByVal strDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新收取异常的单据费用
    '入参:
    '   blnDel True-作废单据,False-重新收费
    '   bln退费异常 是否退费异常单据
    '   strDate 收费时间
    '返回:收取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 15:41:08
    '说明：若blnDel=True And bln退费异常=True表示作废单据时产生的异常单据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivsReplenish As String
    Dim blnOK As Boolean
    
    On Error GoTo errHandle
    If bln退费异常 = False Then
        If zlIsCheckExiseSingularity(lng结算序号) Then
            MsgBox "该异常单据已经被作废，因此，不能再" & IIf(blnDel, "进行作废", "重新收费") & "，请刷新费用列表！", vbInformation, gstrSysName
            Exit Function
        End If
        If Not zlIsCheckExistErrBill(lng结算序号) Then
            MsgBox "该异常单据已经被重新收费，因此，不能再" & IIf(blnDel, "进行作废", "重新收费") & "，请刷新费用列表！", vbInformation, gstrSysName
            Exit Function
        End If
    
        blnOK = frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, IIf(blnDel, EM_ED_异常作废, EM_ED_异常重收), , lng结算序号, mblnNOMoved, , , mobjMsgModule)
    Else
        If blnDel Then
            '对作废收费异常记录产生的异常进行处理
            blnOK = frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, EM_ED_异常作废, , lng结算序号, mblnNOMoved, , , mobjMsgModule, , True)
        Else
            If CheckBillExistReplenishData(0, lng结算序号) Then

                strPrivsReplenish = ";" & GetPrivFunc(glngSys, 1124) & ";"
                If InStr(strPrivsReplenish, ";结算退费;") > 0 Then
                    If MsgBox("选择的记录进行了医保补充结算且为异常补充结算退费记录，是否针对该记录进行再次结算退费？", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        blnOK = frmReplenishTheBalanceDel.zlShowMe(Me, 1124, strPrivsReplenish, EM_RBDTY_异常重退, lng结算序号, False, 0, False, strDate)
                    Else
                        Exit Function
                    End If
                Else
                    MsgBox "选择的记录进行了医保补充结算且为异常补充结算退费记录，你不具备操作该记录的权限，不允许进行退费操作！", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                blnOK = frmClinicDelAndView.ShowMe(Me, EM_MULTI_异常重退, mstrPrivs, lng结算序号, , , , strDate)
            End If
        End If
    End If

    If blnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mblnNotClick = True
                mnuViewReFlash_Click
                mblnNotClick = False
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
        ReChargeToErrBill = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReChargeToDelErrBill(ByVal strNo As String, strDelTime As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新对异常退费单进行退费操作
    '入参:strDelTime-异常单据的退费时间
    '
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-22 15:41:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOneCard As Boolean
    On Error GoTo errHandle
    If strNo = "" Then Exit Function
    If InStr(mstrPrivs, ";门诊退费;") = 0 Then Exit Function
    blnOneCard = GetOneCard.RecordCount > 0
    
    If frmMultiBills.ShowMe(gfrmMain, 2, mstrPrivs, strNo, strDelTime, , , blnOneCard) = False Then Exit Function
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
    ReChargeToDelErrBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function CheckTogetherDo(ByVal strNo As String) As Boolean
    Dim intInsure As Integer
    
    intInsure = ChargeExistInsure(strNo)
    If intInsure > 0 Then
        If gclsInsure.GetCapability(support多单据收费必须全退, , intInsure) Then CheckTogetherDo = True
    End If
    
    If gblnMultiBalance Then
        If CheckSingleBalance(strNo) = False Then CheckTogetherDo = True
    End If
End Function

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Charge_Click()
    If frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, EM_ED_收费, , , , , , mobjMsgModule) = True Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改记录内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mblnNotClick = True
                mnuViewReFlash_Click
                mblnNotClick = False
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
        End If
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim strNo As String, strNos As String
    Dim strDate As String, blnDel As Boolean
    '重大改进后，传入的是结算序号
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("结算序号"))
    If tbPage.Selected.Index = 1 Then
        strNo = mshList.TextMatrix(mshList.Row, GetColNum("单据号"))
        '退费申请界面，根据单据查看
        If isSimple(strNo) Then
            '简单收费
            frmSimpleCharge.mlngModul = mlngModul
            frmSimpleCharge.mstrPrivs = mstrPrivs
            frmSimpleCharge.mbytInState = 1
            frmSimpleCharge.mstrDelete = IIf(blnDel, strDate, "")     '只有退费单据才传入时间以区别正常状态单据
            frmSimpleCharge.mstrInNO = strNo
            frmSimpleCharge.mblnNOMoved = mblnNOMoved
            frmSimpleCharge.Show 1, Me
        Else
            '正常收费
            frmCharge.mlngModul = mlngModul
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 1
            frmCharge.mblnDelete = blnDel
            frmCharge.mstrTime = strDate
            frmCharge.mstrInNO = strNo
            frmCharge.mblnNOMoved = mblnNOMoved
            Set frmCharge.mobjMsgModule = mobjMsgModule
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        End If
        Exit Sub
    End If
    If strNo = "" Then
        MsgBox "当前没有单据可以查阅！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    If Not (gbln退费申请模式 And tbPage.Selected.Index = 1) Then
        '是否查看退费单据
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("退费符号"))) < 0
        strDate = mshList.TextMatrix(mshList.Row, GetColNum("收费时间"))
    End If
    
    If Val(strNo) < 0 Then
        If blnDel Then
            frmClinicDelAndView.ShowMe Me, EM_MULTI_查看, mstrPrivs, Val(strNo), , , mblnNOMoved, strDate
        Else
            frmClinicDelAndView.ShowMe Me, EM_MULTI_查看, mstrPrivs, Val(strNo), , , mblnNOMoved
        End If
    Else
        If CheckBalance(Val(strNo)) Then
            strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
            strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
            If UBound(Split(strNos, ",")) > 0 Then
                frmMultiBills.ShowMe gfrmMain, 0, mstrPrivs, strNo, IIf(blnDel, strDate, ""), , , , mblnNOMoved      '只有退费单据才传入时间以区别正常状态单据
            ElseIf isSimple(strNo) Then
                '简单收费
                frmSimpleCharge.mlngModul = mlngModul
                frmSimpleCharge.mstrPrivs = mstrPrivs
                frmSimpleCharge.mbytInState = 1
                frmSimpleCharge.mstrDelete = IIf(blnDel, strDate, "")     '只有退费单据才传入时间以区别正常状态单据
                frmSimpleCharge.mstrInNO = strNo
                frmSimpleCharge.mblnNOMoved = mblnNOMoved
                frmSimpleCharge.Show 1, Me
            Else
                '正常收费
                frmCharge.mlngModul = mlngModul
                frmCharge.mstrPrivs = mstrPrivs
                frmCharge.mbytInState = 1
                frmCharge.mblnDelete = blnDel
                frmCharge.mstrTime = strDate
                frmCharge.mstrInNO = strNo
                frmCharge.mblnNOMoved = mblnNOMoved
                Set frmCharge.mobjMsgModule = mobjMsgModule
                frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
            End If
        Else
            strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
            frmMultiBills.ShowMe gfrmMain, 0, mstrPrivs, strNo, IIf(blnDel, strDate, ""), , , , mblnNOMoved      '只有退费单据才传入时间以区别正常状态单据
        End If
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    ShowBills IIf(gbln退费申请模式 And tbPage.Selected.Index = 1, mstrFilter2, mstrFilter)
    If mblnNotClick = False Then Call CheckErrBill
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mshFact_GotFocus()
    Call SetActiveList(mshFact)
End Sub

Private Sub picExtendInfo_Resize()
    On Error Resume Next
    With vsfExtendInfo
        .Top = 0
        .Left = 0
        .Width = picExtendInfo.Width
        .Height = picExtendInfo.Height
    End With
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        picVsc.Top = picVsc.Top + Y
        picVsc.Height = picVsc.Height - Y
        mshFact.Top = mshFact.Top + Y
        mshFact.Height = mshFact.Height - Y
        tbSub.Top = tbSub.Top + Y
        tbSub.Height = tbSub.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub picSubBalance_Resize()
    On Error Resume Next
    With vsSubBalance
        .Top = 0
        .Left = 0
        .Width = picSubBalance.Width
        .Height = picSubBalance.Height
    End With
End Sub

Private Sub picSubInvoice_Resize()
    On Error Resume Next
    With vsSubInvoice
        .Height = picSubInvoice.Height
        .Width = picSubInvoice.Width
        .Left = 0
        .Top = 0
    End With
End Sub

Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshDetail.Width + X < 1000 Or mshFact.Width - X < 1000 Then Exit Sub
        picVsc.Left = picVsc.Left + X
        mshDetail.Width = mshDetail.Width + X
        mshFact.Left = mshFact.Left + X
        mshFact.Width = mshFact.Width - X
        tbSub.Left = tbSub.Left + X
        tbSub.Width = tbSub.Width - X
        Me.Refresh
    End If
End Sub

Private Sub picVsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strFilter As String
    If Me.Visible = False Then Exit Sub
    If mblnNotClick Then Exit Sub
    If tbPage.ItemCount = 4 Then
        Unload frmChargeFilter
    End If
    If tbPage.Selected.Index = 1 Then
        strFilter = mstrFilter2
        Call SaveFlexState(mshList, App.ProductName & "\" & Me.Name & Val(tbPage.Tag))
    Else
        strFilter = mstrFilter
        Call SaveFlexState(mshList, App.ProductName & "\" & Me.Name & Val(tbPage.Tag))
    End If
    picCons.Visible = tbPage.Selected.Index = 2 Or tbPage.Selected.Index = 3
    
    SQLCondition.Default = True
    frmChargeFilter.mblnDateMoved = False
    
    tbPage.Tag = Item.Index
    Call 权限控制
    Call ShowBills(strFilter)
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '定位
            mnuViewGo_Click
        Case "Filter" '过滤
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "Charge"
            If tbPage.Selected.Index = 3 Or tbPage.Selected.Index = 2 Then
                mnuEditReCharge_Click
            Else
                mnuEdit_Charge_Click
            End If
'        Case "Modi"
'            mnuEdit_Modi_Click
        Case "Del"
            mnuEdit_DelMulti_Click
        Case "Cancel"
            mnuEditCancelBill_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "轧帐"
            mnuFileRollingCurtain_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call ExcPlugInFun(ButtonMenu.Tag)
End Sub

Private Sub LoadPlugInMnu()
    Dim strTmp As String
    Dim arrTmp As Variant
    Dim i As Integer
    Dim blnHave As Boolean, blnTool As Boolean
    Dim strTemp As String
    Dim intToolCounter As Integer
    
    blnHave = Not gobjPlugIn Is Nothing
    
    mnuEdit_Extra.Visible = blnHave
    tbr.Buttons("Extra").Visible = blnHave
    
    If blnHave Then
        blnTool = False
        On Error Resume Next
        strTmp = gobjPlugIn.GetFuncNames(glngSys, glngModul, 3)
        Call zlPlugInErrH(Err, "GetFuncNames")
        Err.Clear: On Error GoTo 0
        
        If strTmp = "" Then
            mnuEdit_Extra.Visible = False
            tbr.Buttons("Extra").Visible = False
            Exit Sub
        End If
        
        strTmp = Replace(strTmp, "Auto:", "")
        arrTmp = Split(strTmp, ",")
        intToolCounter = 0
        For i = 0 To UBound(arrTmp)
            If i <> 0 Then
                Load mnuEdit_ExtraItem(i)
            End If
            mnuEdit_ExtraItem(i).Caption = Replace(CStr(arrTmp(i)), "InTool:", "")
            mnuEdit_ExtraItem(i).Tag = Replace(CStr(arrTmp(i)), "InTool:", "")
            
            If InStr(CStr(arrTmp(i)), "InTool:") > 0 Then
                strTemp = Split(CStr(arrTmp(i)), ":")(1)
                blnTool = True
                If intToolCounter <> 0 Then
                    tbr.Buttons("Extra").ButtonMenus.Add tbr.Buttons("Extra").ButtonMenus.Count + 1, strTemp, strTemp
                    intToolCounter = intToolCounter + 1
                End If
                tbr.Buttons("Extra").ButtonMenus(tbr.Buttons("Extra").ButtonMenus.Count).Text = strTemp
                tbr.Buttons("Extra").ButtonMenus(tbr.Buttons("Extra").ButtonMenus.Count).Tag = strTemp
            End If
        Next
        tbr.Buttons("Extra").Visible = blnTool
    End If
End Sub

Private Sub mnuEdit_ExtraItem_Click(Index As Integer)
    Call ExcPlugInFun(mnuEdit_ExtraItem(Index).Tag)
End Sub

Private Sub ExcPlugInFun(ByVal strFunName As String)
    Dim lngPatiID As Long
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    
    If strNo = "" Or strNo = "首张单据" Then
        MsgBox "未选中任何单据，不能执行此操作！", vbExclamation, gstrSysName: Exit Sub
    End If
        
    If Not gobjPlugIn Is Nothing Then
        lngPatiID = zlGet病人ID(strNo)
        On Error Resume Next
        Call gobjPlugIn.ExecuteFunc(glngSys, glngModul, strFunName, lngPatiID, strNo, 0, "", 4)
        Call zlPlugInErrH(Err, "ExecuteFunc")
        Err.Clear: On Error GoTo 0
    End If
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte
    Dim intCurrentRow As Integer
    
    On Error GoTo errHandler
    
    '表头
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    If tbPage.Selected Is Nothing Then Exit Sub
    Select Case tbPage.Selected.Index
    Case 1
        objOut.Title.Text = "退费申请记录清单"
    Case 2
        objOut.Title.Text = "收费异常记录清单"
    Case 3
        objOut.Title.Text = "退费异常记录清单"
    Case Else
        If glngSys Like "8??" Then
            objOut.Title.Text = "药店收费单据清单"
        Else
            objOut.Title.Text = "收退单据记录清单"
        End If
    End Select
    
    '表项
    If tbPage.Selected.Index = 0 Then
        objRow.Add "时间：" & Format(SQLCondition.DateB, "yyyy-mm-dd hh:mm:ss") & " 至 " & Format(SQLCondition.DateE, "yyyy-mm-dd hh:mm:ss")
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    mshList.Redraw = False
    intCurrentRow = mshList.Row
    mblnPrinting = True
    
    '表体
    Set objOut.Body = mshList
    '输出
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mblnPrinting = False
    mshList.Row = intCurrentRow
    mshList.Redraw = True
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mblnPrinting = False
    mshList.Row = intCurrentRow
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub
 
Private Sub SetMenu(blnUsed As Boolean)
'功能：根据有无记录设置菜单可用状态
    Dim blnApply As Boolean
    
    If gbln退费申请模式 And tbPage.Selected.Index = 1 Then blnApply = True
    
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEditReCharge.Enabled = blnUsed
    
    mnuEdit_Adjust.Enabled = blnUsed And Not blnApply
'    mnuEdit_Modi.Enabled = blnUsed And Not blnApply
'    tbr.Buttons("Modi").Enabled = blnUsed And Not blnApply
'    mnuEdit_Del.Enabled = blnUsed And Not blnApply
    tbr.Buttons("Del").Enabled = blnUsed And Not blnApply
    mnuEdit_View.Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed
    
    mnuEdit_Print.Enabled = blnUsed And Not blnApply
    mnuFeeDetial_Print.Enabled = mnuEdit_Print.Enabled
    mnuEdit_Print_Supplemental.Enabled = blnUsed And Not blnApply
    mnuFeeDetial_Supplemental.Enabled = blnUsed And Not blnApply
    mnuEdit_PrintProve.Enabled = blnUsed And Not blnApply
    mnuFeeDetial_PrintProve.Enabled = blnUsed And Not blnApply
    mnuEdit_PrintList.Enabled = blnUsed And Not blnApply
    mnuFeeDetial_PrintList.Enabled = blnUsed And Not blnApply
    mnuEdit_PrintDel.Enabled = blnUsed And Not blnApply
    
    mnuViewGo.Enabled = blnUsed And Not blnApply
    tbr.Buttons("Go").Enabled = blnUsed And Not blnApply
        
    mnuEdit_Apply.Enabled = blnUsed And Not (tbPage.Selected.Index = 1) And gbln退费申请模式
    mnuEdit_UnApply.Enabled = blnUsed And blnApply
    mnuEdit_Audit.Enabled = blnUsed And blnApply
    mnuEdit_RefuseApply.Visible = blnUsed And blnApply
    mnuEdit_RefuseApply.Visible = blnUsed And blnApply
    mnuEdit_UnAudit.Enabled = blnUsed And blnApply
    
    mnuEdit_DelMulti.Enabled = Not blnApply
    mnuEditWriteCard.Enabled = blnUsed
End Sub

Private Sub 权限控制()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:权限控制
    '编制:刘兴洪
    '日期:2011-09-02 15:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, blnErrRefund As Boolean
    Dim blnErrPage As Boolean   '异常单据页面
    If Not tbPage.Selected Is Nothing Then
        blnErrPage = tbPage.Selected.Index = 2 Or tbPage.Selected.Index = 3
        blnErrRefund = tbPage.Selected.Index = 3
    Else
        blnErrPage = False
        blnErrRefund = False
    End If
    If tbPage.Selected.Index = 3 Then
        tbr.Buttons("Charge").Caption = "退费"
    Else
        tbr.Buttons("Charge").Caption = "收费"
    End If
    mnuEditReCharge.Caption = IIf(blnErrRefund, "重新退费(&R)", "重新收费")
    If glngSys Like "8??" Then
        mshFact.Visible = False
        picVsc.Visible = False
        mnuEdit_Simple.Visible = False
        mnuEdit_Charge_.Visible = False
        mnuEditReCharge.Visible = False
    End If
    mnuEdit_Apply_.Visible = gbln退费申请模式
    '---------------------------------------------------------------
    blnHavePrivs = InStr(mstrPrivs, ";门诊收费;") > 0
    mnuEdit_Charge.Visible = blnHavePrivs
    mnuEdit_Simple.Visible = blnHavePrivs
    mnuEdit_Charge_.Visible = blnHavePrivs And Not blnErrPage
    mnuEditReCharge.Visible = blnHavePrivs And blnErrPage
    mnuEditCancelBill.Visible = blnHavePrivs And blnErrPage And Not blnErrRefund
    tbr.Buttons("Charge").Visible = blnHavePrivs
    '工具栏的作废按钮
    tbr.Buttons("Cancel").Visible = blnHavePrivs And blnErrPage And Not blnErrRefund
    '-----------------------------------------------------------
    '重打仅是控制这里的打印,不控制保存时的打印,所以与报表的权限不同。
    mnuEdit_Print.Visible = Not (InStr(mstrPrivs, ";重打票据;") = 0 Or InStr(mstrPrivs, "收据打印") = 0) And Not blnErrPage
    mnuEditInvoicePrint.Visible = gTy_Module_Para.byt票据分配规则 <> 0 And Not (InStr(mstrPrivs, ";重打票据;") = 0 Or InStr(mstrPrivs, "收据打印") = 0) And Not blnErrPage
    
    '52328: 补打票据
    mnuEdit_Print_Supplemental.Visible = InStr(mstrPrivs, ";补打票据;") > 0 And InStr(mstrPrivs, ";收据打印;") > 0 And Not blnErrPage
    
    mnuEditMakeupPrn.Visible = InStr(mstrPrivs, ";补打票据;") > 0 And InStr(mstrPrivs, ";收据打印;") > 0 And Not blnErrPage
    mnuEdit_PrintProve.Visible = InStr(mstrPrivs, ";证明打印;") > 0 And Not blnErrPage
    mnuEdit_PrintList.Visible = InStr(mstrPrivs, ";打印清单;") > 0 And Not blnErrPage
    mnuEdit_PrintDel.Visible = (InStr(mstrPrivs, ";重打票据;") > 0 Or InStr(mstrPrivs, ";补打票据;") > 0) And InStr(mstrPrivs, ";收据打印;") > 0 And Not blnErrPage
    
    blnHavePrivs = InStr(mstrPrivs, ";重打票据;") > 0 Or InStr(mstrPrivs, ";收据打印;") > 0 _
        Or InStr(mstrPrivs, ";证明打印;") > 0 Or InStr(mstrPrivs, ";打印清单;") > 0
    mnuEdit_View_.Visible = blnHavePrivs And Not blnErrPage
    
'    mnuEdit_Modi.Visible = InStr(mstrPrivs, ";记录修改;") > 0 And Not blnErrPage
'    tbr.Buttons("Modi").Visible = InStr(mstrPrivs, ";记录修改;") > 0 And Not blnErrPage
    
    mnuEdit_Adjust.Visible = InStr(mstrPrivs, ";记录调整;") > 0 And Not blnErrPage
    mnuEdit_Adjust_.Visible = (InStr(mstrPrivs, ";记录修改;") > 0 Or InStr(mstrPrivs, ";记录调整;") > 0) And Not blnErrPage
    '-------------------------------------------------------------
    '门诊退费控制
    blnHavePrivs = InStr(mstrPrivs, ";门诊退费;") > 0
'    mnuEdit_Del.Visible = blnHavePrivs And Not blnErrPage
    mnuEdit_Del_.Visible = blnHavePrivs
    tbr.Buttons("Del").Visible = blnHavePrivs And Not blnErrPage
    tbr.Buttons("Del_").Visible = blnHavePrivs
    mnuEdit_DelMulti.Visible = blnHavePrivs And Not blnErrPage
    mnuEdit_Apply.Visible = blnHavePrivs And Not blnErrPage And gbln退费申请模式
    mnuEdit_UnApply.Visible = blnHavePrivs And Not blnErrPage And gbln退费申请模式
    '-------------------------------------------------------------
    '问题:36076
    mnuEditMzToZyDel.Visible = InStr(mstrPrivs, ";转住院费用退费;") > 0 And Not mbln立即销帐 And Not blnErrPage
    mnuEditSplitMzToZy.Visible = InStr(mstrPrivs, ";转住院费用退费;") > 0 And Not mbln立即销帐 And Not blnErrPage
    '-------------------------------------------------------------
    mnuEdit_Audit.Visible = InStr(mstrPrivs, ";退费审核;") > 0 And Not blnErrPage And gbln退费申请模式
    mnuEdit_RefuseApply.Visible = InStr(mstrPrivs, ";退费审核;") > 0 And Not blnErrPage And gbln退费申请模式 And tbPage.Selected.Index = 1
    mnuEdit_UnAudit.Visible = InStr(mstrPrivs, ";退费审核;") > 0 And Not blnErrPage And gbln退费申请模式
    '-------------------------------------------------------------
    mnuEditSplitMzToZy.Visible = InStr(mstrPrivs, ";门诊信息写卡;") > 0 And mstrWriteCardTypeIDs <> "" And Not blnErrPage And tbPage.Selected.Index = 0
    mnuEditWriteCard.Visible = InStr(mstrPrivs, ";门诊信息写卡;") > 0 And mstrWriteCardTypeIDs <> "" And Not blnErrPage And tbPage.Selected.Index = 0
        
    '收费轧帐管理
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";轧帐;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("轧帐").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim sngTabWidth As Single
    mstrPrivs = gstrPrivs
    mlngModul = glngModul: mblnFirst = True
    
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    If Not gobjSquare Is Nothing Then
        If Not gobjSquare.objSquareCard Is Nothing Then
            mstrWriteCardTypeIDs = gobjSquare.objSquareCard.zlGetAvailabilityWriteCardType
        End If
    End If
    Call InitTabSub
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即销帐", glngSys, 1131)) = 1
    If glngSys Like "8??" Then
        Caption = "药店收费管理": Me.mnuEdit_Charge.Caption = "药店收费(&A)"
    End If
    i = Val(zlDatabase.GetPara("异常单据查询", glngSys, mlngModul, 0, Array(lbl缺省, cboDate)))
    With cboDate
        .Clear
        .AddItem "今日"
        .ListIndex = .NewIndex
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "最近两天"
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "最近三天"
        If i = 2 Then .ListIndex = .NewIndex
        .AddItem "最近一周"
        If i = 3 Then .ListIndex = .NewIndex
        .AddItem "本月"
        If i = 4 Then .ListIndex = .NewIndex
        .AddItem "自定义"
        If i = 5 Then .ListIndex = .NewIndex
        dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpEndDate.MaxDate = dtpStartDate.MaxDate
        dtpEndDate.Value = dtpEndDate.MaxDate
        dtpStartDate.Value = DateAdd("d", -7, dtpEndDate.MaxDate)
    End With
    
    If Not gbln退费申请模式 Then
        tbPage.Item(1).Visible = False
        picCons.Left = picCons.Left - 1200
    End If
    mblnNotClick = True
    tbPage.Item(0).Selected = True
    mblnNotClick = False
    'tbPage.Visible = gbln退费申请模式
    Call RestoreWinState(Me, App.ProductName)
    Set stbThis.Panels(5).Picture = Me.Picture
    
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("刷新方式", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    '创建并检测税控打印对象
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(1)
        End If
        On Error GoTo 0
    End If
    
     '创建第三方票据打印部件
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.编号, UserInfo.姓名)
    End If
    On Error GoTo 0
    
    '权限设置
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1121_1")
    Call 权限控制
    
    Call ClearErrInvoice
    
    If InStr(mstrPrivs, "LED与语音") = 0 Then gblnLED = False
    
    '缺省过滤条件(当天内)
    mstrInsure = "" '缺省包含医保的普通费用
    If gbln退费申请模式 Then
        mstrFilter2 = " And A.申请时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And Nvl(A.状态,0) = 0"
    End If
    mstrFilter = " And 登记时间 Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And 操作员姓名||''=[7]"
    
    SQLCondition.Default = True
    SQLCondition.DateB = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    SQLCondition.DateE = DateAdd("s", -1, DateAdd("d", 1, SQLCondition.DateB))
    frmChargeFilter.mblnDateMoved = False
    
    mbln收费 = True
    mbln退费 = False
    
    Call SetHeader
    Call SetInvoiceList
    Call SetBalanceList
    Call SetFactList
    Call SetDetail
    Call SetExtendInfo
    tbSub.Item(2).Visible = False
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "请刷新清单或重新设置过滤条件"
    
    '初始化消息处理对象模块
    Call zlMsgModuleInit
    
    Call LoadPlugInMnu
End Sub

Private Sub ClearErrInvoice()
'功能：清除操作员上次异常退出时只填了实际票号而没有实际打印的单据的费用记录中的票据号
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "Select 登记时间, 实际票号" & vbNewLine & _
            "From 门诊费用记录 A," & vbNewLine & _
            "     (Select Max(NO) NO From 门诊费用记录 Where 登记时间 > Sysdate - 1 And 操作员姓名 = [1] And 记录性质 = 1) B" & vbNewLine & _
            "Where A.记录性质 = 1 And A.NO = B.NO And A.实际票号 Is Not Null And Not Exists (Select 1 From 票据打印内容 C Where C.NO = B.NO)"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.姓名)
    If rsTmp.RecordCount > 0 Then
        strSql = "Select NO From 门诊费用记录 Where 登记时间 = [1] And 记录性质 = 1 And 实际票号 = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(rsTmp!登记时间), CStr(rsTmp!实际票号))
        For i = 1 To rsTmp.RecordCount
            strSql = "Zl_票据起始号_Update('" & rsTmp!NO & "','',1)"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            rsTmp.MoveNext
        Next
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long
    Dim sngVsc As Single, sngHsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
     mshList.MousePointer = 0
    
    '靠齐控件宽度和高度
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    tbPage.Top = cbrH + 20
    cbrH = cbrH + IIf(tbPage.Visible, tbPage.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height + 10, 0) 'tbPage.Top +
'    sngVsc = (mshDetail.Height + 1000) / (mshDetail.Height + mshList.Height)
    sngVsc = 0.5
    sngHsc = mshFact.Width / (mshFact.Width + mshDetail.Width)
    If mblnMax Then
        sngVsc = 0.5: sngHsc = 0.35
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    If picVsc.Visible = False Then sngVsc = 0
    tbPage.Width = ScaleWidth
    picCons.Top = tbPage.Top + 20
    mshList.Left = Me.ScaleLeft
    mshList.Top = Me.ScaleTop + cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = 0
    picHsc.Width = mshList.Width
    
    mshDetail.Left = 0
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - picHsc.Height - mshList.Height
    mshDetail.Width = (Me.ScaleWidth - IIf(sngHsc = 0, 0, picVsc.Width)) * (1 - sngHsc)
    
    picVsc.Top = mshDetail.Top
    picVsc.Left = mshDetail.Left + mshDetail.Width
    picVsc.Height = mshDetail.Height
    
    mshFact.Top = mshDetail.Top
    mshFact.Left = picVsc.Left + picVsc.Width
    mshFact.Height = mshDetail.Height
    mshFact.Width = Me.ScaleWidth - picVsc.Width - mshDetail.Width
    
    tbSub.Top = mshDetail.Top
    tbSub.Left = picVsc.Left + picVsc.Width
    tbSub.Height = mshDetail.Height
    tbSub.Width = Me.ScaleWidth - picVsc.Width - mshDetail.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
       
    mstrFilter = ""
    mstrFilter2 = ""
    Unload frmChargeFilter
    Unload frmChargeGo
    Call SaveWinState(Me, App.ProductName)
    If Not mobjInExise Is Nothing Then Call mobjInExise.CloseWindows
    Set mobjInExise = Nothing
    Call SaveFlexState(mshList, App.ProductName & "\" & Me.Name & Val(tbPage.Tag))
    Call SaveFlexState(mshDetail, App.ProductName & "\" & Me.Name)
    Call SaveFlexState(mshFact, App.ProductName & "\" & Me.Name)
    Call SaveFlexState(vsSubBalance, App.ProductName & "\" & Me.Name)
    Call SaveFlexState(vsSubInvoice, App.ProductName & "\" & Me.Name)
    '刷新方式
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "刷新方式", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
            Exit For
        End If
    Next
    '保存参数
    If cboDate.ListIndex < 5 Then
        zlDatabase.SetPara "异常单据查询", cboDate.ListIndex, glngSys, mlngModul, InStr(1, mstrPrivs, ";参数设置;") > 0
    End If
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
    '拆卸消息对象
    Call zlMsgModuleUnload
End Sub

Private Sub mnuViewGo_Click()
    If SQLCondition.int门诊标志 <= 1 Then
        frmChargeGo.lbl标识号.Caption = "门诊号"
    ElseIf SQLCondition.int门诊标志 = 2 Then
        frmChargeGo.lbl标识号.Caption = "住院号"
    Else
        frmChargeGo.lbl标识号.Caption = "门诊/住院号"
    End If
    frmChargeGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmChargeGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "正在定位满足条件的单据,按ESC终止 ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '比较条件
        blnFill = True
        With frmChargeGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("首张单据")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("首张发票")) = .txtFact.Text
            End If
            If .cbo操作员.ListIndex > 0 Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("操作员")) = zlStr.NeedName(.cbo操作员.Text)
            End If
            If .txt姓名.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("姓名"))) Like "*" & UCase(.txt姓名.Text) & "*"
            End If
            If .txt标识号.Text <> "" Then
                blnFill = blnFill And (mshList.TextMatrix(i, GetColNum("门诊号")) = .txt标识号.Text Or mshList.TextMatrix(i, GetColNum("住院号")) = .txt标识号.Text)
            End If
        End With
        
        '满足则退出
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
            
            mlngGo = i + 1
            
            stbThis.Panels(2).Text = "找到一条记录"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '按ESC取消
        If mblnGo = False Then
            stbThis.Panels(2).Text = "用户取消定位操作"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "已定位到清单尾部"
    Screen.MousePointer = 0
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.COLS - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mnuEdit_Print_Click()
    Call PrintBill(0)
End Sub

Private Sub mnuEdit_Print_Supplemental_Click()
    Call PrintBill(1)
End Sub

Private Function GetAllNos(ByVal lngBalance As Long) As String
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strNos As String
    strSql = "Select Distinct a.No From 门诊费用记录 A, 病人预交记录 B Where a.结帐id = b.结帐id And b.结算序号 = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalance)
    Do While Not rsTmp.EOF
        strNos = strNos & "," & "'" & rsTmp!NO & "'"
        rsTmp.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetAllNos = strNos
End Function

Private Function GetInvoiceRelatedNos(ByVal strNo As String, Optional ByRef strInvoices As String) As String
    '功能：通过一张单据号获取票据打印的关联单据
    '参数：strNo - 单据号


    '      strInvoices - 回收票据号
    '返回：关联单据号
    '问题号：83602
    Dim strSql As String, rsNos As ADODB.Recordset
    Dim strNos As String, blnNotRule As Boolean '是否按实际打印分配票号的
    Dim strReturnInvoices As String
    
    On Error GoTo ErrHand
    '判断票据规则是否有变
    strSql = "Select 1 From  票据打印明细 " & _
            " Where 票种 = 1 And 是否回收 <> 1 And NO = [1] And Rownum < 2"
    Set rsNos = zlDatabase.OpenSQLRecord(strSql, "判断票号分配方式", strNo)
    
    blnNotRule = rsNos.EOF
    '根据票号分配规则查找关联单据号
    '预定规则分配票号
    If blnNotRule = False Then
        strSql = "" & _
        "   Select Distinct a.No, a.票号" & _
        "   From 票据打印明细 A, 票据打印明细 B, 票据打印明细 C" & _
        "   Where a.No = b.No And a.票种 = b.票种 And a.是否回收 <> 1" & _
        "       And b.票号 = c.票号 And b.票种 = c.票种 And b.是否回收 <> 1" & _
        "       And c.票种 = 1 And c.是否回收 <> 1 And c.No = [1]" & _
        " Order By 票号"
    Else
        strSql = "" & _
        " Select Distinct b.No, a.号码 as 票号" & _
        " From 票据使用明细 A, 票据打印内容 B, 票据打印内容 C" & _
        " Where a.打印id = b.Id And a.票种 = 1 And a.原因<>6" & _
        "       And Not Exists (Select 1 From 票据使用明细 Where 打印id = a.打印id And 号码 = a.号码 And 票种 = a.票种 And 性质 = 2)" & _
        "       And b.Id = c.Id And b.数据性质 = c.数据性质" & _
        "       And c.数据性质 = 1 And c.No = [1]" & _
        " Order By 票号"
    End If

    Set rsNos = zlDatabase.OpenSQLRecord(strSql, "获取重打单据号", strNo)
    strNos = "": strReturnInvoices = ""
    Do While Not rsNos.EOF
        If InStr(strNos & ",", ",'" & NVL(rsNos!NO) & "',") = 0 Then
            strNos = strNos & ",'" & NVL(rsNos!NO) & "'"
        End If
        If InStr(strReturnInvoices & ",", "," & NVL(rsNos!票号) & ",") = 0 Then
            strReturnInvoices = strReturnInvoices & "," & NVL(rsNos!票号)
        End If
        rsNos.MoveNext
    Loop
    strInvoices = IIf(strReturnInvoices = "", "", Mid(strReturnInvoices, 2))
    GetInvoiceRelatedNos = IIf(strNos = "", "'" & strNo & "'", Mid(strNos, 2))
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PrintBill(bytMode As Byte)
'功能：当前收款记录重新打印一张票据
'bytMode=0-重打,1-补打
    Dim strNo As String, strNos As String, blnMediCare As Boolean
    Dim intInsure As Integer, blnVirtualPrint As Boolean, lng结帐ID As Long, lng病人ID As Long, blnDel As Boolean
    Dim strUseType  As String, lngShareUseID As Long, intInvoiceFormat As Integer
    Dim intOldInvoiceFormat As Integer, lngBalance As Long, lngPJ结帐ID As Long
    Dim strReclaimInvoice As String '回收的票据25187
    Dim blnPrintFact As Boolean '是否已经打印过票据
    Dim blnOnePatiPrint As Boolean '按病人一次打印
    Dim strPrintNos As String
    Dim blnLocalNo As Boolean  '指定单据打印
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    If Me.ActiveControl Is mshDetail Then
        '83602,冉俊明,2015-3-31,重打部分单据
        If mshDetail.IsSubtotal(mshDetail.Row) Then
            strNo = mshDetail.TextMatrix(mshDetail.Row + 1, 0)
        Else
            strNo = mshDetail.TextMatrix(mshDetail.Row, 0)
        End If
        blnLocalNo = True
    Else
        If bytMode = 0 Then
            If MsgBox("你确定要重打本次结算的所有单据吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        blnLocalNo = False
    End If
    
    If strNo = "" Then
        MsgBox "当前没有单据可以重打票据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lngBalance = Val(mshList.TextMatrix(mshList.Row, GetColNum("结算序号")))
    blnMediCare = mshList.TextMatrix(mshList.Row, GetColNum("医保")) = "√"
    blnDel = mshList.TextMatrix(mshList.Row, GetColNum("符号")) = "2"   '记录状态为2的，目前是禁用了打印菜单项的
    blnPrintFact = Trim(mshList.TextMatrix(mshList.Row, GetColNum("首张发票"))) <> ""
  
    '按病人打印
    blnOnePatiPrint = False
    If bytMode = 0 Then
        If zlIsOnePatiPrint(strNo, strPrintNos, blnOnePatiPrint) = False Then Exit Sub
    End If
    If blnOnePatiPrint Then
        If blnLocalNo Then
            If MsgBox("当前选择的单据【" & strNo & "】是按病人补打的单据，你是否重新对补打的单据进行重打?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        If CheckBillExistReplenishData(1, 0, strPrintNos) = True Then
            MsgBox "选择的记录进行了医保补充结算，不允许进行重打或补打票据操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If CheckBillExistReplenishData(0, lngBalance) = True Then
            MsgBox "选择的记录进行了医保补充结算，不允许进行重打或补打票据操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '是否已转入后备数据表中
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '此时已转入在线数据表
    End If
    
    If bytMode = 0 Then
        If Not blnOnePatiPrint Then '按病人补打发票的，不进行相关的单据检查(暂不考虑,以后有需求再加上相关限制).
            If Not BillOperCheck(2, mshList.TextMatrix(mshList.Row, GetColNum("操作员")), _
                CDate(mshList.TextMatrix(mshList.Row, GetColNum("收费时间"))), "重打", strNo, , 1) Then Exit Sub
        End If
    Else
        If Trim(mshList.TextMatrix(mshList.Row, GetColNum("首张发票"))) <> "" _
            And Not (Me.ActiveControl Is mshDetail And mnuFeeDetial_Supplemental.Enabled) Then
            MsgBox "当前单据已打印过票据,不能进行补打！", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    If blnOnePatiPrint Then
        '1.按病人打印时，需要全部重打
        strNos = "'" & Replace(strPrintNos, ",", "','") & "'"
        '重打票据时,回收票号
        strReclaimInvoice = zlGetReclaimInvoice(Replace(Replace(strNos, "'", ""), ",", ";"))
    ElseIf Me.ActiveControl Is mshDetail Then
        '2.选择明细列表打印
        If bytMode = 0 Then
            '2.1选择明细列表重打
            '83602,冉俊明,2015-3-31,重打部分单据
            strNos = GetInvoiceRelatedNos(strNo, strReclaimInvoice)
        ElseIf bytMode = 1 Then
            '2.2选择明细列表补打
            If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 Then
                '按实际打印分配票号且分别打印，则只补打当前选择单据
                strNos = "'" & strNo & "'"
            Else
                strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
            End If
        End If
    Else
        '3.选择结算列表打印时，无论补打、重打都全部打
        strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
        If bytMode = 0 Then
            '重打票据时,回收票号
            strReclaimInvoice = zlGetReclaimInvoice(Replace(Replace(strNos, "'", ""), ",", ";"))
        End If
    End If
    
    If blnMediCare Then
        intInsure = ChargeExistInsure(strNo, lng病人ID, lng结帐ID, , blnDel)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
        End If
    End If
    
    lng病人ID = zlGet病人ID(strNo)
    lngPJ结帐ID = zlGet结帐ID(strNo)
    strUseType = zl_GetInvoiceUserType(lng病人ID, 0, intInsure)
    lngShareUseID = zl_GetInvoiceShareID(mlngModul, strUseType)
    intInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, strUseType, intOldInvoiceFormat, blnOnePatiPrint)
    '单据有剩余数量的才可以重打，北京医保，即使退完了也可以重新打印
    If Not blnVirtualPrint Then
        If Not BillExistMoney(strNos, 1) Then
            MsgBox "单据中的项目已经全部退费,不能进行打印！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If (Me.ActiveControl Is mshDetail And bytMode = 0 Or blnOnePatiPrint) And strReclaimInvoice <> "" Then
            '提醒回收票据
            MsgBox "注意:" & vbCrLf & "    你需要回收以下票据：" & vbCrLf & _
                    Replace("    " & strReclaimInvoice, ",", "，"), vbInformation, gstrSysName
    End If
    
    '重打时，仍使用原票据打印格式
    If bytMode = 0 And blnOnePatiPrint = False Then intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, intOldInvoiceFormat, intInvoiceFormat)
    
    Dim strPriceGrade As String
    If gintPriceGradeStartType >= 2 Then
        strPriceGrade = GetPriceGradeFromNos(strNos)
    Else
        strPriceGrade = gstr普通价格等级
    End If
    If RePrintCharge(IIf(bytMode = 0, 1, 2), strNos, Me, 0, strReclaimInvoice, , , _
        intInvoiceFormat, blnVirtualPrint, blnDel, lngShareUseID, strUseType, blnOnePatiPrint, strPriceGrade) Then

        '银医一卡通写卡，85950
        Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, 0, strNos)
        
        '81688:李南春,2015/5/18,评价器
        If Not gobjPlugIn Is Nothing And bytMode = 1 Then
            On Error Resume Next
            Call gobjPlugIn.OutPatiInvoicePrintAfter(lng病人ID, lngPJ结帐ID)
            Err.Clear
        End If
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("当前操作已更改单据清单内容,要刷新清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub SetApplyHeader()
    Dim strHead As String
    Dim i As Long
    
    strHead = "单据号,4,850|申请人,1,800|申请时间,1,1850|申请原因,1,3000|审核状态,4,1000" & _
            "|审核人,1,800|审核时间,1,1850|审核原因,1,3500|记录性质,1,0|结算序号,1,0"
    
    With mshList
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name & tbPage.Selected.Index)
        .RowHeight(0) = 350
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If

        .Redraw = True
    End With

End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    If gbln退费申请模式 And tbPage.Selected.Index = 1 Then
        Call SetApplyHeader
        Exit Sub
    End If
    If tbPage.Selected.Index = 0 Then
        strHead = "医保,4,450|首张单据,4,1000|首张发票,4,1000|姓名,1,1200" & _
            "|性别,4,500|年龄,4,500|门诊号,1,800|住院号,1,800|费别,4,750|应收金额,7,1000|实收金额,7,1000|操作员,4,1200" & _
            "|收费时间,4,1850|结算序号,1,0|符号,1,0|退费符号,1,0|红票已打印,1,0"
    Else
        strHead = "医保,4,450|首张单据,4,1000|姓名,1,1200" & _
            "|性别,4,500|年龄,4,500|门诊号,1,800|住院号,1,800|费别,4,750|应收金额,7,1000|实收金额,7,1000|操作员,4,1200" & _
            "|收费时间,4,1850|结算序号,1,0|符号,1,0|退费符号,1,0|红票已打印,1,0"
    End If
    With mshList
        .Redraw = False
        .ExplorerBar = flexExSort
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name & tbPage.Selected.Index)
        
        i = GetColNum("住院号")
        .ColWidth(i) = IIf(SQLCondition.int门诊标志 = 1, 0, IIf(.ColWidth(i) <= 0, 800, .ColWidth(i)))
        i = GetColNum("门诊号")
        .ColWidth(i) = IIf(SQLCondition.int门诊标志 = 2, 0, IIf(.ColWidth(i) <= 0, 800, .ColWidth(i)))
        
        .RowHeight(0) = 350
        
        '恢复上次行
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .COLS - 1
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i

        .Redraw = True
    End With
End Sub

Private Sub ShowApplyBills(ByVal strFilter As String, ByVal blnSort As Boolean)
'功能：显示退费申请单据
    Dim i As Long, j As Long, k As Long, strSql As String
    
    On Error GoTo errH
    
    '问题号:53953
    If Not blnSort Then
        strSql = "Select a.No As 首张单据, a.申请人, To_Char(a.申请时间, 'YYYY-MM-DD HH24:MI:SS') As 申请时间, a.申请原因," & vbNewLine & _
                "        Decode(Nvl(a.状态, 0), 1, '通过', 2, '拒绝', '申请') As 审核状态, a.审核人, To_Char(a.审核时间, 'YYYY-MM-DD HH24:MI:SS') As 审核时间," & vbNewLine & _
                "        a.审核原因, a.记录性质, b.结算序号" & vbNewLine & _
                " From 病人退费申请 A," & vbNewLine & _
                "      (Select Distinct m.No, Nvl(n.结算序号, n.结帐id) As 结算序号" & vbNewLine & _
                "        From 门诊费用记录 M, 病人预交记录 N" & vbNewLine & _
                "        Where m.结帐id = n.结帐id And m.记录性质 = 1 And m.记录状态 In (1, 3)) B" & vbNewLine & _
                " Where a.No = b.No And (1 = 1 " & strFilter & ")" & vbNewLine & _
                " Order By a.申请时间 Desc, a.审核时间 Desc"
        With SQLCondition
            Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .ApplyName, .ApplyDateB, .ApplyDateE, .AuditName, .AuditDateB, .AuditDateE)
        End With
    End If
    
    mshList.Redraw = False
    mshList.Clear
    mshList.Rows = 2

    mshDetail.Clear
    mshDetail.Rows = 2

    mshFact.Clear
    mshFact.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        
        k = GetColNum("审核状态")
        For i = 1 To mshList.Rows - 1
            If Trim(mshList.TextMatrix(i, k)) = "通过" Then
                '审核通过的用蓝色表示
                mshList.Cell(flexcpForeColor, i, 0, i, mshList.COLS - 1) = &HC00000
            ElseIf Trim(mshList.TextMatrix(i, k)) = "拒绝" Then
                '审核拒绝的用红色表示
                mshList.Cell(flexcpForeColor, i, 0, i, mshList.COLS - 1) = &HC0
            Else
                mshList.Cell(flexcpForeColor, i, 0, i, mshList.COLS - 1) = vbBlack
            End If
        Next
        
        stbThis.Panels(2).Text = "共 " & mrsList.RecordCount & " 次结算"
        Call SetMenu(True)
    End If
    
    Call SetHeader
    Call SetApplyDetail
    Call SetApplyFactList
    
    
        
    '触发AfterRowColChange事件
    mshList.Row = -1
    If mlngCurRow >= mshList.FixedRows And mlngCurRow < mshList.Rows Then
        mshList.Row = mlngCurRow
    Else
        mshList.Row = mshList.FixedRows
    End If
    
    mshList.Redraw = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    mshList.Redraw = True
End Sub

Private Sub ShowBills(Optional ByVal strFilter As String, Optional blnSort As Boolean)
'功能:按条件读取单据列表(过滤功能)
'参数:strFilter=以"AND"开始的条件串
'     blnSort=不重新读取数据,仅重新显示已排序的内容
    Dim i As Long, j As Long, k As Long, l As Long
    Dim strSql As String
    Dim dtStartDate As Date, dtEndDate As Date
    Dim strErrWhere As String
    Dim strWhere As String
    Dim strFeeTable As String
    Dim strTemp As String, strSQL1 As String
    
    On Error GoTo errH
    strErrWhere = "": strWhere = ""
    If gbln退费申请模式 And tbPage.Selected.Index = 1 Then
        Call ShowApplyBills(strFilter, blnSort)
        Exit Sub
    End If
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("正在读取单据列表,请稍候 ...", Me)
        
        If mbln收费 And mbln退费 Then
            '所有费用记录
            strWhere = " Where Mod(记录性质, 10) = 1 And 记录状态 IN([13],[14],[15]) "
        ElseIf mbln收费 Then
            '原始收费记录
            strWhere = " Where 记录性质 = 1 And 记录状态 IN([13],[15]) "
        ElseIf mbln退费 Then
            '退费记录以及重收记录
            strWhere = " Where (Mod(记录性质, 10) = 1 And 记录状态 = [14] Or 记录性质 = 11 And 记录状态 In ([13],[15])) "
        End If
        
        Select Case SQLCondition.int门诊标志
        Case 1 '门诊
            strWhere = strWhere & " And  门诊标志 in (1,4)"
        Case 2 '住院
            strWhere = strWhere & " And  门诊标志 =2"
        Case Else   '所有
        End Select
        
        strErrWhere = ""
        If tbPage.Selected.Index = 2 Or tbPage.Selected.Index = 3 Then
            Select Case cboDate.ListIndex
            Case 0 '今日
                dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
            Case 1 '最近两天
                dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 2 '最近三天
                dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 3  '本周
                dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 4  '本月
                dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-01") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case Else
                dtStartDate = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
            End Select
            lblDateShow.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
            lblDateShow.Caption = lblDateShow.Caption & "~" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
            
            '收费异常记录
            If tbPage.Selected.Index = 2 Then
                strErrWhere = _
                " Where Nvl(费用状态,0) = 1 And 记录性质 = 1 And 记录状态 = 1" & vbNewLine & _
                "       And 登记时间 Between [1] and [2] And 操作员姓名=[3]  " & _
                "       And Not Exists (Select 1" & vbNewLine & _
                "                       From 门诊费用记录 B" & vbNewLine & _
                "                       Where a.No = b.No" & vbNewLine & _
                "                             And Mod(b.记录性质, 10) = 1 And b.记录状态 = 2) "
            Else
                '退费异常记录
                strErrWhere = _
                " Where Nvl(费用状态,0) = 1 And (Mod(记录性质, 10) = 1 And 记录状态 = 2 Or 记录性质 = 11 And 记录状态 = 1)" & vbNewLine & _
                "       And 登记时间 Between [1] and [2] And 操作员姓名=[3]  " & _
                "       And Exists (Select 1" & vbNewLine & _
                "                   From 病人预交记录 B " & vbNewLine & _
                "                   Where a.结帐id=b.结帐ID and Nvl(b.校对标志,0) <> 0) "
            End If
            strWhere = strErrWhere
        Else
            strWhere = strWhere & " And Nvl(费用状态,0) <> 1  "
        End If
        
        If tbPage.Selected.Index = 0 Then strWhere = strWhere & " " & strFilter
        'strFilter中可能因为费别条件而有子查询
        strFeeTable = _
            " Select 结帐ID,记录状态,Max(Decode(n.打印ID, Null, 0, 1)) As 红票已打印" & vbNewLine & _
            " From 门诊费用记录 A,票据打印内容 M, 票据使用明细 N" & _
                strWhere & " And a.No = m.No(+) And m.Id = n.打印id(+) And n.性质(+) = 1 And n.原因(+) = 6" & vbNewLine & _
            " Group By 结帐ID,记录状态"
        
        '注意:一张收费单据可能使用多个票据号,部分退重打的首张单据显示为开始票据号(序号=1)
        '符号:1=该单据未退过,3-该单据被退过,2-该单据为退的记录
        '因为以前一张单据填有多个票据号,暂处理为取开始号(新程序已是只填一个号码)
'        IIf(mbln退费 = False, " And Exists (Select 1 From 门诊费用记录 Where 记录性质=1 And 记录状态 In (1,3) And 结帐id=z.结帐id)", "")
        strSQL1 = _
        "      Select Distinct a.结帐ID,nvl(B.结算序号,a.结帐ID) as 结算序号, " & _
        "               Max(Decode(Nvl(t.险类,0),0,0,1)) as 医保, " & _
        "               Max(decode(A.记录状态,2,1,0)) as 退费标志,Max(a.红票已打印) as 红票已打印" & _
        "      From (" & strFeeTable & ") A, 病人预交记录 B, 保险结算记录 T" & _
        "      Where A.结帐ID=B.结帐ID(+) And  A.结帐id=t.记录ID(+) And t.性质(+)=1 " & mstrInsure & _
        "      Group by a.结帐ID,nvl(B.结算序号,a.结帐ID)"
        strTemp = ""
        If frmChargeFilter.mblnDateMoved Then
            strTemp = Replace(strSQL1, "门诊费用记录", "H门诊费用记录")
            strTemp = Replace(strTemp, "病人预交记录", "H病人预交记录")
            strTemp = Replace(strTemp, "票据打印内容", "H票据打印内容")
            strTemp = Replace(strTemp, "票据使用明细", "H票据使用明细")
        End If
        strSQL1 = " With C_结算信息 As (" & strSQL1 & ")" & IIf(strTemp = "", "", ",C_H结算信息 As (" & strTemp & ")")

        If tbPage.Selected.Index = 0 Then
            strSql = _
            " Select Decode(Max(J.医保),1,'√','') as 医保,Min(A.NO) As 首张单据,Min(A.实际票号) As 首张发票," & _
            "       A.姓名,A.性别,A.年龄,Decode(A.门诊标志,2,'',A.标识号) As 门诊号,Decode(A.门诊标志,2,A.标识号,'') As 住院号," & _
            "       Min(A.费别) as 费别, " & _
            "       To_Char(decode(max(J.退费标志),1,-1,1)*Sum(a.应收金额), '999999999" & gstrDec & "') as 应收金额," & _
            "       To_Char(decode(max(J.退费标志),1,-1,1)*Sum(a.实收金额), '999999999" & gstrDec & "') as 实收金额," & _
            "       A.操作员姓名 as 操作员,To_Char(Min(A.登记时间),'YYYY-MM-DD HH24:MI:SS') as 收费时间,J.结算序号," & _
            "       Max(A.记录状态) as 符号,Min(a.执行状态) As 退费符号,Max(j.红票已打印) as 红票已打印" & _
            " From 门诊费用记录 A,C_结算信息 J,部门表 B,医疗付款方式 C" & _
            " Where A.结帐ID=J.结帐ID And A.开单部门ID=B.ID " & _
            "       And A.付款方式=C.编码(+) " & IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.名称=[16]", "") & _
            "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & _
            " Group by A.姓名,A.性别,A.年龄,A.门诊标志,A.标识号,A.操作员姓名,J.结算序号"
        Else
            strSql = _
            " Select Decode(Max(J.医保),1,'√','') as 医保 ,Min(A.NO) As 首张单据," & _
            "       A.姓名,A.性别,A.年龄,Decode(A.门诊标志,2,'',A.标识号) As 门诊号,Decode(A.门诊标志,2,A.标识号,'') As 住院号," & _
            "       Min(A.费别) as 费别,   " & _
            "       To_Char(decode(max(J.退费标志),1,-1,1)*Sum(a.应收金额), '999999999" & gstrDec & "') as 应收金额," & _
            "       To_Char(decode(max(J.退费标志),1,-1,1)*Sum(a.实收金额), '999999999" & gstrDec & "') as 实收金额," & _
            "       A.操作员姓名 as 操作员,To_Char(Min(A.登记时间),'YYYY-MM-DD HH24:MI:SS') as 收费时间,J.结算序号," & _
            "       Max(A.记录状态) as 符号,Min(a.执行状态) As 退费符号,Max(j.红票已打印) as 红票已打印" & _
            " From 门诊费用记录 A,C_结算信息 J,部门表 B,医疗付款方式 C" & _
            " Where  A.开单部门ID=B.ID And A.付款方式=C.编码(+) And A.结帐ID=J.结帐ID " & IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.名称=[16]", "") & _
            "       And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
            " Group by A.姓名,A.性别,A.年龄,A.门诊标志,A.标识号,A.操作员姓名,J.结算序号 "
        End If
        
        If frmChargeFilter.mblnDateMoved Then
            strSql = strSql & vbNewLine & _
                    " Union All" & vbNewLine & _
                      Replace(Replace(strSql, "门诊费用记录", "H门诊费用记录"), "C_结算信息", "C_H结算信息")
        End If
        
        strSql = "Select * From (" & strSQL1 & strSql & ")  " & _
                " Order By 收费时间 Desc"
                
        With SQLCondition
            If SQLCondition.Default Then SQLCondition.Operator = UserInfo.姓名
            If strErrWhere <> "" Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名)
            Else
                Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, _
                            .DateB, .DateE, .ChargeKind, .NOB, .NOE, .PayKind, .Operator, _
                            .PatientName, .PatientIdentity, .FactB, .FactE, .DeptID _
                                , 1, 2, 3, .PayKindName, .Doctor, .FeeItems, .PatientID)
            End If
        End With
    End If
    
    If tbPage.Selected.Index = 2 Then
        If mrsList.RecordCount = 0 Then
            tbPage.Selected.Caption = "收费异常记录"
        Else
            tbPage.Selected.Caption = "收费异常记录(" & mrsList.RecordCount & ")"
        End If
    End If
    
    If tbPage.Selected.Index = 3 Then
        If mrsList.RecordCount = 0 Then
            tbPage.Selected.Caption = "退费异常记录"
        Else
            tbPage.Selected.Caption = "退费异常记录(" & mrsList.RecordCount & ")"
        End If
    End If
    
    mshList.Clear 1
    mshList.Rows = 2
    
    mshDetail.Clear 1
    mshDetail.Rows = 2
    
    mshFact.Clear
    mshFact.Rows = 2
    
    vsSubBalance.Clear 1
    vsSubBalance.Rows = 2
    
    vsSubInvoice.Clear 1
    vsSubInvoice.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "当前设置没有过滤出任何单据"
        Call SetMenu(False)
    Else
        If tbPage.Selected.Index = 0 Then
         '求实收合计金额,目前此处没有按序号进行汇总,如果有部分退费的情况,实收金额会不正确
            If Not blnSort Then
                strFeeTable = "Select Distinct 结帐id From 门诊费用记录 A" & strWhere
                '一次结算的所有单据
                strFeeTable = _
                    "Select Distinct m.Id, m.记录性质, m.No, m.记录状态, m.序号, " & vbNewLine & _
                    "       m.付款方式, m.病人科室id, m.实收金额, m.开单部门id, m.结帐id" & vbNewLine & _
                    "From 门诊费用记录 M,(" & strFeeTable & ") N" & vbNewLine & _
                    "Where m.结帐id = n.结帐id"

                If frmChargeFilter.mblnDateMoved Then
                    strTemp = Replace(strFeeTable, "门诊费用记录", "H门诊费用记录")
                    strTemp = Replace(strTemp, "病人预交记录", "H病人预交记录")
                    strTemp = Replace(strTemp, "票据打印内容", "H票据打印内容")
                    strTemp = Replace(strTemp, "票据使用明细", "H票据使用明细")
                    strFeeTable = strFeeTable & " Union All " & strTemp
                End If
                
                strSql = "With 门诊费用 As (" & strFeeTable & ")" & vbNewLine & _
                        " Select " & IIf(mbln收费 = False And mbln退费, -1, 1) & "*Sum(a.实收金额) As 金额," & vbNewLine & _
                        "        Count(Distinct Decode(a.记录性质, 11, Null, a.结帐id)) As 单据" & vbNewLine
                If mstrInsure = "" Then
                    strSql = strSql & _
                        " From 门诊费用 A, 部门表 B, 医疗付款方式 C" & vbNewLine & _
                        " Where A.开单部门ID = B.ID And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
                        "       And A.付款方式=C.编码(+) " & IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.名称=[16]", "")
                Else
                    strSql = strSql & _
                        " From (Select a.实收金额, a.记录性质, a.结帐id" & vbNewLine & _
                        "       From 门诊费用 A, 部门表 B, 医疗付款方式 C, 保险结算记录 T" & vbNewLine & _
                        "       Where a.开单部门id = b.Id And (b.站点 = '" & gstrNodeNo & "' Or b.站点 Is Null) And a.付款方式 = c.编码(+)" & vbNewLine & _
                                IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.名称=[16]", "") & _
                        "             And a.结帐id = t.记录id(+) And t.性质(+) = 1 " & mstrInsure & vbNewLine & _
                        "       Group By a.Id, a.实收金额, a.记录性质, a.结帐id) A"
                End If
                strSql = "Select 单据, 金额 From (" & strSql & ")"
                If strErrWhere <> "" Then
                        Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.姓名, 1, 2, 3)
                Else
                    With SQLCondition
                        Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .ChargeKind, .NOB, .NOE, .PayKind, .Operator, .PatientName, Val(.PatientIdentity), .FactB, .FactE, .DeptID _
                                        , 1, 2, 3, .PayKindName, .Doctor, .FeeItems, .PatientID)
                    End With
                End If
            End If
            Set mshList.DataSource = mrsList
            stbThis.Panels(2).Text = "共 " & NVL(mrsTotal!单据, 0) & " 次结算,合计:" & Format(NVL(mrsTotal!金额, 0), gstrDec)
        Else
            Set mshList.DataSource = mrsList
            stbThis.Panels(2).Text = "共 " & mrsList.RecordCount & " 条异常记录"
        End If
        Call SetMenu(True)
    End If
    
    With mshList
        .Redraw = False
        '设置颜色
        .ForeColor = ForeColor
        k = GetColNum("符号")
        l = GetColNum("退费符号")
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, l)) < 0 Then
                '退费记录用红色
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = &HC0
            ElseIf Val(.TextMatrix(i, k)) = 1 And Val(.TextMatrix(i, l)) >= 0 Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
            Else
                '包含退过费的用蓝色
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = &HC00000
            End If
        Next
        
        Call SetHeader
        Call SetDetail
        Call SetFactList
        
        '触发AfterRowColChange事件
        .Row = -1
        If mlngCurRow >= .FixedRows And mlngCurRow < .Rows Then
            .Row = mlngCurRow
        Else
            .Row = .FixedRows
        End If
        
        .Redraw = True
    End With
    If Not blnSort Then Call zlCommFun.StopFlash
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckBalance(Optional ByVal lngBalanceID As Long, Optional ByVal strNo As String) As Boolean
    Dim strSql As String, rsTemp As ADODB.Recordset
    If strNo = "" Then
        strSql = "Select 1 From 病人预交记录 Where 结算序号= [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
        CheckBalance = rsTemp.EOF
    Else
        strSql = "Select 1 From 病人预交记录 A,门诊费用记录 B Where B.NO= [1] And Mod(B.记录性质,10) = 1 And B.结帐id=A.结帐id And Nvl(A.结算序号,0) < 0 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
        CheckBalance = Not rsTemp.EOF
    End If
End Function

Private Sub DetailSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:针对费用列表信息进行分组显示
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With mshDetail
        For i = 0 To .COLS - 1
            If i < .ColIndex("类别") And i > .ColIndex("说明") Then
                .ColHidden(i) = True
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("实收金额"), gstrDec, &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("单据号"), .ColIndex("应收金额"), gstrDec, &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("类别")
        .OutlineCol = .ColIndex("类别")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("类别")) = strTemp

                strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("单据号"))
                If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 Then
                    '83446,如果是按实际打印分配票号,且多张单据收费分别打印,则在单据显示行中增加显示发票号
                    strTemp = strTemp & Space(2) & "发票号:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("发票号"))
                End If
                strTemp = strTemp & Space(2) & "费别:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("费别"))
                strTemp = strTemp & Space(2) & "开单部门:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单科室"))
                If InStr(mstrPrivs, "显示开单人") <> 0 Then
                   strTemp = strTemp & Space(2) & "开单人:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("开单人"))
                End If
                .MergeRow(i) = True
                .MergeCells = flexMergeRestrictRows
                .Cell(flexcpAlignment, i, .ColIndex("类别"), i, .ColIndex("类别")) = 1
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
'                 If Val(.TextMatrix(i + 1, .ColIndex("记录状态"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                 
                For j = 0 To .COLS - 1
                   If j < .ColIndex("应收金额") Then
                       If j >= .ColIndex("类别") Then
                           .Cell(flexcpText, i, j) = strTemp
                           .Cell(flexcpFontBold, i, j) = False
                       End If
                   ElseIf .ColIndex("实收金额") = j Then
                       .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                       .Cell(flexcpFontBold, i, j) = False
                   ElseIf .ColIndex("应收金额") = j Then
                       .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                       .Cell(flexcpFontBold, i, j) = False
                   End If
                Next
            Else
                .TextMatrix(i, .ColIndex("单价")) = Format(Val(.TextMatrix(i, .ColIndex("单价"))), gstrFeePrecisionFmt)
                .TextMatrix(i, .ColIndex("数量")) = Formatex(Val(.TextMatrix(i, .ColIndex("数量"))), 5)
                .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(.TextMatrix(i, .ColIndex("应收金额"))), gstrDec)
                .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(.TextMatrix(i, .ColIndex("实收金额"))), gstrDec)
            End If
        Next
        Call .AutoSize(.ColIndex("类别"))
        Call .AutoSize(.ColIndex("单价"))
        
        For j = 0 To .COLS - 1
            If j < .ColIndex("应收金额") Then
                .MergeCol(j) = True
            Else
                .MergeCol(j) = False
            End If
        Next
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub ShowApplyDetail(Optional strNo As String)
'参数:strDate:单据的登记时间
    Dim i As Long, j As Long, strSql As String, blnDel As Boolean, strDate As String
    
    On Error GoTo errH
    
    If frmChargeFilter.mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNo, , "1")
    Else
        mblnNOMoved = False
    End If
    
    strSql = _
    " Select C.名称 as 类别,Nvl(E.名称,B.名称) as 名称," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 as 商品名,", "") & "B.规格," & _
            IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
    "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
            IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
    "       A.费别,To_Char(Sum(A.标准单价)" & _
            IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
    "       D.名称 as 执行科室,Nvl(A.费用类型,B.费用类型) as 类型," & _
    "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行',9,'异常收费单','第'||ABS(A.执行状态)||'次退费') as 说明," & _
    "       A.记录状态" & _
    " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,收费项目别名 E,药品规格 X" & _
              IIf(gTy_System_Para.byt药品名称显示 = 2, ",收费项目别名 E1", "") & _
    " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
    "       And A.记录性质=1 And A.NO=[1] And A.记录状态" & IIf(blnDel, "=2", " IN(1,3)") & _
            IIf(strDate <> "", " And A.登记时间=[2]", "") & _
    "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            IIf(gTy_System_Para.byt药品名称显示 = 2, "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
    " Group by Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称)," & IIf(gTy_System_Para.byt药品名称显示 = 2, "E1.名称 ,", "") & " B.规格,A.计算单位,A.费别,D.名称," & _
    "       Nvl(A.费用类型,B.费用类型),A.执行状态,A.记录状态,X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1)" & _
    " Order by Nvl(A.价格父号,A.序号)"

    Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, "")
    
    mshDetail.Redraw = False
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
'    mshDetail.ForeColor = IIf(blnDel, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail

    '原始单据退过的为蓝色
    mshDetail.ForeColor = ForeColor
    For i = 1 To mshDetail.Rows - 1
        If Val(mshDetail.TextMatrix(i, mshDetail.COLS - 1)) = 3 Then
            mshDetail.Cell(flexcpForeColor, i, 0, i, mshDetail.COLS - 1) = &HC00000
        End If
    Next

    Call SetApplyDetail
    
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetApplyDetail()
    Dim strHead As String
    Dim i As Long
    
    If glngSys Like "8??" Then
        strHead = "类别,1,750|名称,1,1800" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,2000", "") & "|规格,1,1000|单位,4,500|数量,7,850|费别,1,750|单价,7,850|应收金额,7,850|实收金额,7,850|发药药店,1,850|类型,1,850|说明,1,1000|记录状态,1,0"
    Else
        strHead = "类别,1,750|名称,1,1800" & IIf(gTy_System_Para.byt药品名称显示 = 2, "|商品名,1,2000", "") & "|规格,1,1000|单位,4,500|数量,7,850|费别,1,750|单价,7,850|应收金额,7,850|实收金额,7,850|执行科室,1,850|类型,1,850|说明,1,1000|记录状态,1,0"
    End If
    
    With mshDetail
        .Redraw = False
        
        .COLS = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .ColHidden(i) = False
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        For i = 0 To .COLS - 1
            If .TextMatrix(0, i) = "商品名" Then
                If gTy_System_Para.byt药品名称显示 = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 2000
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            .RowHeight(i) = 300
        Next i
        .RowHeight(0) = 350
        .ColWidth(.COLS - 1) = 0
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        
        'Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Function ReadListData(ByVal bytType As Byte, ByVal lngBalanceID As Long, ByVal blnDel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:读取相关的明细数据
    '返回:数据获取成功返回true,否则返回False
    '编制:刘尔旋
    '日期:2014-06-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMain As ADODB.Recordset
    Dim strWithTable As String, strWhere As String, i As Long
    Dim strTable As String, lngMainRow As Long, strNo As String
    On Error GoTo errHandle
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("首张单据"))
    
    If frmChargeFilter.mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNo, , "1")
    Else
        mblnNOMoved = False
    End If
    
    If bytType = 2 Then
        '10.29以前数据的获取
        strSql = _
            " Select NO As 单据号, Max(发票号) As 发票号, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, " & _
            "       Sum(数量) As 数量, 单价, Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, 执行科室, " & _
            "       Max(类型) As 类型, 说明,划价人,医疗付款方式,Max(摘要), 记录状态" & vbNewLine & _
            " From (Select a.结帐ID,D1.名称 as 开单科室,A.开单人,a.No,Max(a.实际票号) As 发票号,C.名称 as 类别,Nvl(E.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格," & _
                    IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
            "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                    IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
            "       a.费别,To_Char(Sum(A.标准单价)" & _
                    IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
            "       D.名称 as 执行科室,Max(Nvl(A.费用类型,B.费用类型)) as 类型," & _
            "       Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行',9,'异常收费单','第'||ABS(A.执行状态)||'次退费') as 说明," & _
            "       A.记录状态, Nvl(a.价格父号, a.序号) As 序号, A.划价人,F.名称 As 医疗付款方式,Max(摘要) As 摘要 " & _
            " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 D1,收费项目别名 E,收费项目别名 E1,医疗付款方式 F,药品规格 X" & _
            " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
            "       And A.记录性质=1 And A.结帐ID = [1] And A.记录状态" & IIf(blnDel, "=2", " IN(1,3)") & _
            "       And A.收费细目ID=E.收费细目ID(+) And a.开单部门ID=D1.ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 And A.付款方式=F.编码(+) " & _
            " Group by a.结帐id, D1.名称, a.开单人, A.划价人,F.名称,a.费别,a.No,Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称),E1.名称 , B.规格,A.计算单位,D.名称," & _
            "       A.执行状态,A.记录状态,X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1) )" & _
            " Group By NO, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, 单价, 执行科室, 说明, 记录状态,划价人,医疗付款方式 " & _
            " Order By 单据号, 序号"
    Else
        strSql = _
            " Select NO As 单据号, Max(发票号) As 发票号, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, " & _
            "       Sum(数量) As 数量, 单价, Sum(应收金额) As 应收金额, Sum(实收金额) As 实收金额, 执行科室, " & _
            "       Max(类型) As 类型, Max(说明),划价人,医疗付款方式,Max(摘要),Max(状态), Min(退费状态)" & vbNewLine & _
            " From (Select a.结帐ID,D1.名称 as 开单科室,A.开单人,a.No,Max(a.实际票号) As 发票号,C.名称 as 类别,Nvl(E.名称,B.名称) as 名称,E1.名称 as 商品名,B.规格," & _
                    IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 单位," & _
            "       To_Char(Avg(Nvl(A.付数,1)*" & IIf(blnDel, "-1*", "") & "A.数次)" & _
                    IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ",'9999990.00000') as 数量, " & _
            "       a.费别,To_Char(Sum(A.标准单价)" & _
                    IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as 单价, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.应收金额),'9999999" & gstrDec & "') as 应收金额, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.实收金额),'9999999" & gstrDec & "') as 实收金额, " & _
            "       D.名称 as 执行科室,Max(Nvl(A.费用类型,B.费用类型)) as 类型,Max(Decode(A.记录状态,2,'第'||ABS(A.执行状态)||'次退费',Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部分执行'))) As 说明," & _
            "       Max(A.记录状态) As 状态,Min(A.记录状态) As 退费状态, Nvl(a.价格父号, a.序号) As 序号, A.划价人,F.名称 As 医疗付款方式,Max(摘要) As 摘要 " & _
            " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 D1,收费项目别名 E,收费项目别名 E1,医疗付款方式 F,药品规格 X," & _
            "       (Select Distinct 结帐ID From " & IIf(mblnNOMoved, "H", "") & "病人预交记录 Where 结算序号= [1]) F" & _
            " Where A.收费细目ID=B.ID and A.收费类别=C.编码 And A.执行部门ID=D.ID(+) And A.收费细目ID=X.药品ID(+)" & _
            "       And Mod(A.记录性质,10)=1 And A.结帐ID = F.结帐ID " & _
            "       And A.收费细目ID=E.收费细目ID(+) And E.码类(+)=1 And E.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            "       And A.收费细目ID=E1.收费细目ID(+) And A.开单部门ID=D1.ID(+) And E1.码类(+)=1 And E1.性质(+)=3 And A.付款方式=F.编码(+) " & _
            " Group by a.结帐id, D1.名称, A.划价人,F.名称,a.开单人, a.费别,a.No,Nvl(A.价格父号,A.序号),C.名称,Nvl(E.名称,B.名称),E1.名称 , B.规格,A.计算单位,D.名称," & _
            "       X.药品ID,X." & gstr药房单位 & ",Nvl(X." & gstr药房包装 & ",1) )" & _
            " Group By NO, 序号, 开单科室, 开单人, 费别, 类别, 名称, 商品名, 规格, 单位, 单价, 执行科室, 划价人,医疗付款方式 Having Sum(数量) <> 0" & _
            " Order By 单据号, 序号"
    End If
    
    Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    Set mshDetail.DataSource = rsMain
    
    '83446,如果是按实际打印分配票号,且多张单据收费分别打印,则在单据显示行中增加显示发票号
    If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 Then
        Dim rsInVoice As ADODB.Recordset, lngRow As Long
        strSql = "Select b.No, f_List2str(Cast(Collect(Distinct a.号码 Order By a.号码 Asc) As t_Strlist)) As 号码" & vbNewLine & _
                " From 门诊费用记录 C, 票据使用明细 A, 票据打印内容 B" & vbNewLine & _
                " Where b.No = c.No And a.打印id = b.Id And a.票种 = 1 And a.性质 = 1 And a.原因<>6" & vbNewLine & _
                "   And Not Exists(Select 1 From 票据使用明细 Where 号码 = a.号码 And 票种 = 1 And 性质 = 2)" & vbNewLine & _
                "   And c.结帐id " & IIf(bytType = 2, "=[1]", "In (Select 结帐id From 病人预交记录 Where 结算序号 = [1])") & vbNewLine & _
                " Group By b.No"
        Set rsInVoice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
        If rsInVoice.RecordCount > 0 Then
            Do While Not rsInVoice.EOF
                lngRow = mshDetail.FindRow(NVL(rsInVoice!NO), , 0) '单据号
                If lngRow > 0 And lngRow < mshDetail.Rows Then
                    For i = lngRow To mshDetail.Rows - 1
                        If mshDetail.TextMatrix(i, 0) = NVL(rsInVoice!NO) Then
                            mshDetail.TextMatrix(i, 1) = NVL(rsInVoice!号码) '重新设置发票号
                        Else
                            Exit For
                        End If
                    Next
                End If
                rsInVoice.MoveNext
            Loop
        End If
    End If
    
    Call SetDetail
    
    ReadListData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckInsureDetail(ByVal strNo As String, ByVal lngSN As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保重收记录颜色判断
    '返回:未退记录返回True,存在退费记录返回False
    '编制:刘尔旋
    '日期:2014-08-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    strSql = _
        " Select 1" & vbNewLine & _
        " From 门诊费用记录" & vbNewLine & _
        " Where Nvl(价格父号, 序号) = [2] And NO = [1] And Mod(记录性质,10) = 1 Having" & vbNewLine & _
        "  Sum(付数 * 数次) = (Select 付数 * 数次" & vbNewLine & _
        "                       From 门诊费用记录" & vbNewLine & _
        "                       Where Nvl(价格父号, 序号) = [2] And NO = [1] And 记录性质 = 1 And 记录状态 In (1, 3) And Rownum < 2)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, lngSN)
    CheckInsureDetail = Not rsTmp.EOF
End Function

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant
    
    strHead = "单据号,1,0|发票号,1,0|序号,1,0|开单科室,1,0|开单人,1,0|费别,1,0|类别,4,800|名称,1,2000|商品名,1,2000|" & _
            "规格,1,1200|单位,4,500|数量,7,800|单价,7,1000|应收金额,7,1000|实收金额,7,1000|执行科室,4,1000|类型,4,1000|" & _
            "说明,1,1800|划价人,4,750|医疗付款方式,4,1200|摘要,1,1500|记录状态,1,0"
    
    With mshDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            If .TextMatrix(0, i) = "划价人" Then
                If InStr(mstrPrivs, "显示开单人") = 0 Then
                    .ColHidden(i) = True
                End If
            End If
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        If .Rows < 2 Then .Rows = 2
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        .Redraw = True
        If .TextMatrix(1, .ColIndex("单据号")) <> "" Then Call DetailSplitGroup
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                If Val(.TextMatrix(i, .ColIndex("记录状态"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                .RowHeight(i) = 300
                '医保重收记录修正
'                If mshList.TextMatrix(mshList.Row, GetColNum("医保")) <> "" And Val(mshList.TextMatrix(mshList.Row, GetColNum("符号"))) = 3 And Val(mshList.TextMatrix(mshList.Row, GetColNum("退费符号"))) <> 2 Then
'                    If CheckInsureDetail(.TextMatrix(i, .ColIndex("单据号")), Val(.TextMatrix(i, .ColIndex("序号")))) Then
'                        .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                    End If
'                End If
            End If
        Next i
        
        If gTy_System_Para.byt药品名称显示 = 0 Then
            .ColHidden(.ColIndex("名称")) = False
            .ColHidden(.ColIndex("商品名")) = True
        End If
        If gTy_System_Para.byt药品名称显示 = 1 Then
            .ColHidden(.ColIndex("名称")) = True
            .ColHidden(.ColIndex("商品名")) = False
        End If
        If gTy_System_Para.byt药品名称显示 = 2 Then
            .ColHidden(.ColIndex("名称")) = False
            .ColHidden(.ColIndex("商品名")) = False
        End If
    End With
End Sub

Private Sub ShowInvoice(ByVal strNo As String)
    Dim strSql As String, lngBalanceID As Long, blnOld As Boolean
    Dim rsInVoice As ADODB.Recordset
    
    On Error GoTo errH
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("结算序号")))
    blnOld = CheckBalance(lngBalanceID)
    If blnOld Then
        strSql = _
        "Select Distinct b.Id, b.号码 As 票据号," & vbNewLine & _
        "       Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因," & vbNewLine & _
        "       To_Char(b.使用时间, 'MM-DD HH24:MI') As 使用时间, b.使用人" & vbNewLine & _
        "From 票据打印内容 A, 票据使用明细 B, 门诊费用记录 C" & vbNewLine & _
        "Where a.Id = b.打印id And a.No = c.No And a.数据性质 = 1 And b.票种 = 1 And c.结帐id = [1]" & vbNewLine & _
        "Order By 使用时间"
    Else
        strSql = _
        "Select Distinct b.Id, b.号码 As 票据号," & vbNewLine & _
        "       Decode(b.原因, 1, '正常发出', 2, '作废收回', 3, '重打发出', 4, '重打收回', 6, '红票发出') As 使用原因," & vbNewLine & _
        "       To_Char(b.使用时间, 'MM-DD HH24:MI') As 使用时间, b.使用人" & vbNewLine & _
        "From 票据打印内容 A, 票据使用明细 B, 门诊费用记录 C, 病人预交记录 D" & vbNewLine & _
        "Where a.Id = b.打印id And a.No = c.No And c.结帐id = d.结帐id And a.数据性质 = 1 And b.票种 = 1 And d.结算序号 = [1]" & vbNewLine & _
        "Order By 使用时间"
    End If
    If mblnNOMoved Then
        strSql = Replace(strSql, "票据打印内容", "H票据打印内容")
        strSql = Replace(strSql, "票据使用明细", "H票据使用明细")
        strSql = Replace(strSql, "门诊费用记录", "H门诊费用记录")
        strSql = Replace(strSql, "病人预交记录", "H病人预交记录")
    End If
    Set rsInVoice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    Set vsSubInvoice.DataSource = rsInVoice
    Call SetInvoiceList
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowBalanceList(Optional ByVal strNo As String, Optional ByVal blnSort As Boolean)
    Dim strSql As String, i As Long, lngBalanceID As Long
    
    If tbPage.Selected.Index <> 0 And tbPage.Selected.Index <> 1 Then Exit Sub
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("结算序号")))
    '预交款不显示结算号码、摘要、卡号、交易流水号、交易说明
    If CheckBalance(lngBalanceID) = False Then
        strSql = "Select Decode(Mod(a.记录性质,10),1,'冲预存款',Nvl(a.结算方式,'未结金额')) As 结算方式, Sum(a.冲预交) As 冲预交," & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.结算号码)) As 结算号码, " & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.摘要)) As 摘要, " & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.卡号)) As 卡号," & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.交易流水号)) As 交易流水号, " & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.交易说明)) As 交易说明" & _
                " From 病人预交记录 A, (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = [1]) B" & _
                " Where a.结帐id = b.结帐id" & _
                " Group By Decode(Mod(a.记录性质,10),1,'冲预存款',Nvl(a.结算方式,'未结金额'))"
    Else
        strSql = "Select Decode(Mod(a.记录性质,10),1,'冲预存款',Nvl(a.结算方式,'未结金额')) As 结算方式, Sum(a.冲预交) As 冲预交," & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.结算号码)) As 结算号码, " & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.摘要)) As 摘要, " & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.卡号)) As 卡号," & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.交易流水号)) As 交易流水号, " & _
                "        Decode(Mod(Max(a.记录性质),10),1,'',Max(a.交易说明)) As 交易说明" & _
                " From 病人预交记录 A" & _
                " Where a.结帐id = [1]" & _
                " Group By Decode(Mod(a.记录性质,10),1,'冲预存款',Nvl(a.结算方式,'未结金额'))"
    End If
    Set mrsBalance = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    vsSubBalance.Redraw = False
    vsSubBalance.Clear
    vsSubBalance.Rows = 2
    If Not mrsBalance.EOF Then
        Set vsSubBalance.DataSource = mrsBalance
    End If
    Call SetBalanceList
    vsSubBalance.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowFactList(Optional ByVal strNo As String, Optional ByVal blnSort As Boolean)
    Dim strSql As String, i As Long, lngBalanceID As Long
    
    If tbPage.Selected.Index <> 2 And tbPage.Selected.Index <> 3 Then Exit Sub
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("结算序号")))
    If tbPage.Selected.Index = 2 Then
        strSql = "Select Nvl(A.结算方式,'未结金额') As 结算方式,Sum(A.冲预交) As 冲预交,Decode(Nvl(A.校对标志,0),0,'√',2,'√','×') As 标志" & _
                " From 病人预交记录 A" & _
                " Where A.结算序号 = [1]" & _
                " Group By Nvl(A.结算方式,'未结金额'),Nvl(A.校对标志,0)" & _
                " Order By 标志"
    Else
        strSql = "Select Nvl(A.结算方式,'未退金额') As 结算方式,Sum(A.冲预交) As 冲预交,Decode(Nvl(A.校对标志,0),0,'√',2,'√','×') As 标志" & _
                " From 病人预交记录 A" & _
                " Where A.结算序号 = [1]" & _
                " Group By Nvl(A.结算方式,'未退金额'),Nvl(A.校对标志,0)" & _
                " Order By 标志"
    End If
    If CheckBalance(lngBalanceID) Then strSql = Replace(strSql, "结算序号", "结帐ID")
    Set mrsFact = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    
    mshFact.Redraw = False
    mshFact.Clear
    mshFact.Rows = 2
    If Not mrsFact.EOF Then
        Set mshFact.DataSource = mrsFact
    End If
    Call SetFactList
    mshFact.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitTabSub()
    With tbPage
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionTop
        '.PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .InsertItem 0, "收退单据记录", picTemp.hWnd, 0
        .InsertItem 1, "退费申请记录", picTemp.hWnd, 0
        .InsertItem 2, "收费异常记录", picTemp.hWnd, 0
        .InsertItem 3, "退费异常记录", picTemp.hWnd, 0
        .Item(0).Selected = True
    End With
    
    With tbSub
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .InsertItem 1, "票据信息", picSubInvoice.hWnd, 0
        .InsertItem 2, "结算信息", picSubBalance.hWnd, 0
        .InsertItem 3, "结算关联信息", picExtendInfo.hWnd, 0
        .Item(0).Selected = True
    End With
End Sub

Private Sub SetBalanceList()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Long
    Dim varData As Variant
    
    strHead = "结算方式,4,1000|金额,7,1000|结算号码,4,1000|摘要,1,1200|卡号,1,1000|交易流水号,1,1000|交易说明,1,1200"
    
    With vsSubBalance
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("金额")) = Formatex(.TextMatrix(i, .ColIndex("金额")), 6, , , 2)
            .RowHeight(i) = 300
        Next i
        
        Call RestoreFlexState(vsSubBalance, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        .Redraw = True
    End With
End Sub

Private Sub ShowExtendInfo()
    Dim strSql As String, rsTemp As ADODB.Recordset, lngBalanceID As Long
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("结算序号")))
    
    '89448,缴预交款时使用医疗卡，第一次使用该笔预交收费时，不应该有结算关联信息（记录性质<>1）
    If CheckBalance(lngBalanceID) Then
        '无结算序号的数据
        strSql = _
            " Select a.交易id As ID, b.结算方式, c.名称, b.冲预交 As 金额, a.交易项目 As 项目, a.交易内容 As 内容" & vbNewLine & _
            " From " & IIf(mblnNOMoved, "H", "") & "三方结算交易 A, " & IIf(mblnNOMoved, "H", "") & "病人预交记录 B, 医疗卡类别 C" & vbNewLine & _
            " Where b.结帐id = [1] And b.记录性质 <> 1 And a.交易id = b.Id And b.卡类别id = c.Id(+) Order By ID"
    Else
        strSql = _
            " Select a.交易id As ID, b.结算方式, c.名称, b.冲预交 As 金额, a.交易项目 As 项目, a.交易内容 As 内容" & vbNewLine & _
            " From " & IIf(mblnNOMoved, "H", "") & "三方结算交易 A, " & IIf(mblnNOMoved, "H", "") & "病人预交记录 B, 医疗卡类别 C" & vbNewLine & _
            " Where b.结算序号 = [1] And b.记录性质 <> 1 And a.交易id = b.Id And b.卡类别id = c.Id(+) Order By ID"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    Set vsfExtendInfo.DataSource = rsTemp
    If rsTemp.RecordCount = 0 Then
        '没有第三方交易记录时，隐藏分页
        tbSub.Item(2).Visible = False
        If tbSub.Selected.Index = 2 Then tbSub.Item(0).Selected = True
    Else
        tbSub.Item(2).Visible = True
    End If
    Call SetExtendInfo
End Sub

Private Sub SetExtendInfo()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant

    strHead = "ID,1,0|结算方式,1,0|名称,1,0|金额,1,0|项目,1,1200|内容,1,2000"
    
    With vsfExtendInfo
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColKey(i) = .TextMatrix(0, i)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .FixedAlignment(i) = 4
            If .ColKey(i) = "ID" Or .ColKey(i) = "结算方式" Or .ColKey(i) = "名称" Or .ColKey(i) = "金额" Then .ColHidden(i) = True
        Next
        If .Rows < 2 Then .Rows = 2
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        
        .RowHeight(0) = 350
        '.Row = 1: .Col = 0: .ColSel = .COLS - 1
        .Redraw = True
        
        If .TextMatrix(1, 0) = "" Then Exit Sub

        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTNone, .ColIndex("ID"), .ColIndex("项目"), gstrDec, &H8000000F
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("项目")
        .OutlineCol = .ColIndex("项目")
        
        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("项目")) = strTemp

                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("结算方式"))
                 'strTemp = strTemp & Space(1) & .Cell(flexcpTextDisplay, i + 1, .ColIndex("名称"))
                 strTemp = strTemp & "(" & Format(.Cell(flexcpTextDisplay, i + 1, .ColIndex("金额")), gstrDec) & ")"

                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("项目"), i, .ColIndex("项目")) = 1
                 
                 For j = 0 To .COLS - 1
                    If j <= .ColIndex("内容") Then
                        If j >= .ColIndex("项目") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = False
                        End If
                    End If
                 Next
            End If
        Next
        Call .AutoSize(.ColIndex("项目"))
        For j = 0 To .COLS - 1
            .MergeCol(j) = True
        Next
    End With
End Sub

Private Sub ShowApplyFactList(Optional ByVal strNo As String)
    Dim strSql As String, i As Long
    
    If gTy_Module_Para.byt票据分配规则 <> 0 Then
        strSql = _
        " Select distinct B.ID,B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票正常发出') as 使用原因," & _
        " To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
        " From " & IIf(mblnNOMoved, "H", "") & "票据打印明细 A," & _
                IIf(mblnNOMoved, "H", "") & "票据使用明细 B " & _
        " Where A.票种=1 And A.票号=B.号码" & _
        "             And B.票种=1 And A.NO=[1]" & _
        " Order by ID"
        On Error GoTo errH
        Set mrsFact = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
        If mrsFact.RecordCount = 0 Then GoTo ReadOld:
    Else
ReadOld:
        strSql = _
        " Select B.ID, B.号码 as 票据号,Decode(B.原因,1,'正常发出',2,'作废收回',3,'重打发出',4,'重打收回',6,'红票正常发出') as 使用原因," & _
        " To_Char(B.使用时间,'MM-DD HH24:MI') as 使用时间,B.使用人" & _
        " From " & IIf(mblnNOMoved, "H", "") & "票据打印内容 A," & _
                IIf(mblnNOMoved, "H", "") & "票据使用明细 B" & _
        " Where A.数据性质=1 And A.ID=B.打印ID" & _
        " And B.票种=1 And A.NO=[1]" & _
        " Order by ID"
    End If
    On Error GoTo errH
    Set mrsFact = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    
    
    mshFact.Redraw = False
    mshFact.Clear
    mshFact.Rows = 2
    If Not mrsFact.EOF Then
        Set mshFact.DataSource = mrsFact
    End If
    Call SetApplyFactList
    mshFact.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetApplyFactList()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant
    If tbPage.Selected.Index = 0 Then
        tbSub.Visible = True
        mshFact.Visible = False
    Else
        tbSub.Visible = False
        mshFact.Visible = True
    End If
    strHead = "ID,1,0|票据号,1,850|使用原因,1,850|使用时间,1,1080|使用人,1,800"
    With mshFact
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        Call RestoreFlexState(mshFact, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        For i = 1 To .Rows - 1
            .RowHeight(i) = 300
        Next i
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1

        .Redraw = True
    End With
End Sub

Private Sub SetFactList()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Integer
    Dim varData As Variant
    If tbPage.Selected.Index = 0 Then
        tbSub.Visible = True
        mshFact.Visible = False
    Else
        tbSub.Visible = False
        mshFact.Visible = True
    End If
    If tbPage.Selected.Index = 3 Then
        strHead = "结算方式,4,1000|收费金额,7,1000|退费状态,4,1200"
    Else
        strHead = "结算方式,4,1000|收费金额,7,1000|收费状态,4,1200"
    End If
    
    With mshFact
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) Like "*误差*" Then
                strTemp = Val(.TextMatrix(i, 1))
                If InStr(strTemp, ".") = 0 Then
                    strAcc = "0.00"
                Else
                    strTemp = Split(strTemp, ".")(1)
                    strAcc = "0."
                    If Len(strTemp) < 2 Then
                        strAcc = "0.00"
                    Else
                        For j = 1 To Len(strTemp)
                            strAcc = strAcc & "0"
                        Next j
                    End If
                End If
                .TextMatrix(i, 1) = Format(.TextMatrix(i, 1), strAcc)
            Else
                If .TextMatrix(i, 1) <> "" Then .TextMatrix(i, 1) = Format(.TextMatrix(i, 1), "0.00")
            End If
            .RowHeight(i) = 300
        Next i
        Call RestoreFlexState(mshFact, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 350
        
        .Row = 1: .Col = 0: .ColSel = .COLS - 1
        .Redraw = True
    End With
End Sub

Private Sub SetInvoiceList()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant

    strHead = "ID,1,0|票据号,4,1000|使用原因,4,1000|使用时间,4,1200|使用人,1,1000"
    
    With vsSubInvoice
        .Redraw = flexRDNone
        If .Rows = 1 Then .Rows = 2
        .HighLight = flexHighlightWithFocus
        
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            If .ColWidth(i) = 0 Then .ColHidden(i) = True
            .FixedAlignment(i) = 4
        Next
        
        Call RestoreFlexState(vsSubInvoice, App.ProductName & "\" & Me.Name)
        
        .RowHeight(-1) = 300: .RowHeight(0) = 350
        
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Function zlGet病人ID(ByVal strNo As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人ID
    '编制:刘兴洪
    '日期:2011-04-29 17:05:13
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select Max(病人ID) as 病人ID From 门诊费用记录 Where No=[1] and 记录性质=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    zlGet病人ID = Val(NVL(rsTemp!病人ID))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGet结帐ID(ByVal strNo As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取结帐ID
    '编制:李南春
    '日期:2015-09-25 17:05:13
    '问题:81688
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select Min(结帐ID) as 结帐ID From 门诊费用记录 Where No=[1] and 记录性质=1 And 记录状态 In(1,3) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    zlGet结帐ID = Val(NVL(rsTemp!结帐ID))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlMsgModuleInit() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化消息模块
    '入参:lngModule -模块号
    '     strPivs-权限串
    '出参:objMsgModule-返回消息对象
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set mobjMsgModule = New clsMipModule
    Call mobjMsgModule.InitMessage(glngSys, mlngModul, mstrPrivs)
    Call AddMipModule(mobjMsgModule)
    zlMsgModuleInit = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Public Function zlMsgModuleUnload() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拆卸消息模块
    '入参:objMsgModule-消息对象
    '编制:刘兴洪
    '日期:2014-03-11 11:46:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    
    If mobjMsgModule Is Nothing Then Exit Function
    Call mobjMsgModule.CloseMessage
    Call DelMipModule(mobjMsgModule)
    Set mobjMsgModule = Nothing
    zlMsgModuleUnload = False
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub vsfExtendInfo_GotFocus()
    Call SetActiveList(vsfExtendInfo)
End Sub

Private Sub vsSubBalance_GotFocus()
    Call SetActiveList(vsSubBalance)
End Sub

Private Sub vsSubInvoice_GotFocus()
    Call SetActiveList(vsSubInvoice)
End Sub
