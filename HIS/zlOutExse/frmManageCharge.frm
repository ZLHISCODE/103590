VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageCharge 
   AutoRedraw      =   -1  'True
   Caption         =   "�����շѹ���"
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
            Name            =   "����"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Name            =   "����"
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
         Caption         =   "��"
         Height          =   210
         Left            =   3870
         TabIndex        =   10
         Top             =   45
         Width           =   330
      End
      Begin VB.Label lblȱʡ 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡ��ʾ"
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
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�շ�"
               Key             =   "Charge"
               Description     =   "�շ�"
               Object.ToolTipText     =   "�����շѴ���"
               Object.Tag             =   "�շ�"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˷�"
               Key             =   "Del"
               Description     =   "�˷�"
               Object.ToolTipText     =   "�Ե�ǰѡ�е����˷�"
               Object.Tag             =   "�˷�"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "����"
               Key             =   "Cancel"
               Object.ToolTipText     =   "�����쳣����"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Del_"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "�շ�����"
               Object.Tag             =   "����"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��չ"
               Key             =   "Extra"
               Object.ToolTipText     =   "�����չ����"
               Object.Tag             =   "��չ"
               ImageIndex      =   13
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "ExtraItem"
                     Object.Tag             =   "����"
                     Text            =   "����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3722
            MinWidth        =   3722
            Picture         =   "frmManageCharge.frx":FB52
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "�ֽ�㳮(&E)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "�շ�����(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Insure 
         Caption         =   "�������(&I)"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_Charge 
         Caption         =   "�����շ�(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_Simple 
         Caption         =   "���շ�(&S)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEditReCharge 
         Caption         =   "�����շ�(&R)"
      End
      Begin VB.Menu mnuEditCancelBill 
         Caption         =   "�����շ�(&Z)"
      End
      Begin VB.Menu mnuEdit_Charge_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Adjust 
         Caption         =   "����ʱ��(&J)"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEdit_Adjust_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_DelMulti 
         Caption         =   "�����˷�(&U)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSplitMzToZy 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzToZyDel 
         Caption         =   "תסԺ�����˷�(Q)"
      End
      Begin VB.Menu mnuEdit_Del_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "���ĵ���(&V)"
      End
      Begin VB.Menu mnuEdit_View_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "�ش��շ�Ʊ��(&R)"
      End
      Begin VB.Menu mnuEditInvoicePrint 
         Caption         =   "����Ʊ���ش�Ʊ��(&F)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "�����շ�Ʊ��(&B)"
      End
      Begin VB.Menu mnuEditMakeupPrn 
         Caption         =   "�����˲���Ʊ��(&M)"
      End
      Begin VB.Menu mnuEdit_PrintDel 
         Caption         =   "�ش��˷�Ʊ��(&D)"
      End
      Begin VB.Menu mnuEdit_PrintList 
         Caption         =   "��ӡ�շ��嵥(&L)"
      End
      Begin VB.Menu mnuEdit_PrintProve 
         Caption         =   "��ӡ�վ�֤��(&O)"
      End
      Begin VB.Menu mnuEdit_Apply_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Apply 
         Caption         =   "�˷�����(&P)"
      End
      Begin VB.Menu mnuEdit_UnApply 
         Caption         =   "ȡ������(&D)"
      End
      Begin VB.Menu mnuEdit_Audit 
         Caption         =   "�˷����(&T)"
      End
      Begin VB.Menu mnuEdit_RefuseApply 
         Caption         =   "�ܾ�����(&R)"
      End
      Begin VB.Menu mnuEdit_UnAudit 
         Caption         =   "ȡ�����(&D)"
      End
      Begin VB.Menu mnuEditSplitW 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWriteCard 
         Caption         =   "������Ϣд��(&W)"
      End
      Begin VB.Menu mnuEdit_Extra 
         Caption         =   "��չ"
         Begin VB.Menu mnuEdit_ExtraItem 
            Caption         =   "����"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuView_Tlb_1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "ˢ�·�ʽ(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "������Ҫˢ������(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "��������ʾ�Ƿ�ˢ��(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "�������Զ�ˢ������(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
   Begin VB.Menu mnuFeeDetial 
      Caption         =   "������ϸ�Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuFeeDetial_Print 
         Caption         =   "�ش��շ�Ʊ��(&R)"
      End
      Begin VB.Menu mnuFeeDetial_Supplemental 
         Caption         =   "�����շ�Ʊ��(&B)"
      End
      Begin VB.Menu mnuFeeDetial_PrintList 
         Caption         =   "��ӡ�շ��嵥(&L)"
      End
      Begin VB.Menu mnuFeeDetial_PrintProve 
         Caption         =   "��ӡ�վ�֤��(&O)"
      End
   End
End
Attribute VB_Name = "frmManageCharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mrsList As ADODB.Recordset  '�����б�
Private mrsDetail As ADODB.Recordset
Private mrsTotal As ADODB.Recordset
Private mrsFact As ADODB.Recordset
Private mrsBalance As ADODB.Recordset
Private mbln�������� As Boolean
Private mblnFirst As Boolean
Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    ChargeKind As String
    PayKind As String
    PayKindName As String
    PatientID As Long '����ID
    PatientName As String '��������
    PatientIdentity As String '��ʶ��
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
    int�����־ As Integer  '1-����;2-סԺ;3-�����סԺ 33789
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String, mstrFilter2 As String, mstrInsure As String
Private mbln�շ� As Boolean, mbln�˷� As Boolean

Private mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mstrPrivs As String
Private mlngModul As Long
Private mblnNOMoved As Boolean '����ϸʱ��¼��ǰѡ��ĵ����Ƿ����������ݱ���,����������ʱ�������ж�
Private mobjInExise As Object
Private mblnNotClick As Boolean
Private mstrWriteCardTypeIDs As String   '��ǰ���������п����ID
Private mblnPrinting As Boolean
'-----------------------------------------------------------------------------------
'���㿨���
Private mstrPassWord As String
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��
'��Ϣ��ض������
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
    
    strNo = mshDetail.TextMatrix(mshDetail.Row, mshDetail.ColIndex("���ݺ�"))

    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Ե�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If

    '�Ѿ��˹���(����)�ĵ��ݲ��������
    If mshList.TextMatrix(mshList.Row, GetColNum("ҽ��")) <> "��" Then
        If InStr(mstrPrivs, "�����ҽ������") = 0 Then
            MsgBox "��û��Ȩ�޶Է�ҽ�����˵ĵ��ݽ��е���ʱ�������", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If BillExistDelete(strNo, 1) Then
        MsgBox "�õ��ݰ������˷�����,�����������", vbInformation, gstrSysName
        Exit Sub
    End If

    On Error Resume Next
    Err.Clear

    If frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, EM_ED_����, , , , , , mobjMsgModule, strNo) = True Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
        '�ش�Ľ��󣬴�����ǽ���ID
        strNo = .TextMatrix(.Row, GetColNum("�������"))
    End With
    If Val(strNo) = 0 Then Exit Sub
    If CheckBillExistReplenishData(0, Val(strNo)) = True Then
        MsgBox "ѡ����˷Ѽ�¼������ҽ��������㣬����������˷Ѳ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(strNo) < 0 Then
        blnDel = frmClinicDelAndView.ShowMe(Me, EM_MULTI_�˷�, mstrPrivs, Val(strNo))
    Else
        Call DelOldBill
        Exit Sub
    End If
    If blnDel Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
'���ܣ���ǰ�˿��¼���´�ӡһ��Ʊ��
    Dim strNo As String, lngBalance As Long, blnMediCare As Boolean
    Dim intInsure As Integer, blnVirtualPrint As Boolean
    Dim lng����ID As Long, lng����ID As Long, blnDel As Boolean
    Dim strUseType  As String, lngShareUseID As Long, intInvoiceFormat As Integer
    
    Err = 0: On Error GoTo errHandler
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����ش��˷�Ʊ�ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    lngBalance = Val(mshList.TextMatrix(mshList.Row, GetColNum("�������")))

    If CheckBillExistReplenishData(0, lngBalance) = True Then
        MsgBox "ѡ��ļ�¼������ҽ��������㣬����������ش��˷�Ʊ�ݲ�����", vbInformation, gstrSysName
        Exit Sub
    End If

    blnMediCare = mshList.TextMatrix(mshList.Row, GetColNum("ҽ��")) = "��"
    blnDel = mshList.TextMatrix(mshList.Row, GetColNum("����")) = "3"   '��¼״̬Ϊ2�ģ�Ŀǰ�ǽ����˴�ӡ�˵����
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    If Not BillOperCheck(2, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("�շ�ʱ��"))), "�˷ѵ����ش�", strNo, , 1) Then Exit Sub
    
    If blnMediCare Then
        intInsure = ChargeExistInsure(strNo, lng����ID, lng����ID, , blnDel)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
        End If
    End If
    
    lng����ID = zlGet����ID(strNo)
    strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
    lngShareUseID = zl_GetInvoiceShareID(mlngModul, strUseType)
    intInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, strUseType, , , True)
    
    '��ӡ�˷�Ʊ��(��Ʊ)
    If PrintDelCharge(lngBalance, Me, 0, , , intInvoiceFormat, blnVirtualPrint, blnDel, lngShareUseID, strUseType) Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥���ݣ�Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Դ�ӡ�嵥��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
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
        strNos = GetMultiNOs(strNo, , , True)  '�����Ƕ൥���շ��е�һ��
    End If
    
    If glngSys Like "8??" Then
        If ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me, "NO=" & strNos, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
        End If
    Else
        If ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me) Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strNos, "ҩƷ��λ=" & IIf(gblnҩ����λ, 1, 0), 2)
        End If
    End If
End Sub

Private Sub mnuEdit_PrintProve_Click()
    Dim strNo As String, strNos As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Դ�ӡ֤����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
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
        strNos = GetMultiNOs(strNo, , , True) '�����Ƕ൥���շ��е�һ��
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

Private Function ReChargeToErrBillBefore(ByVal strNo As String, ByVal lng������� As Long, Optional blnDel As Boolean = False, _
    Optional bln�˷��쳣 As Boolean = False, Optional ByVal strDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:10.34֮ǰ������ȡ�쳣�ĵ��ݷ���
    '���:
    '   blnDel True-���ϵ���,False-�����շ�
    '   bln�˷��쳣 �Ƿ��˷��쳣����
    '   strDate �շ�ʱ��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 15:41:08
    '˵������blnDel=True And bln�˷��쳣=True��ʾ���ϵ���ʱ�������쳣����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivsReplenish As String
    
    On Error GoTo errHandle
    If bln�˷��쳣 = False Then
        If zlIsCheckExiseSingularity(lng�������) Then
            MsgBox "���쳣�����Ѿ������ϣ���ˣ�������" & IIf(blnDel, "��������", "�����շ�") & "����ˢ�·����б�", vbInformation, gstrSysName
            Exit Function
        End If
        If Not zlIsCheckExistErrBill(lng�������) Then
            MsgBox "���쳣�����Ѿ��������շѣ���ˣ�������" & IIf(blnDel, "��������", "�����շ�") & "����ˢ�·����б�", vbInformation, gstrSysName
            Exit Function
        End If
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInState = IIf(blnDel, 5, 4)
        frmCharge.mstrInNO = strNo
        frmCharge.mbln�˷��쳣 = False
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        If blnDel Then
            '�������շ��쳣��¼�������쳣���д���
            frmCharge.mlngModul = mlngModul
            frmCharge.mstrPrivs = mstrPrivs
            frmCharge.mbytInState = 5
            frmCharge.mstrInNO = strNo
            frmCharge.mbln�˷��쳣 = True
            Set frmCharge.mobjMsgModule = mobjMsgModule
            frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
        Else
            If CheckBillExistReplenishData(0, lng�������) Then
                strPrivsReplenish = ";" & GetPrivFunc(glngSys, 1124) & ";"
                If InStr(strPrivsReplenish, ";�����˷�;") > 0 Then
                    If MsgBox("ѡ��ļ�¼������ҽ�����������Ϊ�쳣��������˷Ѽ�¼���Ƿ���Ըü�¼�����ٴν����˷ѣ�", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        gblnOK = frmReplenishTheBalanceDel.zlShowMe(Me, 1124, strPrivsReplenish, EM_RBDTY_�쳣����, Val(strNo), False, 0, False, strDate)
                    Else
                        Exit Function
                    End If
                Else
                    MsgBox "ѡ��ļ�¼������ҽ�����������Ϊ�쳣��������˷Ѽ�¼���㲻�߱������ü�¼��Ȩ�ޣ�����������˷Ѳ�����", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                gblnOK = frmMultiBills.ShowMe(Me, 2, mstrPrivs, strNo, strDate)
            End If
        End If
    End If
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    Dim blnDel As Boolean, lng������� As Long
    Dim strDelTime As String, bln�˷��쳣 As Boolean
    Dim strNo As String
    
    '����
    If tbPage.Selected.Index <> 2 Then Exit Sub
    With mshList
        strNo = .TextMatrix(.Row, GetColNum("���ŵ���"))
        lng������� = Val(.TextMatrix(.Row, GetColNum("�������")))
        strDelTime = .TextMatrix(.Row, GetColNum("�շ�ʱ��"))
    End With
    If lng������� = 0 Then
        MsgBox "��������Ҫ���ϵ��쳣���ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If lng������� < 0 Then
        Call ReChargeToErrBill(lng�������, True, False, strDelTime)
    Else
        'V10.34.0�汾��ǰ����
        Call ReChargeToErrBillBefore(strNo, lng�������, True, False, strDelTime)
    End If
End Sub

Private Function IsCancelFee(ByVal strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ������쳣��
    '����:���˺�
    '����:2012-03-01 01:04:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    strSql = "Select 1 From ������ü�¼ where ��¼����=1 and NO=[1] and ��¼״̬=3 And RowNum=1 And nvl(����״̬,0)=1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    IsCancelFee = Not rsTemp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mnuEditInvoicePrint_Click()
    '����Ʊ���ش�Ʊ��
    If frmFromInvoiceToPrint.zlRePrintBill(Me, mlngModul, mstrPrivs, 0) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
        If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    '����:����תסԺ�˷�
    '����:36076
    If InStr(1, mstrPrivs, ";תסԺ�����˷�;") = 0 Or mbln�������� Then Exit Sub
    
    If mobjInExise Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjInExise = CreateObject("zl9InExse.clsInExse")
        If Err <> 0 Then
            MsgBox "ע��:" & "    סԺ���ò�������ʧ��,���ܽ����˷�,����ϵͳ��Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        Err = 0
    End If
    If mobjInExise Is Nothing Then Exit Sub
'    CallMzFeeTOZyFeeDel(ByVal frmMain As Object, cnMain As ADODB.Connection, ByVal strDBUser As String, ByVal lngSys As Long, _
'    ByVal lngModule As Long, ByVal strPrivs As String,ByVal int���� As Integer, optional lng����ID as long =0) As Boolean
    If mobjInExise.CallMzFeeTOZyFeeDel(Me, gcnOracle, gstrDBUser, glngSys, mlngModul, mstrPrivs, 1) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    Dim strNo As String, lng������� As Long
    Dim bln�˷��쳣 As Boolean, blnDel As Boolean
    Dim strDelTime As String
    
    Err = 0: On Error GoTo errHandler
    If tbPage.Selected.Index <> 2 And tbPage.Selected.Index <> 3 Then Exit Sub
    
    With mshList
        bln�˷��쳣 = tbPage.Selected.Index = 3
        strNo = .TextMatrix(.Row, GetColNum("���ŵ���"))
        lng������� = Val(.TextMatrix(.Row, GetColNum("�������")))
        strDelTime = .TextMatrix(.Row, GetColNum("�շ�ʱ��"))
    End With
    If strNo = "" Then
        MsgBox "�����������շѻ��˷ѵ��쳣���ݣ�", vbInformation, gstrSysName

        Exit Sub
    End If
    
    If bln�˷��쳣 Then
        '�ж�ָ���ĵ����Ƿ��쳣���շ����ϲ����쳣
        blnDel = zlIsErrChargeCancel(strNo)
    End If
    
    If lng������� < 0 Then
        Call ReChargeToErrBill(lng�������, blnDel, bln�˷��쳣, strDelTime)
    Else
        'V10.34.0�汾��ǰ����
        Call ReChargeToErrBillBefore(strNo, lng�������, blnDel, bln�˷��쳣, strDelTime)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mnuEditWriteCard_Click()
    Dim lngCardTypeID As Long, strExpend As String, lng����ID As Long
    Dim lng������� As Long, strNo As String, lng��¼״̬ As Long
    Dim rsTemp As ADODB.Recordset, strSql As String
    
    '����:��������Ϣд�뿨��
    '����:56615
    If InStr(mstrPrivs, ";������Ϣд��;") = 0 Or mstrWriteCardTypeIDs = "" Then Exit Sub
    If gbln�˷�����ģʽ And tbPage.Selected.Index = 1 Then Exit Sub
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ�������д����", vbInformation, gstrSysName
        Exit Sub
    End If
    '�Ƿ�鿴�˷ѵ���
    lng��¼״̬ = Val(mshList.TextMatrix(mshList.Row, GetColNum("����")))
    
    '1.����δ��ȫִ��(ִ��״̬=0,2)
    strSql = "Select  A.����ID,B.�������" & _
        " From ������ü�¼ A,����Ԥ����¼ B " & vbNewLine & _
        " Where A.����ID=B.����ID and  Nvl(A.���ӱ�־,0)<>9 And A.NO=[1] And A.��¼����=1  " & _
        "       And A.��¼״̬ =[2]  and Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, lng��¼״̬)
    If rsTemp.EOF Then Exit Sub
    
    lng����ID = Val(NVL(rsTemp!����ID))
    lng������� = Val(NVL(rsTemp!�������))
    If lng����ID = 0 Or lng������� = 0 Then Exit Sub
    
    If InStr(1, mstrWriteCardTypeIDs, ",") = 0 Then lngCardTypeID = Val(mstrWriteCardTypeIDs)
    Call gobjSquare.objSquareCard.zlMzInforWriteToCard(Me, mlngModul, lngCardTypeID, lng����ID, lng�������, strExpend)
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
    Dim blnErrPage As Boolean   '�쳣����ҳ��
    
    intFrom = gint������Դ
    blnUnit = gblnҩ����λ
    blnErr = gblnShowErr
    bytInvoice = gTy_Module_Para.bytƱ�ݷ������
        
    With frmSetExpence
        .mlngModul = mlngModul
        .mstrPrivs = mstrPrivs
        .mbytInFun = 0
        .mblnSetDrugStore = False
        .Show 1, Me
    End With
    
 
    '��������ز���,����ˢ��
    If gblnҩ����λ <> blnUnit Or gblnShowErr <> blnErr Or intFrom <> gint������Դ Or bytInvoice <> gTy_Module_Para.bytƱ�ݷ������ Then
        '���˵�:����Ʊ���ش�Ʊ��
        If Not tbPage.Selected Is Nothing Then
            blnErrPage = tbPage.Selected.Index = 2
        Else
            blnErrPage = False
        End If
        mnuEditInvoicePrint.Visible = gTy_Module_Para.bytƱ�ݷ������ <> 0 And Not (InStr(mstrPrivs, ";�ش�Ʊ��;") = 0 Or InStr(mstrPrivs, "�վݴ�ӡ") = 0) And Not blnErrPage
        
        frmChargeGo.lbl��ʶ��.Caption = "��ʶ��"
        If gbln�˷�����ģʽ And tbPage.Selected.Index = 1 Then
            frmChargeFilter.lbl��ʶ��.Caption = "��ʶ��"
        ElseIf gint������Դ = 1 Then
            frmChargeFilter.opt����(0).Value = True
        ElseIf gint������Դ = 2 Then
            frmChargeFilter.opt����(1).Value = True
        End If
        ShowBills IIf(gbln�˷�����ģʽ And tbPage.Selected.Index = 1, mstrFilter2, mstrFilter)
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
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    If strNo <> "" Then
        With mshList
            If gbln�˷�����ģʽ And tbPage.Selected.Index = 1 Then
                Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                        "NO=" & .TextMatrix(.Row, GetColNum("���ŵ���")))
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
                    strNos = GetMultiNOs(strNo, , , True)  '�����Ƕ൥���շ��е�һ��
                End If
                
                strColValue = .TextMatrix(.Row, GetColNum("סԺ��")): strTmp = "סԺ��" '����:33789
                strNos = Replace(strNos, "'", "")
                If strColValue = "" Then strColValue = .TextMatrix(.Row, GetColNum("�����")): strTmp = "�����"
                Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                        "NO=" & strNos, strTmp & "=" & strColValue, _
                        "������=" & .TextMatrix(.Row, GetColNum("������")))
            End If
        End With
    Else
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    End If
End Sub

Private Sub mnuViewFilter_Click()

    If gbln�˷�����ģʽ Then frmChargeFilter.mblnApply = tbPage.Selected.Index = 1
    frmChargeFilter.mstrPrivs = mstrPrivs
    frmChargeFilter.opt����(IIf(gint������Դ = 1, 0, 1)).Value = True
    frmChargeFilter.Show 1, Me
    
    If gblnOK Then
        With frmChargeFilter
            
            If gbln�˷�����ģʽ And tbPage.Selected.Index = 1 Then
                mstrFilter2 = .mstrFilter
                SQLCondition.ApplyName = zlStr.NeedName(.cboApply.Text)
                SQLCondition.ApplyDateB = .dtpApplyB.Value
                SQLCondition.ApplyDateE = .dtpApplyE.Value
                SQLCondition.AuditName = zlStr.NeedName(.cboAudit.Text)
                SQLCondition.AuditDateB = .dtpAuditB.Value
                SQLCondition.AuditDateE = .dtpAuditE.Value
                SQLCondition.int�����־ = IIf(.opt����(0).Value, 0, IIf(.opt����(1).Value, 1, 2)) + 1
            Else
                mstrFilter = .mstrFilter
                mbln�շ� = .chk�շ�.Value = 1
                mbln�˷� = .chk�˷�.Value = 1
                
                'ҽ������
                If .chk��ͨ.Value = 1 And .chkҽ��.Value = 0 Then
                    mstrInsure = " And Nvl(t.����,0) = 0"
                ElseIf .chk��ͨ.Value = 0 And .chkҽ��.Value = 1 Then
                    mstrInsure = " And Nvl(t.����,0) <> 0"
                Else
                    mstrInsure = ""
                End If
                
                SQLCondition.Default = False
                SQLCondition.DateB = .dtpBegin.Value
                SQLCondition.DateE = .dtpEnd.Value
                SQLCondition.int�����־ = IIf(.opt����(0).Value, 0, IIf(.opt����(1).Value, 1, 2)) + 1
                If .cbo�ѱ�.ListIndex > 0 Then SQLCondition.ChargeKind = zlStr.NeedName(.cbo�ѱ�.Text)
                If .cbo���ʽ.ListIndex > 0 Then
                    SQLCondition.PayKind = zlStr.NeedCode(.cbo���ʽ.Text)
                    SQLCondition.PayKindName = zlStr.NeedName(.cbo���ʽ.Text)
                Else
                    SQLCondition.PayKind = ""
                    SQLCondition.PayKindName = ""
                End If
                
                SQLCondition.PatientName = gstrLike & UCase(.txt����.Text) & "%"
                SQLCondition.PatientIdentity = Val(.txt��ʶ��.Text)
                SQLCondition.PatientID = .mlngPrePatient
                SQLCondition.NOB = .txtNOBegin.Text
                SQLCondition.NOE = .txtNoEnd.Text
                SQLCondition.FactB = .txtFactBegin.Text
                SQLCondition.FactE = .txtFactEnd.Text
                SQLCondition.DeptID = .cbo����.ItemData(.cbo����.ListIndex)
                SQLCondition.Doctor = .txt������.Text
                If .cbo����Ա.ListIndex = -1 Then
                    SQLCondition.Operator = UserInfo.����
                Else
                    SQLCondition.Operator = zlStr.NeedName(.cbo����Ա.Text)
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
    Dim blnDo As Boolean '�Ƿ��Ѵ��Ʊ��
    
    If Button <> 2 Then Exit Sub '�����Ҽ��˳�
    If tbPage.Selected Is Nothing Then Exit Sub
    If Not Me.ActiveControl Is mshDetail _
        Or Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 2 _
        Or Not tbPage.Selected.Index = 0 Then Exit Sub
    
    '��ʵ�ʴ�ӡ����Ʊ���Ұ����ݷֱ��ӡʱ���ſ���ѡ��ĳ�ŵ��ݽ��в���
    If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ Then
        If mshDetail.IsSubtotal(mshDetail.Row) Then '������
            If mshDetail.Cell(flexcpTextDisplay, mshDetail.Row + 1, mshDetail.ColIndex("��Ʊ��")) <> "" Then blnDo = True
        Else
            If mshDetail.Cell(flexcpTextDisplay, mshDetail.Row, mshDetail.ColIndex("��Ʊ��")) <> "" Then blnDo = True
        End If
        
        '�����Ҽ��˵�
        If blnDo Then '�ش�
            Call SetPrintMenu(True)
        Else '����
            Call SetPrintMenu(False)
        End If
    Else
        If mnuEdit_Print.Visible = False _
            Or Trim(mshList.TextMatrix(mshList.Row, GetColNum("���ŷ�Ʊ"))) = "" Then
            Call SetPrintMenu(False, False)
            Exit Sub
        Else
            Call SetPrintMenu(True)
        End If
    End If
End Sub

Private Sub SetPrintMenu(Optional ByVal blnEnable As Boolean, Optional ByVal blnPrintVisible As Boolean = True)
    '���ܣ����÷�����ϸ�б��еĲ˵�
    If blnEnable Then
        mnuFeeDetial_Print.Visible = blnEnable And blnPrintVisible: mnuFeeDetial_Print.Enabled = blnEnable
        mnuEdit_Print_Supplemental.Enabled = Not blnEnable
        mnuFeeDetial_Supplemental.Visible = Not blnEnable And blnPrintVisible: mnuFeeDetial_Supplemental.Enabled = Not blnEnable
    Else '�����Ӳ˵����������һ���ɼ����ȵ����ÿɼ��Ǹ�
        mnuEdit_Print_Supplemental.Enabled = Not blnEnable
        mnuFeeDetial_Supplemental.Visible = Not blnEnable And blnPrintVisible: mnuFeeDetial_Supplemental.Enabled = Not blnEnable
        mnuFeeDetial_Print.Visible = blnEnable And blnPrintVisible: mnuFeeDetial_Print.Enabled = blnEnable
    End If
    '�����Ҽ��˵�
    PopupMenu mnuFeeDetial, 2
End Sub

Private Sub mshList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNo As String, strDate As String, blnDel As Boolean, blnAudit As Boolean
    Dim rsTmp As ADODB.Recordset, bytType As Byte
    Dim bln��Ʊ�Ѵ�ӡ As Boolean
    
    If mblnPrinting Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    If NewRow < mshList.FixedRows Then Exit Sub
    
    With mshList
        If tbPage.Selected.Index <> 1 Then
            strNo = .TextMatrix(NewRow, GetColNum("�������"))
            blnDel = Val(.TextMatrix(NewRow, GetColNum("�˷ѷ���"))) < 0
            If strNo = "" Then Exit Sub
            bytType = IIf(CheckBalance(Val(strNo)), 2, 1)
            bln��Ʊ�Ѵ�ӡ = Val(.TextMatrix(NewRow, GetColNum("��Ʊ�Ѵ�ӡ"))) = 1
        Else
            strNo = .TextMatrix(NewRow, GetColNum("���ŵ���"))
            If strNo = "" Then Exit Sub
        End If
    End With
    If mrsList Is Nothing Then Exit Sub
    If mrsList.State = 0 Then Exit Sub
    If mrsList.RecordCount = 0 Then Exit Sub
    
    Call SetMenuCaption
    mlngGo = NewRow
    mlngCurRow = NewRow: mlngTopRow = mshList.TopRow
    
    If gbln�˷�����ģʽ Then
        If tbPage.Selected.Index = 1 Then
            blnAudit = mshList.TextMatrix(NewRow, GetColNum("���״̬")) = "ͨ��" Or mshList.TextMatrix(NewRow, GetColNum("���״̬")) = "�ܾ�"
            
            mnuEdit_UnApply.Enabled = Not blnAudit
            mnuEdit_Audit.Enabled = Not blnAudit
            mnuEdit_RefuseApply.Visible = Not blnAudit And Val(mshList.TextMatrix(NewRow, GetColNum("�������"))) > 0
            mnuEdit_RefuseApply.Enabled = Not blnAudit And Val(mshList.TextMatrix(NewRow, GetColNum("�������"))) > 0
            mnuEdit_UnAudit.Enabled = mshList.TextMatrix(NewRow, GetColNum("���״̬")) = "ͨ��"
            mnuEditWriteCard.Enabled = False
        Else
            strDate = mshList.TextMatrix(NewRow, GetColNum("�Ǽ�ʱ��"))
            blnDel = Val(mshList.TextMatrix(NewRow, GetColNum("����"))) = 2
            
            mnuEdit_Apply.Enabled = Not blnDel
            
            mnuEdit_Adjust.Enabled = Not blnDel
            tbr.Buttons("Del").Enabled = Not blnDel
            mnuEdit_Print.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("���ŷ�Ʊ"))) <> ""
            mnuFeeDetial_Print.Enabled = mnuEdit_Print.Enabled
            mnuEdit_Print_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("���ŷ�Ʊ"))) = ""
            mnuFeeDetial_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("���ŷ�Ʊ"))) = ""
            mnuEdit_PrintProve.Enabled = Not blnDel
            mnuFeeDetial_PrintProve.Enabled = Not blnDel
            mnuEdit_PrintList.Enabled = Not blnDel
            mnuFeeDetial_PrintList.Enabled = Not blnDel
            mnuEdit_PrintDel.Enabled = blnDel And IIf(bln��Ʊ�Ѵ�ӡ, InStr(mstrPrivs, ";�ش�Ʊ��;") > 0, InStr(mstrPrivs, ";����Ʊ��;") > 0)
            mnuEdit_PrintDel.Caption = (IIf(bln��Ʊ�Ѵ�ӡ, "�ش��˷�Ʊ��(&D)", "�����˷�Ʊ��(&B)"))
            mnuEditWriteCard.Enabled = strNo <> ""
        End If
    Else
        strDate = mshList.TextMatrix(NewRow, GetColNum("�Ǽ�ʱ��"))
        blnDel = Val(mshList.TextMatrix(NewRow, GetColNum("�˷ѷ���"))) < 0
        
        mnuEdit_Adjust.Enabled = Not blnDel

        tbr.Buttons("Del").Enabled = Not blnDel
        
        mnuEdit_Print.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("���ŷ�Ʊ"))) <> ""
        mnuFeeDetial_Print.Enabled = mnuEdit_Print.Enabled
        mnuEdit_Print_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("���ŷ�Ʊ"))) = ""
        mnuFeeDetial_Supplemental.Enabled = Not blnDel And Trim(mshList.TextMatrix(NewRow, GetColNum("���ŷ�Ʊ"))) = ""
        mnuEdit_PrintProve.Enabled = Not blnDel
        mnuFeeDetial_PrintProve.Enabled = Not blnDel
        mnuEdit_PrintList.Enabled = Not blnDel
        mnuFeeDetial_PrintList.Enabled = Not blnDel
        mnuEdit_PrintDel.Enabled = blnDel And IIf(bln��Ʊ�Ѵ�ӡ, InStr(mstrPrivs, ";�ش�Ʊ��;") > 0, InStr(mstrPrivs, ";����Ʊ��;") > 0)
        mnuEdit_PrintDel.Caption = (IIf(bln��Ʊ�Ѵ�ӡ, "�ش��˷�Ʊ��(&D)", "�����˷�Ʊ��(&B)"))
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
    
    '����AfterRowColChange�¼�
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
    '����:���ò˵���Caption����
    '����:���˺�
    '����:2011-09-04 11:40:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, blnDel As Boolean
    If tbPage.Selected Is Nothing Then Exit Sub
    If Not tbPage.Selected.Index = 2 Then Exit Sub
    With mshList
        If .Row <= 0 Then Exit Sub
        strNo = .TextMatrix(.Row, GetColNum("���ŵ���"))
        blnDel = Val(.TextMatrix(.Row, GetColNum("����"))) = 2
        If strNo = "" Then Exit Sub
    End With
    mnuEditCancelBill.Caption = IIf(blnDel, "�����˷�(&Z)", "�����շ�(&Z)")
    tbr.Buttons("Cancel").Caption = IIf(blnDel, "�˷�", "����")
    tbr.Buttons("Cancel").ToolTipText = IIf(blnDel, "�����˷��쳣����:" & strNo, "�����쳣����:" & strNo)
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
    If Me.ActiveControl Is mshDetail Then Exit Sub 'ʹ�÷�����ϸ�б�ĵ����˵�
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
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
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    lngBalance = mshList.TextMatrix(mshList.Row, GetColNum("�������"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����˷ѣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If CheckBalance(lngBalance) = False Then
        blnOneCard = GetOneCard.RecordCount > 0
        If frmMultiBills.ShowMe(gfrmMain, 1, mstrPrivs, strNo, "", , , blnOneCard) Then
            If mnuViewRefeshOptionItem(1).Checked Then
                If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mnuViewReFlash_Click
                End If
            ElseIf mnuViewRefeshOptionItem(2).Checked Then
                mnuViewReFlash_Click
            End If
        End If
        Exit Sub
    End If
    
    'Ȩ�޼��
    If Not BillOperCheck(2, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), _
        CDate(mshList.TextMatrix(mshList.Row, GetColNum("�շ�ʱ��"))), "�˷�", strNo, , 1) Then Exit Sub
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    intInsure = ChargeExistInsure(strNo)
    
    If zlCheckIsInvoiceListPrinted(strNo, mblnNOMoved) Then
        '����ӡ��ϸ���д�ӡʱ,�����շѴ������ж൥�ݴ���
        strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
    Else
        strTempNos = GetMultiNOs(strNo, , mblnNOMoved, True)
        strNos = GetMultiNOs(strNo, , mblnNOMoved, False)
        If InStr(strTempNos, ",") > 0 And InStr(strNos, ",") = 0 Then
            '�϶��ǰ����ݷֱ��ӡ��
            'Ҫ�൥���˷�,�ͱ���������������
            '1.ҽ���൥�ݱ���ȫ��ʱ,���밴������Ž����˷�
            '2.�����˻�ȫ��ʱ,���밴������Ž����˷�
            If intInsure <> 0 Then
                If gclsInsure.GetCapability(support�൥���շѱ���ȫ��, , intInsure) Then
                    strNos = strTempNos
                End If
            ElseIf zlIsExistsSquareCard(strTempNos, True) Then
                '���һ��ͨ���㲿���Ƿ����ȫ�˵�
                strNos = strTempNos
            End If
        End If
    End If
    
    If zlCheckIsMzToZY(strNo, 1) Then
        MsgBox "ע��:" & vbCrLf & _
                      "    �õ����Ѿ����������תסԺ���� " & vbCrLf & _
                      "    ���Ѿ�������������תסԺ����,�������˷�", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    'ҽ������ƥ���ж�(ȷ��ʱ�����ظ��ж�һ��,��Ϊ��Ҫ��ȡ����ҽ������)
    If intInsure > 0 Then
        '�����˷�Ȩ�޼��
        If InStr(mstrPrivs, "�����շ�") = 0 Then
            MsgBox "��û��Ȩ�޶�ҽ�����˵ĵ����˷ѣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(strNos, ",") > 0 Then
            If gclsInsure.GetCapability(support�൥���շѱ���ȫ��, , intInsure) Then
                MsgBox "��ǰҽ�������������һ�ŵ����˷ѣ�", vbInformation, gstrSysName
                Call mnuEdit_DelMulti_Click
                Exit Sub
            End If
        End If
    Else
        If InStr(mstrPrivs, "�����ҽ������") = 0 Then
            MsgBox "��û��Ȩ�޶Է�ҽ�����˽����˷Ѳ�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
        
    If gblnMultiBalance And InStr(strNos, ",") > 0 Then
        If CheckSingleBalance(strNos) = False Then
            MsgBox "���ŵ���ʹ�ö��ֽ��㷽ʽģʽ�²����������һ�ŵ����˷ѣ�", vbInformation, gstrSysName
            Call mnuEdit_DelMulti_Click
            Exit Sub
        End If
    End If
    
    '���˺�:�൥���˷ѣ�Ҫ����Ǵ��ڽ��㿨�����ڽ��㿨��ֻ��ȫ��
    If UBound(Split(strNos, ",")) > 0 Then
        '�൥���˷�
        If zlIsExistsSquareCard(strNos) = True Then
            '���ö൥���˷�
            'If MsgBox("ע��:" & vbCrLf & "    ���ŵ��ݴ��ڽ��㿨����,���ܶ����е�һ�ŵ����˷�,�Ƿ���ö൥���˷�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Call mnuEdit_DelMulti_Click
            Exit Sub
        End If
    End If
    
            
            
    '�Ƿ���ִ��
    i = BillCanDelete(strNo, 1, blnHaveExe, , blnFlagPrint)
    If i <> 0 Then
        Select Case i
            Case 1 '�õ��ݲ�����
                MsgBox "ָ���ĵ��ݲ����ڣ�", vbInformation, gstrSysName
            Case 2 '�Ѿ�ȫ����ȫִ��
                '�������˷��Զ���ҩ
                MsgBox "�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�", vbInformation, gstrSysName
            Case 3 'δ��ȫִ�в���ʣ������Ϊ0
                MsgBox "�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�", vbInformation, gstrSysName
        End Select
        Exit Sub
    ElseIf intInsure > 0 And blnHaveExe Then
        MsgBox "��ҽ���շѵ����а����Ѿ�ִ�е���Ŀ,�����˷ѣ�", vbInformation, gstrSysName
        Exit Sub
    ElseIf intInsure = 0 And blnHaveExe Then
        If GetOneCardBalance(Val(mshList.TextMatrix(mshList.Row, GetColNum("����ID")))).RecordCount > 0 Then
            MsgBox "�õ������ڴ�����ִ�е���Ŀ,ʹ����һ��ͨ����,���ܽ��в����˷ѣ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    If blnHaveExe Then
        MsgBox "ע��:�õ������ڴ�����ִ�е���Ŀ����ǰ��ִ�е��ǲ����˷ѡ�", vbInformation, gstrSysName
    End If
    If blnFlagPrint Then
        If MsgBox("ע��:����ҽ���������Ѵ�ӡ���Ƿ�����˷ѣ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
            
    On Error GoTo errH
    If Not isSimple(strNo) Then
        On Error Resume Next    '���������ڲ�����ʱ������һ���˳���unload����
        frmCharge.mlngModul = mlngModul
        frmCharge.mstrPrivs = mstrPrivs
        frmCharge.mbytInFun = 0
        frmCharge.mbytInState = 3
        frmCharge.mstrInNO = strNo
        Set frmCharge.mobjMsgModule = mobjMsgModule
        frmCharge.Show IIf(gfrmMain Is Nothing, 0, 1), Me
    Else
        If MsgBox("ȷʵҪ���õ����˷���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
        If gblnBillPrint Then
            If gobjBillPrint.zlEraseBill("'" & strNo & "'", 0) = False Then Exit Sub
        End If
        
        '���շѲ�֧��ҽ��,Ҳ���ṩ�����˷�
        strSql = "zl_������շ�_DELETE('" & strNo & "','" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, "'" & strNo & "'")
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        
        gblnOK = True
    End If
    
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("�������"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))

    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_�˷�����, mstrPrivs, Val(strBalance), , , mblnNOMoved) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "�����˷ѳɹ�!"
        End If
    Else
        strNos = Replace(GetMultiNOs(strNo, , , True), "'", "")
        arrTmp = Split(strNos, ",")
        '�൥��һ���շѵ���ʷ���ݱ���һ�������ȡ�����룬�Լ��ܾ����룬��Ϊ�ڹ�������ֻ��ѡ�����ŵ��ݣ�
        '��������ŵ��ݽ������룬�еĵ��ݾ�ѡ�񲻵����޷���������
'        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNO)

'        If blnTogetherDo Then
        If UBound(arrTmp) > 0 Then
            If MsgBox("����[" & strNos & "]����һ�������˷ѣ���ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        '71917,Ƚ����,2014-4-17,�ڲ����˷�����ʱ�����˷�����ԭ��
        If Not frmInputBox.InputBox(Me, "����ԭ��", "����������ԭ��", 100, 2, True, False, strReason, False) Then Exit Sub
        
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                '71917,Ƚ����,2014-4-17,�ڲ����˷�����ʱ�����˷�����ԭ��
                strSql = "Zl_�����˷�����_Apply(0,'" & arrTmp(i) & "',1,'" & UserInfo.���� & "'," & strDate & ",'" & strReason & "')"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "�Ե���[" & strNos & "]�����˷ѳɹ�!"
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
    
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("�������"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    strDate = mshList.TextMatrix(mshList.Row, GetColNum("����ʱ��"))
    strApplicant = mshList.TextMatrix(mshList.Row, GetColNum("������"))

    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_ȡ������, mstrPrivs, Val(strBalance), , , mblnNOMoved, , strDate) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "ȡ���˷�����ɹ�!"
        End If
    Else
        If MsgBox("ȷʵҪȡ������[" & strNo & "]���˷�������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        If InStr(1, mstrPrivs, "���в���Ա") = 0 Then
            If mshList.TextMatrix(mshList.Row, GetColNum("������")) <> UserInfo.���� Then
                MsgBox "��û��Ȩ��ȡ�����˵��˷����뵥��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
            
        strNos = Replace(GetMultiNOs(strNo, , , True), "'", "")
        If CheckBalance(, strNo) Then strNos = strNo
        arrTmp = Split(strNos, ",")
        '�൥��һ���շѵ���ʷ���ݱ���һ�������ȡ�����룬�Լ��ܾ����룬��Ϊ�ڹ�������ֻ��ѡ�����ŵ��ݣ�
        '��������ŵ��ݽ������룬�еĵ��ݾ�ѡ�񲻵����޷���������
'        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNO)

'        If blnTogetherDo Then
        If UBound(arrTmp) > 0 Then
            If MsgBox("����[" & strNos & "]����һ��ȡ���˷����룬��ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                '71917,Ƚ����,2014-4-17,�ڲ����˷�����ʱ�����˷�����ԭ��
                strSql = "Zl_�����˷�����_Apply(1,'" & arrTmp(i) & "',1,'" & strApplicant & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'))"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "�Ե���[" & strNos & "]ȡ���˷�����ɹ�!"
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
    
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("�������"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    strApplyDate = mshList.TextMatrix(mshList.Row, GetColNum("����ʱ��"))
    
    If InStr(1, mstrPrivs, "���в���Ա") = 0 Then
        If mshList.TextMatrix(mshList.Row, GetColNum("�����")) <> "" And mshList.TextMatrix(mshList.Row, GetColNum("�����")) <> UserInfo.���� Then
            MsgBox "��û��ȡ�����������˵��˷����뵥��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_ȡ�����, mstrPrivs, Val(strBalance), , , mblnNOMoved, , strApplyDate) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "ȡ����˳ɹ�!"
        End If
    Else
        If MsgBox("ȷʵҪȡ������[" & strNo & "]�������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
    
        strNos = Replace(GetMultiNOs(strNo), "'", "")
        arrTmp = Split(strNos, ",")
        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNo)
        
        If blnTogetherDo Then
            If MsgBox("����[" & strNos & "]����һ��ȡ����ˣ���ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            Else
                For i = 0 To UBound(arrTmp)
                    If BillExistDelete(arrTmp(i), 1) Then
                        MsgBox "����[" & arrTmp(i) & "]���˷ѣ�����ȡ����ˡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next
            End If
        Else
            If BillExistDelete(strNo, 1) Then
                MsgBox "����[" & strNo & "]���˷ѣ�����ȡ����ˡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                strSql = "Zl_�����˷�����_Audit('" & arrTmp(i) & "',1,To_Date('" & strApplyDate & "','YYYY-MM-DD HH24:MI:SS')," & _
                     "NULL,NULL,NULL,3)"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "�ѶԵ���[" & strNos & "]ȡ����ˣ�"
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
    
    strBalance = mshList.TextMatrix(mshList.Row, GetColNum("�������"))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    strApplyDate = mshList.TextMatrix(mshList.Row, GetColNum("����ʱ��"))

    If Val(strBalance) < 0 Then
        If frmClinicDelAndView.ShowMe(Me, EM_MULTI_�˷����, mstrPrivs, Val(strBalance), , , mblnNOMoved, , strApplyDate) = True Then
            mblnNotClick = True
            mnuViewReFlash_Click
            mblnNotClick = False
            stbThis.Panels(2).Text = "��������ˣ�"
        End If
    Else
        If MsgBox("ȷʵҪ������[" & strNo & "]�����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            Exit Sub
        End If
        
        strNos = Replace(GetMultiNOs(strNo), "'", "")
        arrTmp = Split(strNos, ",")
        If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNo)
        
        If blnTogetherDo Then
            If MsgBox("����[" & strNos & "]����һ����ˣ���ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        
        '����Ƿ������ִ�е���Ŀ
        For i = 0 To UBound(arrTmp)
            Call BillCanDelete(arrTmp(i), 1, blnHaveExe)
            If blnHaveExe Then
                strInfos = strInfos & "," & arrTmp(i)
            End If
        Next
        If strInfos <> "" Then
            strInfos = Mid(strInfos, 2)
            If MsgBox("����[" & strInfos & "]�д�����ִ�е���Ŀ����ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
'        If Not frmInputBox.InputBox(Me, "���ԭ��", "���������ԭ��", 100, 2, True, False, strReason, False) Then Exit Sub

        strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        
        On Error GoTo errH
        gcnOracle.BeginTrans: blnTrans = True
            For i = 0 To UBound(arrTmp)
                strSql = "Zl_�����˷�����_Audit('" & arrTmp(i) & "',1,To_Date('" & strApplyDate & "','YYYY-MM-DD HH24:MI:SS')," & _
                    "'" & UserInfo.���� & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),'" & strReason & "',1)"
                Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mblnNotClick = True
        mnuViewReFlash_Click
        mblnNotClick = False
        stbThis.Panels(2).Text = "����[" & strNos & "]����ˣ�"
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
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    strApplyDate = mshList.TextMatrix(mshList.Row, GetColNum("����ʱ��"))
    
    If InStr(1, mstrPrivs, "���в���Ա") = 0 Then
        If mshList.TextMatrix(mshList.Row, GetColNum("�����")) <> "" And mshList.TextMatrix(mshList.Row, GetColNum("�����")) <> UserInfo.���� Then
            MsgBox "��û��Ȩ�޾ܾ�������˵��˷����뵥��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    strNos = Replace(GetMultiNOs(strNo, , , True), "'", "")
    arrTmp = Split(strNos, ",")
    '�൥��һ���շѵ���ʷ���ݱ���һ�������ȡ�����룬�Լ��ܾ����룬��Ϊ�ڹ�������ֻ��ѡ�����ŵ��ݣ�
    '��������ŵ��ݽ������룬�еĵ��ݾ�ѡ�񲻵����޷���������
'    If UBound(arrTmp) > 0 Then blnTogetherDo = CheckTogetherDo(strNo)

'    If blnTogetherDo Then
    If UBound(arrTmp) > 0 Then
        If MsgBox("����[" & strNos & "]����һ��ܾ����룬��ȷ��Ҫ������", vbQuestion + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
            Exit Sub
        Else
            For i = 0 To UBound(arrTmp)
                If BillExistDelete(arrTmp(i), 1) Then
                    MsgBox "����[" & arrTmp(i) & "]���˷ѣ����ܾܾ����롣", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End If
    End If
    If Not frmInputBox.InputBox(Me, "�ܾ�ԭ��", "������ܾ�ԭ��", 100, 2, False, False, strReason, False) Then Exit Sub
    
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrTmp)
            strSql = "Zl_�����˷�����_Audit('" & arrTmp(i) & "',1,To_Date('" & strApplyDate & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                 UserInfo.���� & "',To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),'" & strReason & "',2)"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Call mnuViewReFlash_Click
    stbThis.Panels(2).Text = "�ѶԵ���[" & strNos & "]�ܾ����룡"
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
    '����:����Ƿ�����쳣���շѵ���
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 15:27:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtStartDate As Date, dtEndDate As Date, strErrWhere As String
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim blnDel As Boolean, strLast As String
    
    If InStr(mstrPrivs, ";�����շ�;") = 0 Then Exit Function
    
    On Error GoTo errHandle
    Select Case cboDate.ListIndex
       Case 0 '����
           dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
       Case 1 '�������
           dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case 2 '�������
           dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case 3  '���һ��
           dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case 4  '����
           dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-01") & " 00:00:00")
           dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
       Case Else
           dtStartDate = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
           dtEndDate = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
       End Select
       lblDateShow.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
       lblDateShow.Caption = lblDateShow.Caption & "��" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
       
    '�շ��쳣��¼
    strSql = _
    " Select Count(distinct nvl(B.�������,B.����ID)) as ����" & vbNewLine & _
    " From ������ü�¼ A,����Ԥ����¼ B" & _
    " Where A.����ID=B.����ID And Nvl(A.����״̬, 0) = 1 And A.��¼���� = 1  " & _
    "       And A.��¼״̬ = 1 And A.�Ǽ�ʱ�� Between [1] And [2] " & _
    "       And A.����Ա���� = [3] " & vbNewLine & _
    "       And Not Exists (Select 1 From ������ü�¼ Q Where a.No = Q.No And Mod(Q.��¼����, 10) = 1 And Q.��¼״̬ = 2) "
        

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
    CheckErrBill = rsTemp!���� <> 0
    If rsTemp!���� <> 0 Then
        tbPage.Item(2).Caption = "�շ��쳣��¼(" & rsTemp!���� & ")"
        If tbPage.Selected.Index <> 2 Then
            If MsgBox("�����շ��쳣��¼,�Ƿ����շ��쳣��¼?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tbPage.Item(2).Selected = True
                Call mnuEditReCharge_Click
                Exit Function
            End If
        End If
    Else
        tbPage.Item(2).Caption = "�շ��쳣��¼"
    End If
    
    '�˷��쳣��¼
    strSql = "" & _
        " Select /*+ Rule*/ Count(distinct nvl(B.�������,B.����ID)) as ���� " & vbNewLine & _
        " From ������ü�¼ A,����Ԥ����¼ B" & vbNewLine & _
        " Where Nvl(A.����״̬, 0) = 1 And Mod(A.��¼����, 10) = 1 " & _
        "       And A.��¼״̬ = 2 And A.�Ǽ�ʱ�� Between [1] And [2] And A.����Ա���� = [3] " & _
        "       And Exists (Select 1 From ����Ԥ����¼ Q Where A.����id = Q.����id And Nvl(Q.У�Ա�־, 0) <> 0)  " & _
        "       And A.����id =B.����id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
    
    CheckErrBill = rsTemp!���� <> 0
    If rsTemp!���� <> 0 Then
        tbPage.Item(3).Caption = "�˷��쳣��¼(" & rsTemp!���� & ")"
        If tbPage.Selected.Index <> 3 Then
            If MsgBox("�����˷��쳣��¼,�Ƿ����˷��쳣��¼?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                tbPage.Item(3).Selected = True
                Call mnuEditReCharge_Click
                Exit Function
            End If
        End If
    Else
        tbPage.Item(3).Caption = "�˷��쳣��¼"
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ReChargeToErrBill(ByVal lng������� As Long, Optional ByVal blnDel As Boolean = False, _
    Optional ByVal bln�˷��쳣 As Boolean = False, Optional ByVal strDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ȡ�쳣�ĵ��ݷ���
    '���:
    '   blnDel True-���ϵ���,False-�����շ�
    '   bln�˷��쳣 �Ƿ��˷��쳣����
    '   strDate �շ�ʱ��
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 15:41:08
    '˵������blnDel=True And bln�˷��쳣=True��ʾ���ϵ���ʱ�������쳣����
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrivsReplenish As String
    Dim blnOK As Boolean
    
    On Error GoTo errHandle
    If bln�˷��쳣 = False Then
        If zlIsCheckExiseSingularity(lng�������) Then
            MsgBox "���쳣�����Ѿ������ϣ���ˣ�������" & IIf(blnDel, "��������", "�����շ�") & "����ˢ�·����б�", vbInformation, gstrSysName
            Exit Function
        End If
        If Not zlIsCheckExistErrBill(lng�������) Then
            MsgBox "���쳣�����Ѿ��������շѣ���ˣ�������" & IIf(blnDel, "��������", "�����շ�") & "����ˢ�·����б�", vbInformation, gstrSysName
            Exit Function
        End If
    
        blnOK = frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, IIf(blnDel, EM_ED_�쳣����, EM_ED_�쳣����), , lng�������, mblnNOMoved, , , mobjMsgModule)
    Else
        If blnDel Then
            '�������շ��쳣��¼�������쳣���д���
            blnOK = frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, EM_ED_�쳣����, , lng�������, mblnNOMoved, , , mobjMsgModule, , True)
        Else
            If CheckBillExistReplenishData(0, lng�������) Then

                strPrivsReplenish = ";" & GetPrivFunc(glngSys, 1124) & ";"
                If InStr(strPrivsReplenish, ";�����˷�;") > 0 Then
                    If MsgBox("ѡ��ļ�¼������ҽ�����������Ϊ�쳣��������˷Ѽ�¼���Ƿ���Ըü�¼�����ٴν����˷ѣ�", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
                        blnOK = frmReplenishTheBalanceDel.zlShowMe(Me, 1124, strPrivsReplenish, EM_RBDTY_�쳣����, lng�������, False, 0, False, strDate)
                    Else
                        Exit Function
                    End If
                Else
                    MsgBox "ѡ��ļ�¼������ҽ�����������Ϊ�쳣��������˷Ѽ�¼���㲻�߱������ü�¼��Ȩ�ޣ�����������˷Ѳ�����", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                blnOK = frmClinicDelAndView.ShowMe(Me, EM_MULTI_�쳣����, mstrPrivs, lng�������, , , , strDate)
            End If
        End If
    End If

    If blnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    '����:���¶��쳣�˷ѵ������˷Ѳ���
    '���:strDelTime-�쳣���ݵ��˷�ʱ��
    '
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-22 15:41:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOneCard As Boolean
    On Error GoTo errHandle
    If strNo = "" Then Exit Function
    If InStr(mstrPrivs, ";�����˷�;") = 0 Then Exit Function
    blnOneCard = GetOneCard.RecordCount > 0
    
    If frmMultiBills.ShowMe(gfrmMain, 2, mstrPrivs, strNo, strDelTime, , , blnOneCard) = False Then Exit Function
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
        If gclsInsure.GetCapability(support�൥���շѱ���ȫ��, , intInsure) Then CheckTogetherDo = True
    End If
    
    If gblnMultiBalance Then
        If CheckSingleBalance(strNo) = False Then CheckTogetherDo = True
    End If
End Function

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_Charge_Click()
    If frmClinicCharge.zlEditBill(Me, mlngModul, mstrPrivs, EM_ED_�շ�, , , , , , mobjMsgModule) = True Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ļ�¼����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    '�ش�Ľ��󣬴�����ǽ������
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("�������"))
    If tbPage.Selected.Index = 1 Then
        strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
        '�˷�������棬���ݵ��ݲ鿴
        If isSimple(strNo) Then
            '���շ�
            frmSimpleCharge.mlngModul = mlngModul
            frmSimpleCharge.mstrPrivs = mstrPrivs
            frmSimpleCharge.mbytInState = 1
            frmSimpleCharge.mstrDelete = IIf(blnDel, strDate, "")     'ֻ���˷ѵ��ݲŴ���ʱ������������״̬����
            frmSimpleCharge.mstrInNO = strNo
            frmSimpleCharge.mblnNOMoved = mblnNOMoved
            frmSimpleCharge.Show 1, Me
        Else
            '�����շ�
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
        MsgBox "��ǰû�е��ݿ��Բ��ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    If Not (gbln�˷�����ģʽ And tbPage.Selected.Index = 1) Then
        '�Ƿ�鿴�˷ѵ���
        blnDel = Val(mshList.TextMatrix(mshList.Row, GetColNum("�˷ѷ���"))) < 0
        strDate = mshList.TextMatrix(mshList.Row, GetColNum("�շ�ʱ��"))
    End If
    
    If Val(strNo) < 0 Then
        If blnDel Then
            frmClinicDelAndView.ShowMe Me, EM_MULTI_�鿴, mstrPrivs, Val(strNo), , , mblnNOMoved, strDate
        Else
            frmClinicDelAndView.ShowMe Me, EM_MULTI_�鿴, mstrPrivs, Val(strNo), , , mblnNOMoved
        End If
    Else
        If CheckBalance(Val(strNo)) Then
            strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
            strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
            If UBound(Split(strNos, ",")) > 0 Then
                frmMultiBills.ShowMe gfrmMain, 0, mstrPrivs, strNo, IIf(blnDel, strDate, ""), , , , mblnNOMoved      'ֻ���˷ѵ��ݲŴ���ʱ������������״̬����
            ElseIf isSimple(strNo) Then
                '���շ�
                frmSimpleCharge.mlngModul = mlngModul
                frmSimpleCharge.mstrPrivs = mstrPrivs
                frmSimpleCharge.mbytInState = 1
                frmSimpleCharge.mstrDelete = IIf(blnDel, strDate, "")     'ֻ���˷ѵ��ݲŴ���ʱ������������״̬����
                frmSimpleCharge.mstrInNO = strNo
                frmSimpleCharge.mblnNOMoved = mblnNOMoved
                frmSimpleCharge.Show 1, Me
            Else
                '�����շ�
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
            strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
            frmMultiBills.ShowMe gfrmMain, 0, mstrPrivs, strNo, IIf(blnDel, strDate, ""), , , , mblnNOMoved      'ֻ���˷ѵ��ݲŴ���ʱ������������״̬����
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
    ShowBills IIf(gbln�˷�����ģʽ And tbPage.Selected.Index = 1, mstrFilter2, mstrFilter)
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
    Call Ȩ�޿���
    Call ShowBills(strFilter)
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
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
        Case "����"
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
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    
    If strNo = "" Or strNo = "���ŵ���" Then
        MsgBox "δѡ���κε��ݣ�����ִ�д˲�����", vbExclamation, gstrSysName: Exit Sub
    End If
        
    If Not gobjPlugIn Is Nothing Then
        lngPatiID = zlGet����ID(strNo)
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
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte
    Dim intCurrentRow As Integer
    
    On Error GoTo errHandler
    
    '��ͷ
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    If tbPage.Selected Is Nothing Then Exit Sub
    Select Case tbPage.Selected.Index
    Case 1
        objOut.Title.Text = "�˷������¼�嵥"
    Case 2
        objOut.Title.Text = "�շ��쳣��¼�嵥"
    Case 3
        objOut.Title.Text = "�˷��쳣��¼�嵥"
    Case Else
        If glngSys Like "8??" Then
            objOut.Title.Text = "ҩ���շѵ����嵥"
        Else
            objOut.Title.Text = "���˵��ݼ�¼�嵥"
        End If
    End Select
    
    '����
    If tbPage.Selected.Index = 0 Then
        objRow.Add "ʱ�䣺" & Format(SQLCondition.DateB, "yyyy-mm-dd hh:mm:ss") & " �� " & Format(SQLCondition.DateE, "yyyy-mm-dd hh:mm:ss")
        objOut.UnderAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    mshList.Redraw = False
    intCurrentRow = mshList.Row
    mblnPrinting = True
    
    '����
    Set objOut.Body = mshList
    '���
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
'���ܣ��������޼�¼���ò˵�����״̬
    Dim blnApply As Boolean
    
    If gbln�˷�����ģʽ And tbPage.Selected.Index = 1 Then blnApply = True
    
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
        
    mnuEdit_Apply.Enabled = blnUsed And Not (tbPage.Selected.Index = 1) And gbln�˷�����ģʽ
    mnuEdit_UnApply.Enabled = blnUsed And blnApply
    mnuEdit_Audit.Enabled = blnUsed And blnApply
    mnuEdit_RefuseApply.Visible = blnUsed And blnApply
    mnuEdit_RefuseApply.Visible = blnUsed And blnApply
    mnuEdit_UnAudit.Enabled = blnUsed And blnApply
    
    mnuEdit_DelMulti.Enabled = Not blnApply
    mnuEditWriteCard.Enabled = blnUsed
End Sub

Private Sub Ȩ�޿���()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ȩ�޿���
    '����:���˺�
    '����:2011-09-02 15:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, blnErrRefund As Boolean
    Dim blnErrPage As Boolean   '�쳣����ҳ��
    If Not tbPage.Selected Is Nothing Then
        blnErrPage = tbPage.Selected.Index = 2 Or tbPage.Selected.Index = 3
        blnErrRefund = tbPage.Selected.Index = 3
    Else
        blnErrPage = False
        blnErrRefund = False
    End If
    If tbPage.Selected.Index = 3 Then
        tbr.Buttons("Charge").Caption = "�˷�"
    Else
        tbr.Buttons("Charge").Caption = "�շ�"
    End If
    mnuEditReCharge.Caption = IIf(blnErrRefund, "�����˷�(&R)", "�����շ�")
    If glngSys Like "8??" Then
        mshFact.Visible = False
        picVsc.Visible = False
        mnuEdit_Simple.Visible = False
        mnuEdit_Charge_.Visible = False
        mnuEditReCharge.Visible = False
    End If
    mnuEdit_Apply_.Visible = gbln�˷�����ģʽ
    '---------------------------------------------------------------
    blnHavePrivs = InStr(mstrPrivs, ";�����շ�;") > 0
    mnuEdit_Charge.Visible = blnHavePrivs
    mnuEdit_Simple.Visible = blnHavePrivs
    mnuEdit_Charge_.Visible = blnHavePrivs And Not blnErrPage
    mnuEditReCharge.Visible = blnHavePrivs And blnErrPage
    mnuEditCancelBill.Visible = blnHavePrivs And blnErrPage And Not blnErrRefund
    tbr.Buttons("Charge").Visible = blnHavePrivs
    '�����������ϰ�ť
    tbr.Buttons("Cancel").Visible = blnHavePrivs And blnErrPage And Not blnErrRefund
    '-----------------------------------------------------------
    '�ش���ǿ�������Ĵ�ӡ,�����Ʊ���ʱ�Ĵ�ӡ,�����뱨���Ȩ�޲�ͬ��
    mnuEdit_Print.Visible = Not (InStr(mstrPrivs, ";�ش�Ʊ��;") = 0 Or InStr(mstrPrivs, "�վݴ�ӡ") = 0) And Not blnErrPage
    mnuEditInvoicePrint.Visible = gTy_Module_Para.bytƱ�ݷ������ <> 0 And Not (InStr(mstrPrivs, ";�ش�Ʊ��;") = 0 Or InStr(mstrPrivs, "�վݴ�ӡ") = 0) And Not blnErrPage
    
    '52328: ����Ʊ��
    mnuEdit_Print_Supplemental.Visible = InStr(mstrPrivs, ";����Ʊ��;") > 0 And InStr(mstrPrivs, ";�վݴ�ӡ;") > 0 And Not blnErrPage
    
    mnuEditMakeupPrn.Visible = InStr(mstrPrivs, ";����Ʊ��;") > 0 And InStr(mstrPrivs, ";�վݴ�ӡ;") > 0 And Not blnErrPage
    mnuEdit_PrintProve.Visible = InStr(mstrPrivs, ";֤����ӡ;") > 0 And Not blnErrPage
    mnuEdit_PrintList.Visible = InStr(mstrPrivs, ";��ӡ�嵥;") > 0 And Not blnErrPage
    mnuEdit_PrintDel.Visible = (InStr(mstrPrivs, ";�ش�Ʊ��;") > 0 Or InStr(mstrPrivs, ";����Ʊ��;") > 0) And InStr(mstrPrivs, ";�վݴ�ӡ;") > 0 And Not blnErrPage
    
    blnHavePrivs = InStr(mstrPrivs, ";�ش�Ʊ��;") > 0 Or InStr(mstrPrivs, ";�վݴ�ӡ;") > 0 _
        Or InStr(mstrPrivs, ";֤����ӡ;") > 0 Or InStr(mstrPrivs, ";��ӡ�嵥;") > 0
    mnuEdit_View_.Visible = blnHavePrivs And Not blnErrPage
    
'    mnuEdit_Modi.Visible = InStr(mstrPrivs, ";��¼�޸�;") > 0 And Not blnErrPage
'    tbr.Buttons("Modi").Visible = InStr(mstrPrivs, ";��¼�޸�;") > 0 And Not blnErrPage
    
    mnuEdit_Adjust.Visible = InStr(mstrPrivs, ";��¼����;") > 0 And Not blnErrPage
    mnuEdit_Adjust_.Visible = (InStr(mstrPrivs, ";��¼�޸�;") > 0 Or InStr(mstrPrivs, ";��¼����;") > 0) And Not blnErrPage
    '-------------------------------------------------------------
    '�����˷ѿ���
    blnHavePrivs = InStr(mstrPrivs, ";�����˷�;") > 0
'    mnuEdit_Del.Visible = blnHavePrivs And Not blnErrPage
    mnuEdit_Del_.Visible = blnHavePrivs
    tbr.Buttons("Del").Visible = blnHavePrivs And Not blnErrPage
    tbr.Buttons("Del_").Visible = blnHavePrivs
    mnuEdit_DelMulti.Visible = blnHavePrivs And Not blnErrPage
    mnuEdit_Apply.Visible = blnHavePrivs And Not blnErrPage And gbln�˷�����ģʽ
    mnuEdit_UnApply.Visible = blnHavePrivs And Not blnErrPage And gbln�˷�����ģʽ
    '-------------------------------------------------------------
    '����:36076
    mnuEditMzToZyDel.Visible = InStr(mstrPrivs, ";תסԺ�����˷�;") > 0 And Not mbln�������� And Not blnErrPage
    mnuEditSplitMzToZy.Visible = InStr(mstrPrivs, ";תסԺ�����˷�;") > 0 And Not mbln�������� And Not blnErrPage
    '-------------------------------------------------------------
    mnuEdit_Audit.Visible = InStr(mstrPrivs, ";�˷����;") > 0 And Not blnErrPage And gbln�˷�����ģʽ
    mnuEdit_RefuseApply.Visible = InStr(mstrPrivs, ";�˷����;") > 0 And Not blnErrPage And gbln�˷�����ģʽ And tbPage.Selected.Index = 1
    mnuEdit_UnAudit.Visible = InStr(mstrPrivs, ";�˷����;") > 0 And Not blnErrPage And gbln�˷�����ģʽ
    '-------------------------------------------------------------
    mnuEditSplitMzToZy.Visible = InStr(mstrPrivs, ";������Ϣд��;") > 0 And mstrWriteCardTypeIDs <> "" And Not blnErrPage And tbPage.Selected.Index = 0
    mnuEditWriteCard.Visible = InStr(mstrPrivs, ";������Ϣд��;") > 0 And mstrWriteCardTypeIDs <> "" And Not blnErrPage And tbPage.Selected.Index = 0
        
    '�շ����ʹ���
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";����;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("����").Visible = blnHavePrivs
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
    mbln�������� = Val(zlDatabase.GetPara("����ת����������", glngSys, 1131)) = 1
    If glngSys Like "8??" Then
        Caption = "ҩ���շѹ���": Me.mnuEdit_Charge.Caption = "ҩ���շ�(&A)"
    End If
    i = Val(zlDatabase.GetPara("�쳣���ݲ�ѯ", glngSys, mlngModul, 0, Array(lblȱʡ, cboDate)))
    With cboDate
        .Clear
        .AddItem "����"
        .ListIndex = .NewIndex
        If i = 0 Then .ListIndex = .NewIndex
        .AddItem "�������"
        If i = 1 Then .ListIndex = .NewIndex
        .AddItem "�������"
        If i = 2 Then .ListIndex = .NewIndex
        .AddItem "���һ��"
        If i = 3 Then .ListIndex = .NewIndex
        .AddItem "����"
        If i = 4 Then .ListIndex = .NewIndex
        .AddItem "�Զ���"
        If i = 5 Then .ListIndex = .NewIndex
        dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
        dtpEndDate.MaxDate = dtpStartDate.MaxDate
        dtpEndDate.Value = dtpEndDate.MaxDate
        dtpStartDate.Value = DateAdd("d", -7, dtpEndDate.MaxDate)
    End With
    
    If Not gbln�˷�����ģʽ Then
        tbPage.Item(1).Visible = False
        picCons.Left = picCons.Left - 1200
    End If
    mblnNotClick = True
    tbPage.Item(0).Selected = True
    mblnNotClick = False
    'tbPage.Visible = gbln�˷�����ģʽ
    Call RestoreWinState(Me, App.ProductName)
    Set stbThis.Panels(5).Picture = Me.Picture
    
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    
    '���������˰�ش�ӡ����
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(1)
        End If
        On Error GoTo 0
    End If
    
     '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.���, UserInfo.����)
    End If
    On Error GoTo 0
    
    'Ȩ������
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1121_1")
    Call Ȩ�޿���
    
    Call ClearErrInvoice
    
    If InStr(mstrPrivs, "LED������") = 0 Then gblnLED = False
    
    'ȱʡ��������(������)
    mstrInsure = "" 'ȱʡ����ҽ������ͨ����
    If gbln�˷�����ģʽ Then
        mstrFilter2 = " And A.����ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And Nvl(A.״̬,0) = 0"
    End If
    mstrFilter = " And �Ǽ�ʱ�� Between Trunc(Sysdate) And Trunc(Sysdate+1)-1/24/60/60 And ����Ա����||''=[7]"
    
    SQLCondition.Default = True
    SQLCondition.DateB = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    SQLCondition.DateE = DateAdd("s", -1, DateAdd("d", 1, SQLCondition.DateB))
    frmChargeFilter.mblnDateMoved = False
    
    mbln�շ� = True
    mbln�˷� = False
    
    Call SetHeader
    Call SetInvoiceList
    Call SetBalanceList
    Call SetFactList
    Call SetDetail
    Call SetExtendInfo
    tbSub.Item(2).Visible = False
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "��ˢ���嵥���������ù�������"
    
    '��ʼ����Ϣ�������ģ��
    Call zlMsgModuleInit
    
    Call LoadPlugInMnu
End Sub

Private Sub ClearErrInvoice()
'���ܣ��������Ա�ϴ��쳣�˳�ʱֻ����ʵ��Ʊ�Ŷ�û��ʵ�ʴ�ӡ�ĵ��ݵķ��ü�¼�е�Ʊ�ݺ�
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "Select �Ǽ�ʱ��, ʵ��Ʊ��" & vbNewLine & _
            "From ������ü�¼ A," & vbNewLine & _
            "     (Select Max(NO) NO From ������ü�¼ Where �Ǽ�ʱ�� > Sysdate - 1 And ����Ա���� = [1] And ��¼���� = 1) B" & vbNewLine & _
            "Where A.��¼���� = 1 And A.NO = B.NO And A.ʵ��Ʊ�� Is Not Null And Not Exists (Select 1 From Ʊ�ݴ�ӡ���� C Where C.NO = B.NO)"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.����)
    If rsTmp.RecordCount > 0 Then
        strSql = "Select NO From ������ü�¼ Where �Ǽ�ʱ�� = [1] And ��¼���� = 1 And ʵ��Ʊ�� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CDate(rsTmp!�Ǽ�ʱ��), CStr(rsTmp!ʵ��Ʊ��))
        For i = 1 To rsTmp.RecordCount
            strSql = "Zl_Ʊ����ʼ��_Update('" & rsTmp!NO & "','',1)"
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
    
    '����ؼ���Ⱥ͸߶�
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
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
            Exit For
        End If
    Next
    '�������
    If cboDate.ListIndex < 5 Then
        zlDatabase.SetPara "�쳣���ݲ�ѯ", cboDate.ListIndex, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
    '��ж��Ϣ����
    Call zlMsgModuleUnload
End Sub

Private Sub mnuViewGo_Click()
    If SQLCondition.int�����־ <= 1 Then
        frmChargeGo.lbl��ʶ��.Caption = "�����"
    ElseIf SQLCondition.int�����־ = 2 Then
        frmChargeGo.lbl��ʶ��.Caption = "סԺ��"
    Else
        frmChargeGo.lbl��ʶ��.Caption = "����/סԺ��"
    End If
    frmChargeGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmChargeGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmChargeGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ŵ���")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ŷ�Ʊ")) = .txtFact.Text
            End If
            If .cbo����Ա.ListIndex > 0 Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("����Ա")) = zlStr.NeedName(.cbo����Ա.Text)
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
            If .txt��ʶ��.Text <> "" Then
                blnFill = blnFill And (mshList.TextMatrix(i, GetColNum("�����")) = .txt��ʶ��.Text Or mshList.TextMatrix(i, GetColNum("סԺ��")) = .txt��ʶ��.Text)
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.COLS - 1
            
            mlngGo = i + 1
            
            stbThis.Panels(2).Text = "�ҵ�һ����¼"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            stbThis.Panels(2).Text = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "�Ѷ�λ���嵥β��"
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
    strSql = "Select Distinct a.No From ������ü�¼ A, ����Ԥ����¼ B Where a.����id = b.����id And b.������� = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalance)
    Do While Not rsTmp.EOF
        strNos = strNos & "," & "'" & rsTmp!NO & "'"
        rsTmp.MoveNext
    Loop
    If strNos <> "" Then strNos = Mid(strNos, 2)
    GetAllNos = strNos
End Function

Private Function GetInvoiceRelatedNos(ByVal strNo As String, Optional ByRef strInvoices As String) As String
    '���ܣ�ͨ��һ�ŵ��ݺŻ�ȡƱ�ݴ�ӡ�Ĺ�������
    '������strNo - ���ݺ�


    '      strInvoices - ����Ʊ�ݺ�
    '���أ��������ݺ�
    '����ţ�83602
    Dim strSql As String, rsNos As ADODB.Recordset
    Dim strNos As String, blnNotRule As Boolean '�Ƿ�ʵ�ʴ�ӡ����Ʊ�ŵ�
    Dim strReturnInvoices As String
    
    On Error GoTo ErrHand
    '�ж�Ʊ�ݹ����Ƿ��б�
    strSql = "Select 1 From  Ʊ�ݴ�ӡ��ϸ " & _
            " Where Ʊ�� = 1 And �Ƿ���� <> 1 And NO = [1] And Rownum < 2"
    Set rsNos = zlDatabase.OpenSQLRecord(strSql, "�ж�Ʊ�ŷ��䷽ʽ", strNo)
    
    blnNotRule = rsNos.EOF
    '����Ʊ�ŷ��������ҹ������ݺ�
    'Ԥ���������Ʊ��
    If blnNotRule = False Then
        strSql = "" & _
        "   Select Distinct a.No, a.Ʊ��" & _
        "   From Ʊ�ݴ�ӡ��ϸ A, Ʊ�ݴ�ӡ��ϸ B, Ʊ�ݴ�ӡ��ϸ C" & _
        "   Where a.No = b.No And a.Ʊ�� = b.Ʊ�� And a.�Ƿ���� <> 1" & _
        "       And b.Ʊ�� = c.Ʊ�� And b.Ʊ�� = c.Ʊ�� And b.�Ƿ���� <> 1" & _
        "       And c.Ʊ�� = 1 And c.�Ƿ���� <> 1 And c.No = [1]" & _
        " Order By Ʊ��"
    Else
        strSql = "" & _
        " Select Distinct b.No, a.���� as Ʊ��" & _
        " From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B, Ʊ�ݴ�ӡ���� C" & _
        " Where a.��ӡid = b.Id And a.Ʊ�� = 1 And a.ԭ��<>6" & _
        "       And Not Exists (Select 1 From Ʊ��ʹ����ϸ Where ��ӡid = a.��ӡid And ���� = a.���� And Ʊ�� = a.Ʊ�� And ���� = 2)" & _
        "       And b.Id = c.Id And b.�������� = c.��������" & _
        "       And c.�������� = 1 And c.No = [1]" & _
        " Order By Ʊ��"
    End If

    Set rsNos = zlDatabase.OpenSQLRecord(strSql, "��ȡ�ش򵥾ݺ�", strNo)
    strNos = "": strReturnInvoices = ""
    Do While Not rsNos.EOF
        If InStr(strNos & ",", ",'" & NVL(rsNos!NO) & "',") = 0 Then
            strNos = strNos & ",'" & NVL(rsNos!NO) & "'"
        End If
        If InStr(strReturnInvoices & ",", "," & NVL(rsNos!Ʊ��) & ",") = 0 Then
            strReturnInvoices = strReturnInvoices & "," & NVL(rsNos!Ʊ��)
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
'���ܣ���ǰ�տ��¼���´�ӡһ��Ʊ��
'bytMode=0-�ش�,1-����
    Dim strNo As String, strNos As String, blnMediCare As Boolean
    Dim intInsure As Integer, blnVirtualPrint As Boolean, lng����ID As Long, lng����ID As Long, blnDel As Boolean
    Dim strUseType  As String, lngShareUseID As Long, intInvoiceFormat As Integer
    Dim intOldInvoiceFormat As Integer, lngBalance As Long, lngPJ����ID As Long
    Dim strReclaimInvoice As String '���յ�Ʊ��25187
    Dim blnPrintFact As Boolean '�Ƿ��Ѿ���ӡ��Ʊ��
    Dim blnOnePatiPrint As Boolean '������һ�δ�ӡ
    Dim strPrintNos As String
    Dim blnLocalNo As Boolean  'ָ�����ݴ�ӡ
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    If Me.ActiveControl Is mshDetail Then
        '83602,Ƚ����,2015-3-31,�ش򲿷ֵ���
        If mshDetail.IsSubtotal(mshDetail.Row) Then
            strNo = mshDetail.TextMatrix(mshDetail.Row + 1, 0)
        Else
            strNo = mshDetail.TextMatrix(mshDetail.Row, 0)
        End If
        blnLocalNo = True
    Else
        If bytMode = 0 Then
            If MsgBox("��ȷ��Ҫ�ش򱾴ν�������е�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        blnLocalNo = False
    End If
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����ش�Ʊ�ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lngBalance = Val(mshList.TextMatrix(mshList.Row, GetColNum("�������")))
    blnMediCare = mshList.TextMatrix(mshList.Row, GetColNum("ҽ��")) = "��"
    blnDel = mshList.TextMatrix(mshList.Row, GetColNum("����")) = "2"   '��¼״̬Ϊ2�ģ�Ŀǰ�ǽ����˴�ӡ�˵����
    blnPrintFact = Trim(mshList.TextMatrix(mshList.Row, GetColNum("���ŷ�Ʊ"))) <> ""
  
    '�����˴�ӡ
    blnOnePatiPrint = False
    If bytMode = 0 Then
        If zlIsOnePatiPrint(strNo, strPrintNos, blnOnePatiPrint) = False Then Exit Sub
    End If
    If blnOnePatiPrint Then
        If blnLocalNo Then
            If MsgBox("��ǰѡ��ĵ��ݡ�" & strNo & "���ǰ����˲���ĵ��ݣ����Ƿ����¶Բ���ĵ��ݽ����ش�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        If CheckBillExistReplenishData(1, 0, strPrintNos) = True Then
            MsgBox "ѡ��ļ�¼������ҽ��������㣬����������ش�򲹴�Ʊ�ݲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        If CheckBillExistReplenishData(0, lngBalance) = True Then
            MsgBox "ѡ��ļ�¼������ҽ��������㣬����������ش�򲹴�Ʊ�ݲ�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    If bytMode = 0 Then
        If Not blnOnePatiPrint Then '�����˲���Ʊ�ģ���������صĵ��ݼ��(�ݲ�����,�Ժ��������ټ����������).
            If Not BillOperCheck(2, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), _
                CDate(mshList.TextMatrix(mshList.Row, GetColNum("�շ�ʱ��"))), "�ش�", strNo, , 1) Then Exit Sub
        End If
    Else
        If Trim(mshList.TextMatrix(mshList.Row, GetColNum("���ŷ�Ʊ"))) <> "" _
            And Not (Me.ActiveControl Is mshDetail And mnuFeeDetial_Supplemental.Enabled) Then
            MsgBox "��ǰ�����Ѵ�ӡ��Ʊ��,���ܽ��в���", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    If blnOnePatiPrint Then
        '1.�����˴�ӡʱ����Ҫȫ���ش�
        strNos = "'" & Replace(strPrintNos, ",", "','") & "'"
        '�ش�Ʊ��ʱ,����Ʊ��
        strReclaimInvoice = zlGetReclaimInvoice(Replace(Replace(strNos, "'", ""), ",", ";"))
    ElseIf Me.ActiveControl Is mshDetail Then
        '2.ѡ����ϸ�б��ӡ
        If bytMode = 0 Then
            '2.1ѡ����ϸ�б��ش�
            '83602,Ƚ����,2015-3-31,�ش򲿷ֵ���
            strNos = GetInvoiceRelatedNos(strNo, strReclaimInvoice)
        ElseIf bytMode = 1 Then
            '2.2ѡ����ϸ�б���
            If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ Then
                '��ʵ�ʴ�ӡ����Ʊ���ҷֱ��ӡ����ֻ����ǰѡ�񵥾�
                strNos = "'" & strNo & "'"
            Else
                strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
            End If
        End If
    Else
        '3.ѡ������б��ӡʱ�����۲����ش�ȫ����
        strNos = GetMultiNOs(strNo, , mblnNOMoved, True)
        If bytMode = 0 Then
            '�ش�Ʊ��ʱ,����Ʊ��
            strReclaimInvoice = zlGetReclaimInvoice(Replace(Replace(strNos, "'", ""), ",", ";"))
        End If
    End If
    
    If blnMediCare Then
        intInsure = ChargeExistInsure(strNo, lng����ID, lng����ID, , blnDel)
        If intInsure <> 0 Then
            blnVirtualPrint = gclsInsure.GetCapability(supportҽ���ӿڴ�ӡƱ��, lng����ID, intInsure, CStr(lng����ID))
        End If
    End If
    
    lng����ID = zlGet����ID(strNo)
    lngPJ����ID = zlGet����ID(strNo)
    strUseType = zl_GetInvoiceUserType(lng����ID, 0, intInsure)
    lngShareUseID = zl_GetInvoiceShareID(mlngModul, strUseType)
    intInvoiceFormat = zl_GetInvoicePrintFormat(mlngModul, strUseType, intOldInvoiceFormat, blnOnePatiPrint)
    '������ʣ�������Ĳſ����ش򣬱���ҽ������ʹ������Ҳ�������´�ӡ
    If Not blnVirtualPrint Then
        If Not BillExistMoney(strNos, 1) Then
            MsgBox "�����е���Ŀ�Ѿ�ȫ���˷�,���ܽ��д�ӡ��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If (Me.ActiveControl Is mshDetail And bytMode = 0 Or blnOnePatiPrint) And strReclaimInvoice <> "" Then
            '���ѻ���Ʊ��
            MsgBox "ע��:" & vbCrLf & "    ����Ҫ��������Ʊ�ݣ�" & vbCrLf & _
                    Replace("    " & strReclaimInvoice, ",", "��"), vbInformation, gstrSysName
    End If
    
    '�ش�ʱ����ʹ��ԭƱ�ݴ�ӡ��ʽ
    If bytMode = 0 And blnOnePatiPrint = False Then intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.bytƱ�ݷ������ <> 0, intOldInvoiceFormat, intInvoiceFormat)
    
    Dim strPriceGrade As String
    If gintPriceGradeStartType >= 2 Then
        strPriceGrade = GetPriceGradeFromNos(strNos)
    Else
        strPriceGrade = gstr��ͨ�۸�ȼ�
    End If
    If RePrintCharge(IIf(bytMode = 0, 1, 2), strNos, Me, 0, strReclaimInvoice, , , _
        intInvoiceFormat, blnVirtualPrint, blnDel, lngShareUseID, strUseType, blnOnePatiPrint, strPriceGrade) Then

        '��ҽһ��ͨд����85950
        Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, 0, strNos)
        
        '81688:���ϴ�,2015/5/18,������
        If Not gobjPlugIn Is Nothing And bytMode = 1 Then
            On Error Resume Next
            Call gobjPlugIn.OutPatiInvoicePrintAfter(lng����ID, lngPJ����ID)
            Err.Clear
        End If
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    
    strHead = "���ݺ�,4,850|������,1,800|����ʱ��,1,1850|����ԭ��,1,3000|���״̬,4,1000" & _
            "|�����,1,800|���ʱ��,1,1850|���ԭ��,1,3500|��¼����,1,0|�������,1,0"
    
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
        '�ָ��ϴ���
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
    
    If gbln�˷�����ģʽ And tbPage.Selected.Index = 1 Then
        Call SetApplyHeader
        Exit Sub
    End If
    If tbPage.Selected.Index = 0 Then
        strHead = "ҽ��,4,450|���ŵ���,4,1000|���ŷ�Ʊ,4,1000|����,1,1200" & _
            "|�Ա�,4,500|����,4,500|�����,1,800|סԺ��,1,800|�ѱ�,4,750|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|����Ա,4,1200" & _
            "|�շ�ʱ��,4,1850|�������,1,0|����,1,0|�˷ѷ���,1,0|��Ʊ�Ѵ�ӡ,1,0"
    Else
        strHead = "ҽ��,4,450|���ŵ���,4,1000|����,1,1200" & _
            "|�Ա�,4,500|����,4,500|�����,1,800|סԺ��,1,800|�ѱ�,4,750|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|����Ա,4,1200" & _
            "|�շ�ʱ��,4,1850|�������,1,0|����,1,0|�˷ѷ���,1,0|��Ʊ�Ѵ�ӡ,1,0"
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
        
        i = GetColNum("סԺ��")
        .ColWidth(i) = IIf(SQLCondition.int�����־ = 1, 0, IIf(.ColWidth(i) <= 0, 800, .ColWidth(i)))
        i = GetColNum("�����")
        .ColWidth(i) = IIf(SQLCondition.int�����־ = 2, 0, IIf(.ColWidth(i) <= 0, 800, .ColWidth(i)))
        
        .RowHeight(0) = 350
        
        '�ָ��ϴ���
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
'���ܣ���ʾ�˷����뵥��
    Dim i As Long, j As Long, k As Long, strSql As String
    
    On Error GoTo errH
    
    '�����:53953
    If Not blnSort Then
        strSql = "Select a.No As ���ŵ���, a.������, To_Char(a.����ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ����ʱ��, a.����ԭ��," & vbNewLine & _
                "        Decode(Nvl(a.״̬, 0), 1, 'ͨ��', 2, '�ܾ�', '����') As ���״̬, a.�����, To_Char(a.���ʱ��, 'YYYY-MM-DD HH24:MI:SS') As ���ʱ��," & vbNewLine & _
                "        a.���ԭ��, a.��¼����, b.�������" & vbNewLine & _
                " From �����˷����� A," & vbNewLine & _
                "      (Select Distinct m.No, Nvl(n.�������, n.����id) As �������" & vbNewLine & _
                "        From ������ü�¼ M, ����Ԥ����¼ N" & vbNewLine & _
                "        Where m.����id = n.����id And m.��¼���� = 1 And m.��¼״̬ In (1, 3)) B" & vbNewLine & _
                " Where a.No = b.No And (1 = 1 " & strFilter & ")" & vbNewLine & _
                " Order By a.����ʱ�� Desc, a.���ʱ�� Desc"
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
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        
        k = GetColNum("���״̬")
        For i = 1 To mshList.Rows - 1
            If Trim(mshList.TextMatrix(i, k)) = "ͨ��" Then
                '���ͨ��������ɫ��ʾ
                mshList.Cell(flexcpForeColor, i, 0, i, mshList.COLS - 1) = &HC00000
            ElseIf Trim(mshList.TextMatrix(i, k)) = "�ܾ�" Then
                '��˾ܾ����ú�ɫ��ʾ
                mshList.Cell(flexcpForeColor, i, 0, i, mshList.COLS - 1) = &HC0
            Else
                mshList.Cell(flexcpForeColor, i, 0, i, mshList.COLS - 1) = vbBlack
            End If
        Next
        
        stbThis.Panels(2).Text = "�� " & mrsList.RecordCount & " �ν���"
        Call SetMenu(True)
    End If
    
    Call SetHeader
    Call SetApplyDetail
    Call SetApplyFactList
    
    
        
    '����AfterRowColChange�¼�
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
'����:��������ȡ�����б�(���˹���)
'����:strFilter=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, j As Long, k As Long, l As Long
    Dim strSql As String
    Dim dtStartDate As Date, dtEndDate As Date
    Dim strErrWhere As String
    Dim strWhere As String
    Dim strFeeTable As String
    Dim strTemp As String, strSQL1 As String
    
    On Error GoTo errH
    strErrWhere = "": strWhere = ""
    If gbln�˷�����ģʽ And tbPage.Selected.Index = 1 Then
        Call ShowApplyBills(strFilter, blnSort)
        Exit Sub
    End If
    
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        
        If mbln�շ� And mbln�˷� Then
            '���з��ü�¼
            strWhere = " Where Mod(��¼����, 10) = 1 And ��¼״̬ IN([13],[14],[15]) "
        ElseIf mbln�շ� Then
            'ԭʼ�շѼ�¼
            strWhere = " Where ��¼���� = 1 And ��¼״̬ IN([13],[15]) "
        ElseIf mbln�˷� Then
            '�˷Ѽ�¼�Լ����ռ�¼
            strWhere = " Where (Mod(��¼����, 10) = 1 And ��¼״̬ = [14] Or ��¼���� = 11 And ��¼״̬ In ([13],[15])) "
        End If
        
        Select Case SQLCondition.int�����־
        Case 1 '����
            strWhere = strWhere & " And  �����־ in (1,4)"
        Case 2 'סԺ
            strWhere = strWhere & " And  �����־ =2"
        Case Else   '����
        End Select
        
        strErrWhere = ""
        If tbPage.Selected.Index = 2 Or tbPage.Selected.Index = 3 Then
            Select Case cboDate.ListIndex
            Case 0 '����
                dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtStartDate, "yyyy-mm-dd") & " 23:59:59")
            Case 1 '�������
                dtStartDate = CDate(Format(DateAdd("d", -1, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 2 '�������
                dtStartDate = CDate(Format(DateAdd("d", -2, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 3  '����
                dtStartDate = CDate(Format(DateAdd("d", -7, dtpStartDate.MaxDate), "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case 4  '����
                dtStartDate = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-01") & " 00:00:00")
                dtEndDate = CDate(Format(dtpStartDate.MaxDate, "yyyy-mm-dd") & " 23:59:59")
            Case Else
                dtStartDate = CDate(Format(dtpStartDate.Value, "yyyy-mm-dd") & " 00:00:00")
                dtEndDate = CDate(Format(dtpEndDate.Value, "yyyy-mm-dd") & " 23:59:59")
            End Select
            lblDateShow.Caption = Format(dtStartDate, "yyyy-mm-dd HH:MM:SS")
            lblDateShow.Caption = lblDateShow.Caption & "~" & Format(dtEndDate, "yyyy-mm-dd HH:MM:SS")
            
            '�շ��쳣��¼
            If tbPage.Selected.Index = 2 Then
                strErrWhere = _
                " Where Nvl(����״̬,0) = 1 And ��¼���� = 1 And ��¼״̬ = 1" & vbNewLine & _
                "       And �Ǽ�ʱ�� Between [1] and [2] And ����Ա����=[3]  " & _
                "       And Not Exists (Select 1" & vbNewLine & _
                "                       From ������ü�¼ B" & vbNewLine & _
                "                       Where a.No = b.No" & vbNewLine & _
                "                             And Mod(b.��¼����, 10) = 1 And b.��¼״̬ = 2) "
            Else
                '�˷��쳣��¼
                strErrWhere = _
                " Where Nvl(����״̬,0) = 1 And (Mod(��¼����, 10) = 1 And ��¼״̬ = 2 Or ��¼���� = 11 And ��¼״̬ = 1)" & vbNewLine & _
                "       And �Ǽ�ʱ�� Between [1] and [2] And ����Ա����=[3]  " & _
                "       And Exists (Select 1" & vbNewLine & _
                "                   From ����Ԥ����¼ B " & vbNewLine & _
                "                   Where a.����id=b.����ID and Nvl(b.У�Ա�־,0) <> 0) "
            End If
            strWhere = strErrWhere
        Else
            strWhere = strWhere & " And Nvl(����״̬,0) <> 1  "
        End If
        
        If tbPage.Selected.Index = 0 Then strWhere = strWhere & " " & strFilter
        'strFilter�п�����Ϊ�ѱ����������Ӳ�ѯ
        strFeeTable = _
            " Select ����ID,��¼״̬,Max(Decode(n.��ӡID, Null, 0, 1)) As ��Ʊ�Ѵ�ӡ" & vbNewLine & _
            " From ������ü�¼ A,Ʊ�ݴ�ӡ���� M, Ʊ��ʹ����ϸ N" & _
                strWhere & " And a.No = m.No(+) And m.Id = n.��ӡid(+) And n.����(+) = 1 And n.ԭ��(+) = 6" & vbNewLine & _
            " Group By ����ID,��¼״̬"
        
        'ע��:һ���շѵ��ݿ���ʹ�ö��Ʊ�ݺ�,�������ش�����ŵ�����ʾΪ��ʼƱ�ݺ�(���=1)
        '����:1=�õ���δ�˹�,3-�õ��ݱ��˹�,2-�õ���Ϊ�˵ļ�¼
        '��Ϊ��ǰһ�ŵ������ж��Ʊ�ݺ�,�ݴ���Ϊȡ��ʼ��(�³�������ֻ��һ������)
'        IIf(mbln�˷� = False, " And Exists (Select 1 From ������ü�¼ Where ��¼����=1 And ��¼״̬ In (1,3) And ����id=z.����id)", "")
        strSQL1 = _
        "      Select Distinct a.����ID,nvl(B.�������,a.����ID) as �������, " & _
        "               Max(Decode(Nvl(t.����,0),0,0,1)) as ҽ��, " & _
        "               Max(decode(A.��¼״̬,2,1,0)) as �˷ѱ�־,Max(a.��Ʊ�Ѵ�ӡ) as ��Ʊ�Ѵ�ӡ" & _
        "      From (" & strFeeTable & ") A, ����Ԥ����¼ B, ���ս����¼ T" & _
        "      Where A.����ID=B.����ID(+) And  A.����id=t.��¼ID(+) And t.����(+)=1 " & mstrInsure & _
        "      Group by a.����ID,nvl(B.�������,a.����ID)"
        strTemp = ""
        If frmChargeFilter.mblnDateMoved Then
            strTemp = Replace(strSQL1, "������ü�¼", "H������ü�¼")
            strTemp = Replace(strTemp, "����Ԥ����¼", "H����Ԥ����¼")
            strTemp = Replace(strTemp, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
            strTemp = Replace(strTemp, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
        End If
        strSQL1 = " With C_������Ϣ As (" & strSQL1 & ")" & IIf(strTemp = "", "", ",C_H������Ϣ As (" & strTemp & ")")

        If tbPage.Selected.Index = 0 Then
            strSql = _
            " Select Decode(Max(J.ҽ��),1,'��','') as ҽ��,Min(A.NO) As ���ŵ���,Min(A.ʵ��Ʊ��) As ���ŷ�Ʊ," & _
            "       A.����,A.�Ա�,A.����,Decode(A.�����־,2,'',A.��ʶ��) As �����,Decode(A.�����־,2,A.��ʶ��,'') As סԺ��," & _
            "       Min(A.�ѱ�) as �ѱ�, " & _
            "       To_Char(decode(max(J.�˷ѱ�־),1,-1,1)*Sum(a.Ӧ�ս��), '999999999" & gstrDec & "') as Ӧ�ս��," & _
            "       To_Char(decode(max(J.�˷ѱ�־),1,-1,1)*Sum(a.ʵ�ս��), '999999999" & gstrDec & "') as ʵ�ս��," & _
            "       A.����Ա���� as ����Ա,To_Char(Min(A.�Ǽ�ʱ��),'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,J.�������," & _
            "       Max(A.��¼״̬) as ����,Min(a.ִ��״̬) As �˷ѷ���,Max(j.��Ʊ�Ѵ�ӡ) as ��Ʊ�Ѵ�ӡ" & _
            " From ������ü�¼ A,C_������Ϣ J,���ű� B,ҽ�Ƹ��ʽ C" & _
            " Where A.����ID=J.����ID And A.��������ID=B.ID " & _
            "       And A.���ʽ=C.����(+) " & IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.����=[16]", "") & _
            "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & _
            " Group by A.����,A.�Ա�,A.����,A.�����־,A.��ʶ��,A.����Ա����,J.�������"
        Else
            strSql = _
            " Select Decode(Max(J.ҽ��),1,'��','') as ҽ�� ,Min(A.NO) As ���ŵ���," & _
            "       A.����,A.�Ա�,A.����,Decode(A.�����־,2,'',A.��ʶ��) As �����,Decode(A.�����־,2,A.��ʶ��,'') As סԺ��," & _
            "       Min(A.�ѱ�) as �ѱ�,   " & _
            "       To_Char(decode(max(J.�˷ѱ�־),1,-1,1)*Sum(a.Ӧ�ս��), '999999999" & gstrDec & "') as Ӧ�ս��," & _
            "       To_Char(decode(max(J.�˷ѱ�־),1,-1,1)*Sum(a.ʵ�ս��), '999999999" & gstrDec & "') as ʵ�ս��," & _
            "       A.����Ա���� as ����Ա,To_Char(Min(A.�Ǽ�ʱ��),'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,J.�������," & _
            "       Max(A.��¼״̬) as ����,Min(a.ִ��״̬) As �˷ѷ���,Max(j.��Ʊ�Ѵ�ӡ) as ��Ʊ�Ѵ�ӡ" & _
            " From ������ü�¼ A,C_������Ϣ J,���ű� B,ҽ�Ƹ��ʽ C" & _
            " Where  A.��������ID=B.ID And A.���ʽ=C.����(+) And A.����ID=J.����ID " & IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.����=[16]", "") & _
            "       And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
            " Group by A.����,A.�Ա�,A.����,A.�����־,A.��ʶ��,A.����Ա����,J.������� "
        End If
        
        If frmChargeFilter.mblnDateMoved Then
            strSql = strSql & vbNewLine & _
                    " Union All" & vbNewLine & _
                      Replace(Replace(strSql, "������ü�¼", "H������ü�¼"), "C_������Ϣ", "C_H������Ϣ")
        End If
        
        strSql = "Select * From (" & strSQL1 & strSql & ")  " & _
                " Order By �շ�ʱ�� Desc"
                
        With SQLCondition
            If SQLCondition.Default Then SQLCondition.Operator = UserInfo.����
            If strErrWhere <> "" Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.����)
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
            tbPage.Selected.Caption = "�շ��쳣��¼"
        Else
            tbPage.Selected.Caption = "�շ��쳣��¼(" & mrsList.RecordCount & ")"
        End If
    End If
    
    If tbPage.Selected.Index = 3 Then
        If mrsList.RecordCount = 0 Then
            tbPage.Selected.Caption = "�˷��쳣��¼"
        Else
            tbPage.Selected.Caption = "�˷��쳣��¼(" & mrsList.RecordCount & ")"
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
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        If tbPage.Selected.Index = 0 Then
         '��ʵ�պϼƽ��,Ŀǰ�˴�û�а���Ž��л���,����в����˷ѵ����,ʵ�ս��᲻��ȷ
            If Not blnSort Then
                strFeeTable = "Select Distinct ����id From ������ü�¼ A" & strWhere
                'һ�ν�������е���
                strFeeTable = _
                    "Select Distinct m.Id, m.��¼����, m.No, m.��¼״̬, m.���, " & vbNewLine & _
                    "       m.���ʽ, m.���˿���id, m.ʵ�ս��, m.��������id, m.����id" & vbNewLine & _
                    "From ������ü�¼ M,(" & strFeeTable & ") N" & vbNewLine & _
                    "Where m.����id = n.����id"

                If frmChargeFilter.mblnDateMoved Then
                    strTemp = Replace(strFeeTable, "������ü�¼", "H������ü�¼")
                    strTemp = Replace(strTemp, "����Ԥ����¼", "H����Ԥ����¼")
                    strTemp = Replace(strTemp, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
                    strTemp = Replace(strTemp, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
                    strFeeTable = strFeeTable & " Union All " & strTemp
                End If
                
                strSql = "With ������� As (" & strFeeTable & ")" & vbNewLine & _
                        " Select " & IIf(mbln�շ� = False And mbln�˷�, -1, 1) & "*Sum(a.ʵ�ս��) As ���," & vbNewLine & _
                        "        Count(Distinct Decode(a.��¼����, 11, Null, a.����id)) As ����" & vbNewLine
                If mstrInsure = "" Then
                    strSql = strSql & _
                        " From ������� A, ���ű� B, ҽ�Ƹ��ʽ C" & vbNewLine & _
                        " Where A.��������ID = B.ID And (B.վ��='" & gstrNodeNo & "' Or B.վ�� is Null)" & vbNewLine & _
                        "       And A.���ʽ=C.����(+) " & IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.����=[16]", "")
                Else
                    strSql = strSql & _
                        " From (Select a.ʵ�ս��, a.��¼����, a.����id" & vbNewLine & _
                        "       From ������� A, ���ű� B, ҽ�Ƹ��ʽ C, ���ս����¼ T" & vbNewLine & _
                        "       Where a.��������id = b.Id And (b.վ�� = '" & gstrNodeNo & "' Or b.վ�� Is Null) And a.���ʽ = c.����(+)" & vbNewLine & _
                                IIf(SQLCondition.PayKindName <> "" And strErrWhere = "", " And C.����=[16]", "") & _
                        "             And a.����id = t.��¼id(+) And t.����(+) = 1 " & mstrInsure & vbNewLine & _
                        "       Group By a.Id, a.ʵ�ս��, a.��¼����, a.����id) A"
                End If
                strSql = "Select ����, ��� From (" & strSql & ")"
                If strErrWhere <> "" Then
                        Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, dtStartDate, dtEndDate, UserInfo.����, 1, 2, 3)
                Else
                    With SQLCondition
                        Set mrsTotal = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .ChargeKind, .NOB, .NOE, .PayKind, .Operator, .PatientName, Val(.PatientIdentity), .FactB, .FactE, .DeptID _
                                        , 1, 2, 3, .PayKindName, .Doctor, .FeeItems, .PatientID)
                    End With
                End If
            End If
            Set mshList.DataSource = mrsList
            stbThis.Panels(2).Text = "�� " & NVL(mrsTotal!����, 0) & " �ν���,�ϼ�:" & Format(NVL(mrsTotal!���, 0), gstrDec)
        Else
            Set mshList.DataSource = mrsList
            stbThis.Panels(2).Text = "�� " & mrsList.RecordCount & " ���쳣��¼"
        End If
        Call SetMenu(True)
    End If
    
    With mshList
        .Redraw = False
        '������ɫ
        .ForeColor = ForeColor
        k = GetColNum("����")
        l = GetColNum("�˷ѷ���")
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, l)) < 0 Then
                '�˷Ѽ�¼�ú�ɫ
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = &HC0
            ElseIf Val(.TextMatrix(i, k)) = 1 And Val(.TextMatrix(i, l)) >= 0 Then
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
            Else
                '�����˹��ѵ�����ɫ
                .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = &HC00000
            End If
        Next
        
        Call SetHeader
        Call SetDetail
        Call SetFactList
        
        '����AfterRowColChange�¼�
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
        strSql = "Select 1 From ����Ԥ����¼ Where �������= [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
        CheckBalance = rsTemp.EOF
    Else
        strSql = "Select 1 From ����Ԥ����¼ A,������ü�¼ B Where B.NO= [1] And Mod(B.��¼����,10) = 1 And B.����id=A.����id And Nvl(A.�������,0) < 0 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
        CheckBalance = Not rsTemp.EOF
    End If
End Function

Private Sub DetailSplitGroup()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Է����б���Ϣ���з�����ʾ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim strTemp As String

    On Error GoTo errHandle
    With mshDetail
        For i = 0 To .COLS - 1
            If i < .ColIndex("���") And i > .ColIndex("˵��") Then
                .ColHidden(i) = True
            End If
        Next
        
        .OutlineBar = flexOutlineBarComplete
        .Subtotal flexSTClear
        .MultiTotals = True
        '&H8000000F
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("ʵ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("Ӧ�ս��"), gstrDec, &H8000000F, , True, "%s", , True
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("���")
        .OutlineCol = .ColIndex("���")

        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("���")) = strTemp

                strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("���ݺ�"))
                If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ Then
                    '83446,����ǰ�ʵ�ʴ�ӡ����Ʊ��,�Ҷ��ŵ����շѷֱ��ӡ,���ڵ�����ʾ����������ʾ��Ʊ��
                    strTemp = strTemp & Space(2) & "��Ʊ��:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("��Ʊ��"))
                End If
                strTemp = strTemp & Space(2) & "�ѱ�:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("�ѱ�"))
                strTemp = strTemp & Space(2) & "��������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("��������"))
                If InStr(mstrPrivs, "��ʾ������") <> 0 Then
                   strTemp = strTemp & Space(2) & "������:" & .Cell(flexcpTextDisplay, i + 1, .ColIndex("������"))
                End If
                .MergeRow(i) = True
                .MergeCells = flexMergeRestrictRows
                .Cell(flexcpAlignment, i, .ColIndex("���"), i, .ColIndex("���")) = 1
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
'                 If Val(.TextMatrix(i + 1, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                 
                For j = 0 To .COLS - 1
                   If j < .ColIndex("Ӧ�ս��") Then
                       If j >= .ColIndex("���") Then
                           .Cell(flexcpText, i, j) = strTemp
                           .Cell(flexcpFontBold, i, j) = False
                       End If
                   ElseIf .ColIndex("ʵ�ս��") = j Then
                       .TextMatrix(i, j) = Format(Val(.TextMatrix(i, j)), gstrDec)
                       .Cell(flexcpFontBold, i, j) = False
                   ElseIf .ColIndex("Ӧ�ս��") = j Then
                       .TextMatrix(i, j) = " " & Format(Val(.TextMatrix(i, j)), gstrDec)
                       .Cell(flexcpFontBold, i, j) = False
                   End If
                Next
            Else
                .TextMatrix(i, .ColIndex("����")) = Format(Val(.TextMatrix(i, .ColIndex("����"))), gstrFeePrecisionFmt)
                .TextMatrix(i, .ColIndex("����")) = Formatex(Val(.TextMatrix(i, .ColIndex("����"))), 5)
                .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("Ӧ�ս��"))), gstrDec)
                .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(Val(.TextMatrix(i, .ColIndex("ʵ�ս��"))), gstrDec)
            End If
        Next
        Call .AutoSize(.ColIndex("���"))
        Call .AutoSize(.ColIndex("����"))
        
        For j = 0 To .COLS - 1
            If j < .ColIndex("Ӧ�ս��") Then
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
'����:strDate:���ݵĵǼ�ʱ��
    Dim i As Long, j As Long, strSql As String, blnDel As Boolean, strDate As String
    
    On Error GoTo errH
    
    If frmChargeFilter.mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("������ü�¼", strNo, , "1")
    Else
        mblnNOMoved = False
    End If
    
    strSql = _
    " Select C.���� as ���,Nvl(E.����,B.����) as ����," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & "B.���," & _
            IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
    "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
            IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
    "       A.�ѱ�,To_Char(Sum(A.��׼����)" & _
            IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
    "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
    "       D.���� as ִ�п���,Nvl(A.��������,B.��������) as ����," & _
    "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��',9,'�쳣�շѵ�','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��," & _
    "       A.��¼״̬" & _
    " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,�շ���Ŀ���� E,ҩƷ��� X" & _
              IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
    " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.��¼����=1 And A.NO=[1] And A.��¼״̬" & IIf(blnDel, "=2", " IN(1,3)") & _
            IIf(strDate <> "", " And A.�Ǽ�ʱ��=[2]", "") & _
    "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
            IIf(Not gblnShowErr, " And Nvl(A.���ӱ�־,0)<>9", "") & _
    " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����)," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� ,", "") & " B.���,A.���㵥λ,A.�ѱ�,D.����," & _
    "       Nvl(A.��������,B.��������),A.ִ��״̬,A.��¼״̬,X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1)" & _
    " Order by Nvl(A.�۸񸸺�,A.���)"

    Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, "")
    
    mshDetail.Redraw = False
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
'    mshDetail.ForeColor = IIf(blnDel, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail

    'ԭʼ�����˹���Ϊ��ɫ
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
        strHead = "���,1,750|����,1,1800" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,2000", "") & "|���,1,1000|��λ,4,500|����,7,850|�ѱ�,1,750|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|��ҩҩ��,1,850|����,1,850|˵��,1,1000|��¼״̬,1,0"
    Else
        strHead = "���,1,750|����,1,1800" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,2000", "") & "|���,1,1000|��λ,4,500|����,7,850|�ѱ�,1,750|����,7,850|Ӧ�ս��,7,850|ʵ�ս��,7,850|ִ�п���,1,850|����,1,850|˵��,1,1000|��¼״̬,1,0"
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
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
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
    '����:��ȡ��ص���ϸ����
    '����:���ݻ�ȡ�ɹ�����true,���򷵻�False
    '����:������
    '����:2014-06-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsMain As ADODB.Recordset
    Dim strWithTable As String, strWhere As String, i As Long
    Dim strTable As String, lngMainRow As Long, strNo As String
    On Error GoTo errHandle
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ŵ���"))
    
    If frmChargeFilter.mblnDateMoved Then
        mblnNOMoved = zlDatabase.NOMoved("������ü�¼", strNo, , "1")
    Else
        mblnNOMoved = False
    End If
    
    If bytType = 2 Then
        '10.29��ǰ���ݵĻ�ȡ
        strSql = _
            " Select NO As ���ݺ�, Max(��Ʊ��) As ��Ʊ��, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, " & _
            "       Sum(����) As ����, ����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, ִ�п���, " & _
            "       Max(����) As ����, ˵��,������,ҽ�Ƹ��ʽ,Max(ժҪ), ��¼״̬" & vbNewLine & _
            " From (Select a.����ID,D1.���� as ��������,A.������,a.No,Max(a.ʵ��Ʊ��) As ��Ʊ��,C.���� as ���,Nvl(E.����,B.����) as ����,E1.���� as ��Ʒ��,B.���," & _
                    IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
            "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                    IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
            "       a.�ѱ�,To_Char(Sum(A.��׼����)" & _
                    IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
            "       D.���� as ִ�п���,Max(Nvl(A.��������,B.��������)) as ����," & _
            "       Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��',9,'�쳣�շѵ�','��'||ABS(A.ִ��״̬)||'���˷�') as ˵��," & _
            "       A.��¼״̬, Nvl(a.�۸񸸺�, a.���) As ���, A.������,F.���� As ҽ�Ƹ��ʽ,Max(ժҪ) As ժҪ " & _
            " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� D1,�շ���Ŀ���� E,�շ���Ŀ���� E1,ҽ�Ƹ��ʽ F,ҩƷ��� X" & _
            " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
            "       And A.��¼����=1 And A.����ID = [1] And A.��¼״̬" & IIf(blnDel, "=2", " IN(1,3)") & _
            "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And a.��������ID=D1.ID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 And A.���ʽ=F.����(+) " & _
            " Group by a.����id, D1.����, a.������, A.������,F.����,a.�ѱ�,a.No,Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����),E1.���� , B.���,A.���㵥λ,D.����," & _
            "       A.ִ��״̬,A.��¼״̬,X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1) )" & _
            " Group By NO, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, ����, ִ�п���, ˵��, ��¼״̬,������,ҽ�Ƹ��ʽ " & _
            " Order By ���ݺ�, ���"
    Else
        strSql = _
            " Select NO As ���ݺ�, Max(��Ʊ��) As ��Ʊ��, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, " & _
            "       Sum(����) As ����, ����, Sum(Ӧ�ս��) As Ӧ�ս��, Sum(ʵ�ս��) As ʵ�ս��, ִ�п���, " & _
            "       Max(����) As ����, Max(˵��),������,ҽ�Ƹ��ʽ,Max(ժҪ),Max(״̬), Min(�˷�״̬)" & vbNewLine & _
            " From (Select a.����ID,D1.���� as ��������,A.������,a.No,Max(a.ʵ��Ʊ��) As ��Ʊ��,C.���� as ���,Nvl(E.����,B.����) as ����,E1.���� as ��Ʒ��,B.���," & _
                    IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ��λ," & _
            "       To_Char(Avg(Nvl(A.����,1)*" & IIf(blnDel, "-1*", "") & "A.����)" & _
                    IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ",'9999990.00000') as ����, " & _
            "       a.�ѱ�,To_Char(Sum(A.��׼����)" & _
                    IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ",'999999" & gstrFeePrecisionFmt & "') as ����, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.Ӧ�ս��),'9999999" & gstrDec & "') as Ӧ�ս��, " & _
            "       To_Char(Sum(" & IIf(blnDel, "-1*", "") & "A.ʵ�ս��),'9999999" & gstrDec & "') as ʵ�ս��, " & _
            "       D.���� as ִ�п���,Max(Nvl(A.��������,B.��������)) as ����,Max(Decode(A.��¼״̬,2,'��'||ABS(A.ִ��״̬)||'���˷�',Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��'))) As ˵��," & _
            "       Max(A.��¼״̬) As ״̬,Min(A.��¼״̬) As �˷�״̬, Nvl(a.�۸񸸺�, a.���) As ���, A.������,F.���� As ҽ�Ƹ��ʽ,Max(ժҪ) As ժҪ " & _
            " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� D1,�շ���Ŀ���� E,�շ���Ŀ���� E1,ҽ�Ƹ��ʽ F,ҩƷ��� X," & _
            "       (Select Distinct ����ID From " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ Where �������= [1]) F" & _
            " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And A.�շ�ϸĿID=X.ҩƷID(+)" & _
            "       And Mod(A.��¼����,10)=1 And A.����ID = F.����ID " & _
            "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
            "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And A.��������ID=D1.ID(+) And E1.����(+)=1 And E1.����(+)=3 And A.���ʽ=F.����(+) " & _
            " Group by a.����id, D1.����, A.������,F.����,a.������, a.�ѱ�,a.No,Nvl(A.�۸񸸺�,A.���),C.����,Nvl(E.����,B.����),E1.���� , B.���,A.���㵥λ,D.����," & _
            "       X.ҩƷID,X." & gstrҩ����λ & ",Nvl(X." & gstrҩ����װ & ",1) )" & _
            " Group By NO, ���, ��������, ������, �ѱ�, ���, ����, ��Ʒ��, ���, ��λ, ����, ִ�п���, ������,ҽ�Ƹ��ʽ Having Sum(����) <> 0" & _
            " Order By ���ݺ�, ���"
    End If
    
    Set rsMain = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    Set mshDetail.DataSource = rsMain
    
    '83446,����ǰ�ʵ�ʴ�ӡ����Ʊ��,�Ҷ��ŵ����շѷֱ��ӡ,���ڵ�����ʾ����������ʾ��Ʊ��
    If gTy_Module_Para.bytƱ�ݷ������ = 0 And gTy_Module_Para.bln�ֱ��ӡ Then
        Dim rsInVoice As ADODB.Recordset, lngRow As Long
        strSql = "Select b.No, f_List2str(Cast(Collect(Distinct a.���� Order By a.���� Asc) As t_Strlist)) As ����" & vbNewLine & _
                " From ������ü�¼ C, Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B" & vbNewLine & _
                " Where b.No = c.No And a.��ӡid = b.Id And a.Ʊ�� = 1 And a.���� = 1 And a.ԭ��<>6" & vbNewLine & _
                "   And Not Exists(Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = 1 And ���� = 2)" & vbNewLine & _
                "   And c.����id " & IIf(bytType = 2, "=[1]", "In (Select ����id From ����Ԥ����¼ Where ������� = [1])") & vbNewLine & _
                " Group By b.No"
        Set rsInVoice = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
        If rsInVoice.RecordCount > 0 Then
            Do While Not rsInVoice.EOF
                lngRow = mshDetail.FindRow(NVL(rsInVoice!NO), , 0) '���ݺ�
                If lngRow > 0 And lngRow < mshDetail.Rows Then
                    For i = lngRow To mshDetail.Rows - 1
                        If mshDetail.TextMatrix(i, 0) = NVL(rsInVoice!NO) Then
                            mshDetail.TextMatrix(i, 1) = NVL(rsInVoice!����) '�������÷�Ʊ��
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
    '����:ҽ�����ռ�¼��ɫ�ж�
    '����:δ�˼�¼����True,�����˷Ѽ�¼����False
    '����:������
    '����:2014-08-13
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    strSql = _
        " Select 1" & vbNewLine & _
        " From ������ü�¼" & vbNewLine & _
        " Where Nvl(�۸񸸺�, ���) = [2] And NO = [1] And Mod(��¼����,10) = 1 Having" & vbNewLine & _
        "  Sum(���� * ����) = (Select ���� * ����" & vbNewLine & _
        "                       From ������ü�¼" & vbNewLine & _
        "                       Where Nvl(�۸񸸺�, ���) = [2] And NO = [1] And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum < 2)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, lngSN)
    CheckInsureDetail = Not rsTmp.EOF
End Function

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    Dim varData As Variant
    
    strHead = "���ݺ�,1,0|��Ʊ��,1,0|���,1,0|��������,1,0|������,1,0|�ѱ�,1,0|���,4,800|����,1,2000|��Ʒ��,1,2000|" & _
            "���,1,1200|��λ,4,500|����,7,800|����,7,1000|Ӧ�ս��,7,1000|ʵ�ս��,7,1000|ִ�п���,4,1000|����,4,1000|" & _
            "˵��,1,1800|������,4,750|ҽ�Ƹ��ʽ,4,1200|ժҪ,1,1500|��¼״̬,1,0"
    
    With mshDetail
        .HighLight = flexHighlightWithFocus
        .Redraw = False
        varData = Split(strHead, "|")
        .COLS = UBound(varData) + 1
        For i = 0 To UBound(varData)
            If Split(varData(i), ",")(0) = "ID" Then .ColHidden(i) = True
            .TextMatrix(0, i) = Split(varData(i), ",")(0)
            If .TextMatrix(0, i) = "������" Then
                If InStr(mstrPrivs, "��ʾ������") = 0 Then
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
        If .TextMatrix(1, .ColIndex("���ݺ�")) <> "" Then Call DetailSplitGroup
        For i = 1 To .Rows - 1
            If .IsSubtotal(i) = False Then
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 1 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 2 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbRed
                If Val(.TextMatrix(i, .ColIndex("��¼״̬"))) = 3 Then .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlue
                .RowHeight(i) = 300
                'ҽ�����ռ�¼����
'                If mshList.TextMatrix(mshList.Row, GetColNum("ҽ��")) <> "" And Val(mshList.TextMatrix(mshList.Row, GetColNum("����"))) = 3 And Val(mshList.TextMatrix(mshList.Row, GetColNum("�˷ѷ���"))) <> 2 Then
'                    If CheckInsureDetail(.TextMatrix(i, .ColIndex("���ݺ�")), Val(.TextMatrix(i, .ColIndex("���")))) Then
'                        .Cell(flexcpForeColor, i, 0, i, .COLS - 1) = vbBlack
'                    End If
'                End If
            End If
        Next i
        
        If gTy_System_Para.bytҩƷ������ʾ = 0 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = True
        End If
        If gTy_System_Para.bytҩƷ������ʾ = 1 Then
            .ColHidden(.ColIndex("����")) = True
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End If
        If gTy_System_Para.bytҩƷ������ʾ = 2 Then
            .ColHidden(.ColIndex("����")) = False
            .ColHidden(.ColIndex("��Ʒ��")) = False
        End If
    End With
End Sub

Private Sub ShowInvoice(ByVal strNo As String)
    Dim strSql As String, lngBalanceID As Long, blnOld As Boolean
    Dim rsInVoice As ADODB.Recordset
    
    On Error GoTo errH
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("�������")))
    blnOld = CheckBalance(lngBalanceID)
    If blnOld Then
        strSql = _
        "Select Distinct b.Id, b.���� As Ʊ�ݺ�," & vbNewLine & _
        "       Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��," & vbNewLine & _
        "       To_Char(b.ʹ��ʱ��, 'MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����" & vbNewLine & _
        "From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B, ������ü�¼ C" & vbNewLine & _
        "Where a.Id = b.��ӡid And a.No = c.No And a.�������� = 1 And b.Ʊ�� = 1 And c.����id = [1]" & vbNewLine & _
        "Order By ʹ��ʱ��"
    Else
        strSql = _
        "Select Distinct b.Id, b.���� As Ʊ�ݺ�," & vbNewLine & _
        "       Decode(b.ԭ��, 1, '��������', 2, '�����ջ�', 3, '�ش򷢳�', 4, '�ش��ջ�', 6, '��Ʊ����') As ʹ��ԭ��," & vbNewLine & _
        "       To_Char(b.ʹ��ʱ��, 'MM-DD HH24:MI') As ʹ��ʱ��, b.ʹ����" & vbNewLine & _
        "From Ʊ�ݴ�ӡ���� A, Ʊ��ʹ����ϸ B, ������ü�¼ C, ����Ԥ����¼ D" & vbNewLine & _
        "Where a.Id = b.��ӡid And a.No = c.No And c.����id = d.����id And a.�������� = 1 And b.Ʊ�� = 1 And d.������� = [1]" & vbNewLine & _
        "Order By ʹ��ʱ��"
    End If
    If mblnNOMoved Then
        strSql = Replace(strSql, "Ʊ�ݴ�ӡ����", "HƱ�ݴ�ӡ����")
        strSql = Replace(strSql, "Ʊ��ʹ����ϸ", "HƱ��ʹ����ϸ")
        strSql = Replace(strSql, "������ü�¼", "H������ü�¼")
        strSql = Replace(strSql, "����Ԥ����¼", "H����Ԥ����¼")
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
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("�������")))
    'Ԥ�����ʾ������롢ժҪ�����š�������ˮ�š�����˵��
    If CheckBalance(lngBalanceID) = False Then
        strSql = "Select Decode(Mod(a.��¼����,10),1,'��Ԥ���',Nvl(a.���㷽ʽ,'δ����')) As ���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��," & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.�������)) As �������, " & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.ժҪ)) As ժҪ, " & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.����)) As ����," & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.������ˮ��)) As ������ˮ��, " & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.����˵��)) As ����˵��" & _
                " From ����Ԥ����¼ A, (Select Distinct ����id From ����Ԥ����¼ Where ������� = [1]) B" & _
                " Where a.����id = b.����id" & _
                " Group By Decode(Mod(a.��¼����,10),1,'��Ԥ���',Nvl(a.���㷽ʽ,'δ����'))"
    Else
        strSql = "Select Decode(Mod(a.��¼����,10),1,'��Ԥ���',Nvl(a.���㷽ʽ,'δ����')) As ���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��," & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.�������)) As �������, " & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.ժҪ)) As ժҪ, " & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.����)) As ����," & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.������ˮ��)) As ������ˮ��, " & _
                "        Decode(Mod(Max(a.��¼����),10),1,'',Max(a.����˵��)) As ����˵��" & _
                " From ����Ԥ����¼ A" & _
                " Where a.����id = [1]" & _
                " Group By Decode(Mod(a.��¼����,10),1,'��Ԥ���',Nvl(a.���㷽ʽ,'δ����'))"
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
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("�������")))
    If tbPage.Selected.Index = 2 Then
        strSql = "Select Nvl(A.���㷽ʽ,'δ����') As ���㷽ʽ,Sum(A.��Ԥ��) As ��Ԥ��,Decode(Nvl(A.У�Ա�־,0),0,'��',2,'��','��') As ��־" & _
                " From ����Ԥ����¼ A" & _
                " Where A.������� = [1]" & _
                " Group By Nvl(A.���㷽ʽ,'δ����'),Nvl(A.У�Ա�־,0)" & _
                " Order By ��־"
    Else
        strSql = "Select Nvl(A.���㷽ʽ,'δ�˽��') As ���㷽ʽ,Sum(A.��Ԥ��) As ��Ԥ��,Decode(Nvl(A.У�Ա�־,0),0,'��',2,'��','��') As ��־" & _
                " From ����Ԥ����¼ A" & _
                " Where A.������� = [1]" & _
                " Group By Nvl(A.���㷽ʽ,'δ�˽��'),Nvl(A.У�Ա�־,0)" & _
                " Order By ��־"
    End If
    If CheckBalance(lngBalanceID) Then strSql = Replace(strSql, "�������", "����ID")
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
        .InsertItem 0, "���˵��ݼ�¼", picTemp.hWnd, 0
        .InsertItem 1, "�˷������¼", picTemp.hWnd, 0
        .InsertItem 2, "�շ��쳣��¼", picTemp.hWnd, 0
        .InsertItem 3, "�˷��쳣��¼", picTemp.hWnd, 0
        .Item(0).Selected = True
    End With
    
    With tbSub
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .InsertItem 1, "Ʊ����Ϣ", picSubInvoice.hWnd, 0
        .InsertItem 2, "������Ϣ", picSubBalance.hWnd, 0
        .InsertItem 3, "���������Ϣ", picExtendInfo.hWnd, 0
        .Item(0).Selected = True
    End With
End Sub

Private Sub SetBalanceList()
    Dim strHead As String, strTemp As String
    Dim i As Long, strAcc As String, j As Long
    Dim varData As Variant
    
    strHead = "���㷽ʽ,4,1000|���,7,1000|�������,4,1000|ժҪ,1,1200|����,1,1000|������ˮ��,1,1000|����˵��,1,1200"
    
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
            .TextMatrix(i, .ColIndex("���")) = Formatex(.TextMatrix(i, .ColIndex("���")), 6, , , 2)
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
    lngBalanceID = Val(mshList.TextMatrix(mshList.Row, GetColNum("�������")))
    
    '89448,��Ԥ����ʱʹ��ҽ�ƿ�����һ��ʹ�øñ�Ԥ���շ�ʱ����Ӧ���н��������Ϣ����¼����<>1��
    If CheckBalance(lngBalanceID) Then
        '�޽�����ŵ�����
        strSql = _
            " Select a.����id As ID, b.���㷽ʽ, c.����, b.��Ԥ�� As ���, a.������Ŀ As ��Ŀ, a.�������� As ����" & vbNewLine & _
            " From " & IIf(mblnNOMoved, "H", "") & "�������㽻�� A, " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ B, ҽ�ƿ���� C" & vbNewLine & _
            " Where b.����id = [1] And b.��¼���� <> 1 And a.����id = b.Id And b.�����id = c.Id(+) Order By ID"
    Else
        strSql = _
            " Select a.����id As ID, b.���㷽ʽ, c.����, b.��Ԥ�� As ���, a.������Ŀ As ��Ŀ, a.�������� As ����" & vbNewLine & _
            " From " & IIf(mblnNOMoved, "H", "") & "�������㽻�� A, " & IIf(mblnNOMoved, "H", "") & "����Ԥ����¼ B, ҽ�ƿ���� C" & vbNewLine & _
            " Where b.������� = [1] And b.��¼���� <> 1 And a.����id = b.Id And b.�����id = c.Id(+) Order By ID"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngBalanceID)
    Set vsfExtendInfo.DataSource = rsTemp
    If rsTemp.RecordCount = 0 Then
        'û�е��������׼�¼ʱ�����ط�ҳ
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

    strHead = "ID,1,0|���㷽ʽ,1,0|����,1,0|���,1,0|��Ŀ,1,1200|����,1,2000"
    
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
            If .ColKey(i) = "ID" Or .ColKey(i) = "���㷽ʽ" Or .ColKey(i) = "����" Or .ColKey(i) = "���" Then .ColHidden(i) = True
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
        .Subtotal flexSTNone, .ColIndex("ID"), .ColIndex("��Ŀ"), gstrDec, &H8000000F
        .SubtotalPosition = flexSTAbove

        .Outline .ColIndex("��Ŀ")
        .OutlineCol = .ColIndex("��Ŀ")
        
        For i = 1 To .Rows - 1
            .MergeRow(i) = False
            If .IsSubtotal(i) Then
                .IsCollapsed(i) = flexOutlineExpanded
                strTemp = .Cell(flexcpTextDisplay, i, 0)
                .RowHeight(i) = 350

                .Cell(flexcpText, i, .ColIndex("��Ŀ")) = strTemp

                 strTemp = .Cell(flexcpTextDisplay, i + 1, .ColIndex("���㷽ʽ"))
                 'strTemp = strTemp & Space(1) & .Cell(flexcpTextDisplay, i + 1, .ColIndex("����"))
                 strTemp = strTemp & "(" & Format(.Cell(flexcpTextDisplay, i + 1, .ColIndex("���")), gstrDec) & ")"

                 .MergeRow(i) = True
                 .MergeCells = flexMergeRestrictRows
                 .Cell(flexcpAlignment, i, .ColIndex("��Ŀ"), i, .ColIndex("��Ŀ")) = 1
                 
                 For j = 0 To .COLS - 1
                    If j <= .ColIndex("����") Then
                        If j >= .ColIndex("��Ŀ") Then
                            .Cell(flexcpText, i, j) = strTemp
                            .Cell(flexcpFontBold, i, j) = False
                        End If
                    End If
                 Next
            End If
        Next
        Call .AutoSize(.ColIndex("��Ŀ"))
        For j = 0 To .COLS - 1
            .MergeCol(j) = True
        Next
    End With
End Sub

Private Sub ShowApplyFactList(Optional ByVal strNo As String)
    Dim strSql As String, i As Long
    
    If gTy_Module_Para.bytƱ�ݷ������ <> 0 Then
        strSql = _
        " Select distinct B.ID,B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ��������') as ʹ��ԭ��," & _
        " To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
        " From " & IIf(mblnNOMoved, "H", "") & "Ʊ�ݴ�ӡ��ϸ A," & _
                IIf(mblnNOMoved, "H", "") & "Ʊ��ʹ����ϸ B " & _
        " Where A.Ʊ��=1 And A.Ʊ��=B.����" & _
        "             And B.Ʊ��=1 And A.NO=[1]" & _
        " Order by ID"
        On Error GoTo errH
        Set mrsFact = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
        If mrsFact.RecordCount = 0 Then GoTo ReadOld:
    Else
ReadOld:
        strSql = _
        " Select B.ID, B.���� as Ʊ�ݺ�,Decode(B.ԭ��,1,'��������',2,'�����ջ�',3,'�ش򷢳�',4,'�ش��ջ�',6,'��Ʊ��������') as ʹ��ԭ��," & _
        " To_Char(B.ʹ��ʱ��,'MM-DD HH24:MI') as ʹ��ʱ��,B.ʹ����" & _
        " From " & IIf(mblnNOMoved, "H", "") & "Ʊ�ݴ�ӡ���� A," & _
                IIf(mblnNOMoved, "H", "") & "Ʊ��ʹ����ϸ B" & _
        " Where A.��������=1 And A.ID=B.��ӡID" & _
        " And B.Ʊ��=1 And A.NO=[1]" & _
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
    strHead = "ID,1,0|Ʊ�ݺ�,1,850|ʹ��ԭ��,1,850|ʹ��ʱ��,1,1080|ʹ����,1,800"
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
        strHead = "���㷽ʽ,4,1000|�շѽ��,7,1000|�˷�״̬,4,1200"
    Else
        strHead = "���㷽ʽ,4,1000|�շѽ��,7,1000|�շ�״̬,4,1200"
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
            If .TextMatrix(i, 0) Like "*���*" Then
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

    strHead = "ID,1,0|Ʊ�ݺ�,4,1000|ʹ��ԭ��,4,1000|ʹ��ʱ��,4,1200|ʹ����,1,1000"
    
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
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Function zlGet����ID(ByVal strNo As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ID
    '����:���˺�
    '����:2011-04-29 17:05:13
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select Max(����ID) as ����ID From ������ü�¼ Where No=[1] and ��¼����=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    zlGet����ID = Val(NVL(rsTemp!����ID))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlGet����ID(ByVal strNo As String) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ����ID
    '����:���ϴ�
    '����:2015-09-25 17:05:13
    '����:81688
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errHandle
    
    strSql = "Select Min(����ID) as ����ID From ������ü�¼ Where No=[1] and ��¼����=1 And ��¼״̬ In(1,3) "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    zlGet����ID = Val(NVL(rsTemp!����ID))
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function zlMsgModuleInit() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣģ��
    '���:lngModule -ģ���
    '     strPivs-Ȩ�޴�
    '����:objMsgModule-������Ϣ����
    '����:��ʼ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-03-11 11:46:08
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
    '����:��ж��Ϣģ��
    '���:objMsgModule-��Ϣ����
    '����:���˺�
    '����:2014-03-11 11:46:08
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
