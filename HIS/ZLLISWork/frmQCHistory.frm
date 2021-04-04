VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmQCHistory 
   Caption         =   "历史质控查询"
   ClientHeight    =   8565
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   Icon            =   "frmQCHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11400
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picCalc 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   1812
      Left            =   9432
      ScaleHeight     =   1815
      ScaleWidth      =   2565
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5088
      Visible         =   0   'False
      Width           =   2568
      Begin VSFlex8Ctl.VSFlexGrid vfgCalc 
         Height          =   1260
         Left            =   72
         TabIndex        =   19
         Top             =   192
         Width           =   1932
         _cx             =   3408
         _cy             =   2222
         Appearance      =   2
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   Begin VB.ComboBox cbo科室 
      Height          =   300
      Left            =   2070
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   90
      Width           =   1845
   End
   Begin VB.PictureBox picReport 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   1824
      Left            =   3048
      ScaleHeight     =   1830
      ScaleWidth      =   3060
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4776
      Visible         =   0   'False
      Width           =   3060
      Begin VSFlex8Ctl.VSFlexGrid vfgReport 
         Height          =   672
         Left            =   60
         TabIndex        =   11
         Top             =   252
         Width           =   1656
         _cx             =   2921
         _cy             =   1185
         Appearance      =   2
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   120
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmQCHistory.frx":038A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmQCHistory.frx":0924
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin XtremeDockingPane.DockingPane dkpSub 
         Left            =   45
         Top             =   0
         _Version        =   589884
         _ExtentX        =   450
         _ExtentY        =   423
         _StockProps     =   0
         VisualTheme     =   5
      End
   End
   Begin VB.ComboBox cbo仪器 
      Height          =   300
      Left            =   4905
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   75
      Width           =   2115
   End
   Begin VB.PictureBox picCharts 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4395
      Left            =   4890
      ScaleHeight     =   4395
      ScaleWidth      =   6510
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   495
      Width           =   6510
      Begin XtremeSuiteControls.TabControl tbcCharts 
         Height          =   3975
         Left            =   150
         TabIndex        =   3
         Top             =   165
         Width           =   6105
         _Version        =   589884
         _ExtentX        =   10769
         _ExtentY        =   7011
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picRecord 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   6750
      Left            =   120
      ScaleHeight     =   6750
      ScaleWidth      =   2445
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   2445
      Begin VB.CommandButton cmd刷新 
         Height          =   600
         Left            =   2085
         Picture         =   "frmQCHistory.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   90
         Width           =   330
      End
      Begin MSComCtl2.DTPicker dtp日期 
         Height          =   300
         Index           =   0
         Left            =   435
         TabIndex        =   5
         Top             =   75
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   64356355
         CurrentDate     =   39110
      End
      Begin MSComCtl2.DTPicker dtp日期 
         Height          =   300
         Index           =   1
         Left            =   435
         TabIndex        =   6
         Top             =   390
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   64356355
         CurrentDate     =   39110
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgReagent 
         Height          =   1605
         Left            =   60
         TabIndex        =   8
         Top             =   3135
         Visible         =   0   'False
         Width           =   2430
         _cx             =   4286
         _cy             =   2831
         Appearance      =   2
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgItem 
         Height          =   4830
         Left            =   45
         TabIndex        =   15
         Top             =   720
         Width           =   2445
         _cx             =   4313
         _cy             =   8520
         Appearance      =   2
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
      Begin VB.Label lbl日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   420
         Width           =   180
      End
      Begin VB.Label lbl日期 
         BackStyle       =   0  'Transparent
         Caption         =   "日期"
         Height          =   180
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   135
         Width           =   405
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8190
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmQCHistory.frx":7710
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
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
   Begin C1Chart2D8.Chart2D chtCopy 
      Height          =   435
      Left            =   1260
      TabIndex        =   13
      Top             =   60
      Visible         =   0   'False
      Width           =   765
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   1349
      _ExtentY        =   767
      _StockProps     =   0
      ControlProperties=   "frmQCHistory.frx":7FA2
   End
   Begin VB.PictureBox picData 
      BackColor       =   &H00FDD6C6&
      BorderStyle     =   0  'None
      Height          =   1812
      Left            =   6288
      ScaleHeight     =   1815
      ScaleWidth      =   2565
      TabIndex        =   16
      Top             =   5016
      Visible         =   0   'False
      Width           =   2568
      Begin VSFlex8Ctl.VSFlexGrid vfgRecord 
         Height          =   1260
         Left            =   72
         TabIndex        =   17
         Top             =   192
         Width           =   1932
         _cx             =   3408
         _cy             =   2222
         Appearance      =   2
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16635590
         ForeColorSel    =   -2147483640
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   60
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmQCHistory.frx":8601
      Left            =   675
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmQCHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mColG  '质控品表列
    ID = 0: 选择: 批号: 质控品: 靶值: SD: 水平: 开始日期: 结束日期: 输入均值: 输入SD: 输入cv
End Enum
Private Enum mColL  '质控数据表列
    序号 = 0: ID: 日期: 次数
End Enum
Private Enum mColR  '质控报告表列
    ID = 0: 检验项目id: 标记: 日期: 标本号: 项目: 结果: 质控品: 水平
End Enum
Private Enum mTab   '质控图窗格:依次为统计数据(频数图)、Levey_Jennings图、Z_分数图、Youden图、累积和图、Monica图、Grubbs表格、Grubbs图表
    LJ = 0: FQ: ZS: YD: MN: CS: Grubbs: GS
End Enum

Private Enum mColC  '质控品表列
    ID = 0: 预设均值: 预设SD: 预设CV: 本月均值: 本月sd: 本月CV: 累计均值: 累计sd: 累计CV
End Enum


Const conPane_Record = 201
Const conPane_Charts = 202
Const conPane_Report = 203
Const conPane_Data = 204
Const conPane_Calc = 205
'-----------------------------------------------------
'窗体变量
'-----------------------------------------------------
Private mstrPrivs As String     '当前使用者权限串
Private mlngListWidth As Long   '列表窗体的设计宽度
Private mblnCusum As Boolean    '当前仪器是否应用累积和规则，决定是否提供累积和图形

Private mfrmRptTxt As frmQCTodayReport  '报告内容子窗体
Private mfrmChartFQ As frmQCChartFQ '数据统计窗格
Private mfrmChartLJ As frmQCChartLJ     'LJ控制图窗格
Private mfrmChartZS As frmQCChartZS     'Z-分数图窗格
Private mfrmChartYD As frmQCChartYD     'Youden图窗格
Private mfrmChartCS As frmQCChartCS     '累积和图窗格
Private mfrmChartMN As frmQCChartMN     'Monica图窗格

Private mfrmGrubbs As frmQCGrubbs      'Grubbs 表格
Private mfrmChartGS As frmQCChartGS    'Grubbs 图表
'-----------------------------------------------------
'临时变量
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrCustom As CommandBarControlCustom
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rptCol As ReportColumn
Dim rptRcd As ReportRecord
Dim RptItem As ReportRecordItem
Dim rptRow As ReportRow

Dim lngCount As Long
Private mblnShowAll As Boolean              '显示所有失控报告
Private mstr期间  As String                 '存期间的上下限
Private mEditMode As Integer                '编辑模式 0=非编辑 1=正在编辑
Private mstrPigeonhole As String            '归档人

Private mLastStartDate As Date, mLastEndDate As Date
Private mLastCell As String '焦点离开前的单元格，用于弃用与采信功能
Private mint显示失效记录 As Integer '0-不显示，1-显示
Private mintLJ图补位显示     As Integer '0-不补位, 1-补位 (默认)
Private Const ID_MENU_MOUSE = 90                                    '右键菜单
Private mlngItemID As Long                                          '当前选中的项目ID
Private mLastItemID As Long                                         '上次显示的项目ID，避免重复刷新
'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Private Function zlRefRecord() As Long
    '功能：刷新质控结果记录
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim date开始 As Date, date结束 As Date
    
    Err = 0: On Error GoTo ErrHand
    If mlngItemID = 0 Then Exit Function
    '|| '-' || Decode(Nvl(R.弃用结果, 0), 0, 999, R.弃用结果)
    gstrSql = "Select R.id,Q.检验时间 As 日期,Q.时间, To_Char(Q.测试次数, '000')  As 次数," & vbNewLine & _
            "       Q.质控品id, Zl_lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果," & vbNewLine & _
            "       Nvl(T.标记, 0) As 标记, Q.检验人,R.弃用结果" & vbNewLine & _
            "From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T" & vbNewLine & _
            "Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) /* And Nvl(R.是否检验, 0) = 1*/ And Q.仪器id + 0 = [1] And" & vbNewLine & _
            "      R.检验项目id + 0 = [2] And" & vbNewLine & _
            IIf(mint显示失效记录 = 1, "", "Nvl(R.弃用结果, 0)=0 And ") & _
            "      (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By Q.检验时间,  Q.测试次数, Q.质控品id"
            'Nvl(弃用结果, 0) * -1 +
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, _
                CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)), mlngItemID, _
                Format(Me.dtp日期(0).Value, "yyyy-MM-dd"), Format(Me.dtp日期(1).Value, "yyyy-MM-dd"))
    With Me.vfgRecord
        .Redraw = flexRDNone
        .Clear
        .FixedCols = 3
        .Cols = .FixedCols
        .ExtendLastCol = False '不自动扩展最后一列的宽度
        .Rows = 6 + Me.vfgReagent.Rows - 1
        
        .ColWidth(0) = 1200
        .TextMatrix(mColL.ID, 0) = "": .RowHidden(mColL.ID) = True
        .TextMatrix(mColL.序号, 0) = ""
        .TextMatrix(mColL.序号, 1) = "靶值": .ColWidth(1) = 500
        .TextMatrix(mColL.序号, 2) = "SD": .ColWidth(2) = 500
        
        .TextMatrix(mColL.日期, 0) = "日期" & vbNewLine & "时间"  ': .ColWidth(mColL.日期) = 1050
        .RowHeight(mColL.日期) = 600
        
        .TextMatrix(mColL.次数, 0) = "次数" ': .ColWidth(mColL.次数) = 600 ': .ColHidden(mColL.次数) = True
        .TextMatrix(.Rows - 2, 0) = "实际日期": .RowHidden(.Rows - 2) = True
        .TextMatrix(.Rows - 1, 0) = "检验人": .RowHidden(.Rows - 1) = True '.ColWidth(.Cols - 1) = 800
        For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
            strTemp = Split(Me.vfgReagent.TextMatrix(lngCount, mColG.批号), ", ")(0)
            .TextMatrix(mColL.ID, 0) = .TextMatrix(mColL.ID, 0) & "|" & strTemp & "=" & Me.vfgReagent.TextMatrix(lngCount, mColG.ID)
            .TextMatrix(lngCount + mColL.次数, 0) = strTemp
            .TextMatrix(lngCount + mColL.次数, 1) = Me.vfgReagent.TextMatrix(lngCount, mColG.靶值)
            .TextMatrix(lngCount + mColL.次数, 2) = Me.vfgReagent.TextMatrix(lngCount, mColG.SD)
            
            If Me.vfgReagent.Cell(flexcpChecked, lngCount, mColG.选择) = flexChecked Then
                '.ColWidth(lngCount + mColL.次数) = 900
                .RowHidden(lngCount + mColL.次数) = False
            Else
                '.ColWidth(lngCount + mColL.次数) = 0
                .RowHidden(lngCount + mColL.次数) = True
            End If
        Next
        .ColAlignment(0) = flexAlignLeftCenter
'        For lngCount = 0 To .Rows - 1
'            .FixedAlignment(lngCount) = flexAlignCenterCenter
'        Next
        Do While Not rsTemp.EOF
            lngRow = 0
            For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
                If rsTemp!质控品id = Val(Me.vfgReagent.TextMatrix(lngCount, mColG.ID)) Then
                    lngRow = lngCount + mColL.次数
                    date开始 = CDate(Me.vfgReagent.TextMatrix(lngCount, mColG.开始日期))
                    date结束 = CDate(Me.vfgReagent.TextMatrix(lngCount, mColG.结束日期))
                    Exit For
                End If
            Next
            If lngRow <> 0 Then
                lngCol = 0
                '按质控品的日期范围 显示数据
                If CDate(Format(rsTemp!日期, "yyyy-MM-dd")) >= date开始 And _
                   CDate(Format(rsTemp!日期, "yyyy-MM-dd")) <= date结束 Then
                    For lngCount = .FixedCols To .Cols - 1
                        If .TextMatrix(.Rows - 2, lngCount) = Format(rsTemp!日期, "yyyy-MM-dd") And _
                            .TextMatrix(mColL.次数, lngCount) = "" & rsTemp!次数 Then
                            lngCol = lngCount: Exit For
                        End If
                    Next
                    If lngCol = 0 Then
                        .Cols = .Cols + 1
                        lngCol = .Cols - 1
                        .ColWidth(lngCol) = 500
                        
                        .TextMatrix(mColL.序号, lngCol) = .Cols - .FixedCols
                        .TextMatrix(mColL.日期, lngCol) = Format(rsTemp!日期, "yy-MM-dd") & vbNewLine & Trim("" & rsTemp!时间)
                        
                        .TextMatrix(mColL.次数, lngCol) = "" & rsTemp!次数
                        .TextMatrix(.Rows - 2, lngCol) = Format(rsTemp!日期, "yyyy-MM-dd")
                        .TextMatrix(.Rows - 1, lngCol) = "" & rsTemp!检验人
                    Else
                        If InStr(1, .TextMatrix(.Rows - 1, lngCol), rsTemp!检验人) = 0 Then
                            .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 1, lngCol) & "," & rsTemp!检验人
                        End If
                    End If
                    .TextMatrix(mColL.ID, lngCol) = .TextMatrix(mColL.ID, lngCol) & "|" & Val("" & rsTemp!质控品id) & "=" & Val("" & rsTemp!ID)
                    .TextMatrix(lngRow, lngCol) = Trim("" & rsTemp!结果)
                    If Left(.TextMatrix(lngRow, lngCol), 1) = "." Then .TextMatrix(lngRow, lngCol) = "0" & .TextMatrix(lngRow, lngCol)
                    
                    Select Case Val("" & rsTemp!标记)
                    Case 1
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HC0FFFF
                        .Cell(flexcpFontBold, lngRow, lngCol) = True
                    Case 2
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HC0C0FF
                        .Cell(flexcpFontBold, lngRow, lngCol) = True
                    End Select
                    
                    '如是弃用结果则标为灰色
                    If Val("" & rsTemp!弃用结果) = 1 Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HC0C0C0
                        .Cell(flexcpFontBold, lngRow, lngCol) = True
                    End If
                End If '-- 在质控品指定日期范围，才显示
            End If
            rsTemp.MoveNext
        Loop
        If .Cols > .FixedCols Then
            .Cell(flexcpAlignment, mColL.序号, .FixedCols, mColL.日期, .Cols - 1) = flexAlignCenterCenter
            .AutoSize 0, .Cols - 1
        End If
        .Redraw = flexRDDirect
        If .Cols > .FixedCols Then .Col = .FixedCols: .Row = mColL.次数 + 1
    End With
    
    zlRefRecord = Me.vfgRecord.Cols - Me.vfgRecord.FixedCols
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlRefRecord = 0
End Function

Private Sub zlRefOthers()
    '功能：根据显示属性，刷新除质控记录外图形和报告
    Dim strLists As String, intLists As Integer
    Dim lngItemID As Long, strFromDate As String, strToDate As String
    Dim str选定的质控品 As String '调用绘图界面要用的质控品日期范围数据
    
    If mlngItemID = 0 Then Exit Sub
    
    If mlngItemID = mLastItemID Then Exit Sub
    If mlngItemID = -1 Then mlngItemID = mLastItemID
    mLastItemID = mlngItemID
    With Me.vfgReagent
        strLists = "": intLists = 0: str选定的质控品 = ""
        For lngCount = 0 To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mColG.选择) = flexChecked Then
                strLists = strLists & "," & .TextMatrix(lngCount, mColG.ID)
                str选定的质控品 = str选定的质控品 & ";" & .TextMatrix(lngCount, mColG.ID) & "=" & Format(CDate("" & .TextMatrix(lngCount, mColG.开始日期)), "yyyy-MM-dd") & "," & Format(CDate("" & .TextMatrix(lngCount, mColG.结束日期)), "yyyy-MM-dd")
                intLists = intLists + 1
            End If
        Next
        If strLists <> "" Then strLists = Mid(strLists, 2)
    End With
    lngItemID = mlngItemID
    strFromDate = Format(Me.dtp日期(0).Value, "yyyy-MM-dd")
    strToDate = Format(Me.dtp日期(1).Value, "yyyy-MM-dd")
    If str选定的质控品 <> "" Then str选定的质控品 = Mid(str选定的质控品, 2)
    '刷新质控报告
    If Me.dkpMan.FindPane(conPane_Report).Closed = False Then
        Call zlRefReport(strLists, lngItemID, strFromDate, strToDate)
    End If
        
    '获得当前选择的控制图，根据质控品变化，决定可显示的控制图形，并刷新数据
    Dim intSelTab As Integer
    For lngCount = 0 To Me.tbcCharts.ItemCount - 1
        If Me.tbcCharts.Item(lngCount).Selected Then intSelTab = lngCount: Exit For
    Next
    Me.tbcCharts.Item(mTab.FQ).Visible = (intLists > 0)
    Me.tbcCharts.Item(mTab.ZS).Visible = (intLists > 0)
    Me.tbcCharts.Item(mTab.YD).Visible = (intLists > 1)
    Me.tbcCharts.Item(mTab.CS).Visible = (intLists > 0 And mblnCusum)
    Me.tbcCharts.Item(mTab.MN).Visible = (intLists > 0)
    
    Me.tbcCharts.Item(mTab.Grubbs).Visible = (intLists > 0)
    Me.tbcCharts.Item(mTab.GS).Visible = (intLists > 0)
    
    If Me.tbcCharts.Item(intSelTab).Visible = False Then Me.tbcCharts.Item(mTab.LJ).Selected = True
    If Me.tbcCharts.Item(mTab.FQ).Selected Then Call mfrmChartFQ.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品)
    If Me.tbcCharts.Item(mTab.LJ).Selected Then Call mfrmChartLJ.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品, mintLJ图补位显示)
    If Me.tbcCharts.Item(mTab.ZS).Selected Then Call mfrmChartZS.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品)
    If Me.tbcCharts.Item(mTab.YD).Selected Then Call mfrmChartYD.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品)
    If Me.tbcCharts.Item(mTab.CS).Selected Then Call mfrmChartCS.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品)
    If Me.tbcCharts.Item(mTab.MN).Selected Then Call mfrmChartMN.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品)
    If Me.tbcCharts.Item(mTab.Grubbs).Selected Then Call mfrmGrubbs.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品)
    If Me.tbcCharts.Item(mTab.GS).Selected Then Call mfrmChartGS.zlRefresh(strLists, lngItemID, strFromDate, strToDate, str选定的质控品)
    
End Sub

Private Sub zlShowQCReport()
    '功能：加载未失控的质控数据
    Dim strLists As String, intLists As Integer
    Dim lngItemID As Long, strFromDate As String, strToDate As String
    Dim str选定的质控品 As String '调用绘图界面要用的质控品日期范围数据
    
    If mlngItemID = 0 Then Exit Sub
    
'    If mlngItemID = mLastItemID Then Exit Sub
    If mlngItemID = -1 Then mlngItemID = mLastItemID
    mLastItemID = mlngItemID
    With Me.vfgReagent
        strLists = "": intLists = 0: str选定的质控品 = ""
        For lngCount = 0 To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mColG.选择) = flexChecked Then
                strLists = strLists & "," & .TextMatrix(lngCount, mColG.ID)
                str选定的质控品 = str选定的质控品 & ";" & .TextMatrix(lngCount, mColG.ID) & "=" & Format(CDate("" & .TextMatrix(lngCount, mColG.开始日期)), "yyyy-MM-dd") & "," & Format(CDate("" & .TextMatrix(lngCount, mColG.结束日期)), "yyyy-MM-dd")
                intLists = intLists + 1
            End If
        Next
        If strLists <> "" Then strLists = Mid(strLists, 2)
    End With
    lngItemID = mlngItemID
    strFromDate = Format(Me.dtp日期(0).Value, "yyyy-MM-dd")
    strToDate = Format(Me.dtp日期(1).Value, "yyyy-MM-dd")
    If str选定的质控品 <> "" Then str选定的质控品 = Mid(str选定的质控品, 2)

    Call frmQCReport.ShowME(strLists, lngItemID, strFromDate, strToDate, Me)
        
   
End Sub


Public Sub zlRefReport(strResList As String, lngItemID, strFromDate As String, strToDate As String)
    '功能：刷新质控报告
    '参数： strResList  当前选择的质控品id串，以逗号分隔
    '       lngItemId   当前项目id
    '       strFromDate 开始日期
    '       strToDate   结束日期
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long
    
    Err = 0: On Error GoTo ErrHand
    '获取失控报告
    gstrSql = "Select R.ID,R.检验项目id, Nvl(T.标记, 0) As 标记, Q.检验时间 As 日期, Q.标本序号 As 标本号,D.中文名 ||'/'||英文名 as 项目, Zl_lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果," & vbNewLine & _
            "       M.批号 || ', ' || M.名称 As 质控品, M.水平, Q.检验人" & vbNewLine & _
            "From 检验质控记录 Q, 检验质控品 M, 检验普通结果 R, 检验质控报告 T,诊治所见项目 D" & vbNewLine & _
            "Where Q.质控品id = M.ID And Q.标本id = R.检验标本id And R.ID = T.结果id And Nvl(R.弃用结果,0)=0 And /*Nvl(R.是否检验, 0) = 1 And*/ " & vbNewLine & _
            "      Instr(',' || [1] || ',', ',' || Q.质控品id || ',') > 0 And R.检验项目id + 0 = D.ID" & IIf(mblnShowAll, "", " And R.检验项目id + 0 = [2]") & " And" & vbNewLine & _
            "      (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By Q.检验时间, R.排列序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList, lngItemID, strFromDate, strToDate)
    With Me.vfgReport
        .Redraw = flexRDNone
        
        .Clear
        
        If mblnShowAll Then
            .ToolTipText = "双击列表中的项目，可显示项目的质控数据。"
        Else
            .ToolTipText = "可在[查看]菜单中选择“显示所有失控报告”"
        End If
        Set .DataSource = rsTemp
        Call .AutoSize(mColR.标记, .Cols - 1)
        .ColWidth(mColR.ID) = 0: .ColHidden(mColR.ID) = True
        .ColWidth(mColR.检验项目id) = 0: .ColHidden(mColR.ID) = True
        .ColWidth(mColR.标记) = 280: .TextMatrix(0, mColR.标记) = ""
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        For lngCount = .FixedRows To .Rows - 1
            Select Case .TextMatrix(lngCount, mColR.标记)
                Case 1: Set .Cell(flexcpPicture, lngCount, mColR.标记) = Me.imgList.ListImages(1).Picture
                Case 2: Set .Cell(flexcpPicture, lngCount, mColR.标记) = Me.imgList.ListImages(2).Picture
            End Select
            .TextMatrix(lngCount, mColR.标记) = ""
            If Left(.TextMatrix(lngCount, mColR.结果), 1) = "." Then .TextMatrix(lngCount, mColR.结果) = "0" & .TextMatrix(lngCount, mColR.结果)
        Next
        .Redraw = flexRDDirect
        
        
        gstrSql = "select 报告人, 归档人 from 检验质控报告 where 结果id = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)))
        If rsTemp.EOF = False Then
            mstrPigeonhole = Trim(Nvl(rsTemp("归档人")))
        Else
            mstrPigeonhole = ""
        End If
        Call vfgReport_AfterRowColChange(.Row, .Col, .Row, .Col)
        
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlRefCalc()
    '功能：刷新质控结果记录
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim date开始 As Date, date结束 As Date
    
    
    Dim intFixeWidth As Integer
    
    Err = 0: On Error GoTo ErrHand
    If mlngItemID = 0 Then Exit Sub
    
    
    
    With Me.vfgCalc
        intFixeWidth = 1200
        .Redraw = flexRDNone
        .Cols = 10
        .Rows = 3
        .FixedRows = 2
        .ExtendLastCol = False '不自动扩展最后一列的宽度
        .MergeCells = flexMergeFree
        
        .ColHidden(mColC.ID) = True
        
        .TextMatrix(0, mColC.预设均值) = "预设": .ColWidth(mColC.预设均值) = intFixeWidth: .ColAlignment(mColC.预设均值) = flexAlignCenterCenter
        .TextMatrix(0, mColC.预设SD) = "预设": .ColWidth(mColC.预设SD) = intFixeWidth: .ColAlignment(mColC.预设SD) = flexAlignCenterCenter
        .TextMatrix(0, mColC.预设CV) = "预设": .ColWidth(mColC.预设CV) = intFixeWidth: .ColAlignment(mColC.预设CV) = flexAlignCenterCenter
        
        .TextMatrix(0, mColC.本月均值) = "本月": .ColWidth(mColC.本月均值) = intFixeWidth: .ColAlignment(mColC.本月均值) = flexAlignCenterCenter
        .TextMatrix(0, mColC.本月sd) = "本月": .ColWidth(mColC.本月sd) = intFixeWidth: .ColAlignment(mColC.本月sd) = flexAlignCenterCenter
        .TextMatrix(0, mColC.本月CV) = "本月": .ColWidth(mColC.本月CV) = intFixeWidth: .ColAlignment(mColC.本月CV) = flexAlignCenterCenter
        
        .TextMatrix(0, mColC.累计均值) = "到本月累计": .ColWidth(mColC.累计均值) = intFixeWidth: .ColAlignment(mColC.累计均值) = flexAlignCenterCenter
        .TextMatrix(0, mColC.累计sd) = "到本月累计": .ColWidth(mColC.累计sd) = intFixeWidth: .ColAlignment(mColC.累计sd) = flexAlignCenterCenter
        .TextMatrix(0, mColC.累计CV) = "到本月累计": .ColWidth(mColC.累计CV) = intFixeWidth: .ColAlignment(mColC.累计CV) = flexAlignCenterCenter
        
        
        .TextMatrix(1, mColC.预设均值) = "均值": .ColAlignment(mColC.预设均值) = flexAlignCenterCenter
        .TextMatrix(1, mColC.预设SD) = "SD": .ColAlignment(mColC.预设SD) = flexAlignCenterCenter
        .TextMatrix(1, mColC.预设CV) = "CV": .ColAlignment(mColC.预设CV) = flexAlignCenterCenter
        
        .TextMatrix(1, mColC.本月均值) = "均值": .ColAlignment(mColC.本月均值) = flexAlignCenterCenter
        .TextMatrix(1, mColC.本月sd) = "SD": .ColAlignment(mColC.本月sd) = flexAlignCenterCenter
        .TextMatrix(1, mColC.本月CV) = "CV": .ColAlignment(mColC.本月CV) = flexAlignCenterCenter
        
        .TextMatrix(1, mColC.累计均值) = "均值": .ColAlignment(mColC.累计均值) = flexAlignCenterCenter
        .TextMatrix(1, mColC.累计sd) = "SD": .ColAlignment(mColC.累计sd) = flexAlignCenterCenter
        .TextMatrix(1, mColC.累计CV) = "CV": .ColAlignment(mColC.累计CV) = flexAlignCenterCenter
        
        .Rows = Me.vfgReagent.Rows + 1
        For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
            .TextMatrix(lngCount + 1, mColC.预设均值) = Format(Me.vfgReagent.TextMatrix(lngCount, mColG.输入均值), "##0.00##")
            .TextMatrix(lngCount + 1, mColC.预设SD) = Format(Me.vfgReagent.TextMatrix(lngCount, mColG.输入SD), "##0.00##")
            .TextMatrix(lngCount + 1, mColC.预设CV) = Format(Round(Val(Me.vfgReagent.TextMatrix(lngCount, mColG.输入cv)) * 100, 4), "##0.00##")
            
            gstrSql = "Select Round(Avg(结果), 4) As 均值, Round(Stddev(结果), 4) As Sd, Count(*) As 次数" & vbNewLine & _
                "From (Select Trunc(Q.检验时间) As 日期," & vbNewLine & _
                "              Avg(zl_Lis_toNumber(Q.质控品ID,R.检验项目ID,R.检验结果,R.ID)) As 结果" & vbNewLine & _
                "       From 检验质控记录 Q, 检验普通结果 R,检验质控报告 T" & vbNewLine & _
                "       Where Q.标本id = R.检验标本id And Q.质控品id = [1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
                "             Nvl(R.弃用结果,0)=0 And R.ID=T.结果ID(+) And Q.检验时间 Between   [3] and [4]  And Nvl(T.标记, 0) <> 2" & vbNewLine & _
                "       Group By Trunc(Q.检验时间))"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReagent.TextMatrix(lngCount, mColG.ID)), mlngItemID, _
                            CDate(CStr(Format(Me.dtp日期(0), "yyyy-MM-dd"))), CDate(CStr(Format(Me.dtp日期(1), "yyyy-MM-dd"))))
                            
            .TextMatrix(lngCount + 1, mColC.本月均值) = IIf(Val(rsTemp("均值") & "") = 0, "", Format(Val(rsTemp("均值") & ""), "##0.00##"))
            .TextMatrix(lngCount + 1, mColC.本月sd) = IIf(Val(rsTemp("sd") & "") = 0, "", Format(Val(rsTemp("sd") & ""), "##0.00##"))
            If Val(rsTemp("Sd") & "") <> 0 And Val(rsTemp("均值") & "") <> 0 Then
                .TextMatrix(lngCount + 1, mColC.本月CV) = Format(Round(Val(rsTemp("Sd") & "") / Val(rsTemp("均值") & "") * 100, 2), "##0.00##")
            Else
                .TextMatrix(lngCount + 1, mColC.本月CV) = ""
            End If
            
            gstrSql = "Select Round(Avg(结果), 4) As 均值, Round(Stddev(结果), 4) As Sd, Count(*) As 次数" & vbNewLine & _
                "From (Select Trunc(Q.检验时间) As 日期," & vbNewLine & _
                "              Avg(zl_Lis_toNumber(Q.质控品ID,R.检验项目ID,R.检验结果,R.ID)) As 结果" & vbNewLine & _
                "       From 检验质控记录 Q, 检验普通结果 R,检验质控报告 T" & vbNewLine & _
                "       Where Q.标本id = R.检验标本id And Q.质控品id = [1] And R.检验项目id + 0 = [2] And" & vbNewLine & _
                "             Nvl(R.弃用结果,0)=0 And R.ID=T.结果ID(+) And Q.检验时间 < [3] And Nvl(T.标记, 0) <> 2" & vbNewLine & _
                "       Group By Trunc(Q.检验时间))"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReagent.TextMatrix(lngCount, mColG.ID)), mlngItemID, _
                            CDate(CStr(Format(Me.dtp日期(0), "yyyy-MM-dd"))), CDate(CStr(Format(Me.dtp日期(1), "yyyy-MM-dd"))))
                            
            .TextMatrix(lngCount + 1, mColC.累计均值) = IIf(Val(rsTemp("均值") & "") = 0, "", Format(Val(rsTemp("均值") & ""), "##0.00##"))
            .TextMatrix(lngCount + 1, mColC.累计sd) = IIf(Val(rsTemp("sd") & "") = 0, "", Format(Val(rsTemp("sd") & ""), "##0.00##"))
            If Val(rsTemp("Sd") & "") <> 0 And Val(rsTemp("均值") & "") <> 0 Then
                .TextMatrix(lngCount + 1, mColC.累计CV) = Format(Round(Val(rsTemp("Sd") & "") / Val(rsTemp("均值") & "") * 100, 2), "##0.00##")
            Else
                .TextMatrix(lngCount + 1, mColC.累计CV) = ""
            End If
        Next
        .MergeRow(0) = True
        .Redraw = flexRDDirect
    End With
    
    
'    '|| '-' || Decode(Nvl(R.弃用结果, 0), 0, 999, R.弃用结果)
'    gstrSql = "Select R.id,Q.检验时间 As 日期,Q.时间, To_Char(Q.测试次数, '000')  As 次数," & vbNewLine & _
'            "       Q.质控品id, Zl_lis_ToNumber(Q.质控品id,R.检验项目id,R.检验结果,R.id) As 结果," & vbNewLine & _
'            "       Nvl(T.标记, 0) As 标记, Q.检验人,R.弃用结果" & vbNewLine & _
'            "From 检验质控记录 Q, 检验普通结果 R, 检验质控报告 T" & vbNewLine & _
'            "Where Q.标本id = R.检验标本id And R.ID = T.结果id(+) /* And Nvl(R.是否检验, 0) = 1*/ And Q.仪器id + 0 = [1] And" & vbNewLine & _
'            "      R.检验项目id + 0 = [2] And" & vbNewLine & _
'            IIf(mint显示失效记录 = 1, "", "Nvl(R.弃用结果, 0)=0 And ") & _
'            "      (Q.检验时间 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
'            "Order By Q.检验时间,  Q.测试次数, Q.质控品id"
'            'Nvl(弃用结果, 0) * -1 +
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, _
'                CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)), mlngItemID, _
'                Format(Me.dtp日期(0).Value, "yyyy-MM-dd"), Format(Me.dtp日期(1).Value, "yyyy-MM-dd"))
'    With Me.vfgCalc
'        .Redraw = flexRDNone
'        .Clear
'        .FixedCols = 3
'        .Cols = .FixedCols
'        .ExtendLastCol = False '不自动扩展最后一列的宽度
'        .Rows = 6 + Me.vfgReagent.Rows - 1
'
'        .ColWidth(0) = 1200
'        .TextMatrix(mColL.ID, 0) = "": .RowHidden(mColL.ID) = True
'        .TextMatrix(mColL.序号, 0) = ""
'        .TextMatrix(mColL.序号, 1) = "靶值": .ColWidth(1) = 700
'        .TextMatrix(mColL.序号, 2) = "SD": .ColWidth(2) = 700
'
'        .TextMatrix(mColL.日期, 0) = "日期" & vbNewLine & "时间"  ': .ColWidth(mColL.日期) = 1050
'        .RowHeight(mColL.日期) = 600
'
'        .TextMatrix(mColL.次数, 0) = "次数" ': .ColWidth(mColL.次数) = 600 ': .ColHidden(mColL.次数) = True
'        .TextMatrix(.Rows - 2, 0) = "实际日期": .RowHidden(.Rows - 2) = True
'        .TextMatrix(.Rows - 1, 0) = "检验人": .RowHidden(.Rows - 1) = True '.ColWidth(.Cols - 1) = 800
'        For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
'            strTemp = Split(Me.vfgReagent.TextMatrix(lngCount, mColG.批号), ", ")(0)
'            .TextMatrix(mColL.ID, 0) = .TextMatrix(mColL.ID, 0) & "|" & strTemp & "=" & Me.vfgReagent.TextMatrix(lngCount, mColG.ID)
'            .TextMatrix(lngCount + mColL.次数, 0) = strTemp
'            .TextMatrix(lngCount + mColL.次数, 1) = Me.vfgReagent.TextMatrix(lngCount, mColG.靶值)
'            .TextMatrix(lngCount + mColL.次数, 2) = Me.vfgReagent.TextMatrix(lngCount, mColG.SD)
'
'
'        Next
'        .ColAlignment(0) = flexAlignLeftCenter
''        For lngCount = 0 To .Rows - 1
''            .FixedAlignment(lngCount) = flexAlignCenterCenter
''        Next
'
'
'        .Redraw = flexRDDirect
'
'    End With
    
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If Me.vfgRecord.Cols <= Me.vfgRecord.FixedCols Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgRecord
    objPrint.Title.Text = Mid(Me.cbo仪器.Text, InStr(1, Me.cbo仪器.Text, ",") + 1) & "质控结果清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cbo科室_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngMachineID As Long                '仪器ID
    
    On Error GoTo errH
    
    lngMachineID = Val(zlDatabase.GetPara("仪器", glngSys, 1209, 0))
    
    If Me.cbo科室.ListCount <= 0 Then Exit Sub
    
    If InStr(1, mstrPrivs, "所有科室") > 0 Then
        gstrSql = "Select Distinct D.ID, D.编码, D.名称, D.质控水平数" & vbNewLine & _
                "From 检验仪器 D, 检验质控品 M, 检验质控品项目 Q" & vbNewLine & _
                "Where D.ID = M.仪器id And M.ID = Q.质控品id And Nvl(D.微生物, 0) <> 1 and d.使用小组id = [1] " & vbNewLine
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo科室.ItemData(Me.cbo科室.ListIndex)))
    Else
'        gstrSql = "Select Distinct D.ID, D.编码, D.名称, D.质控水平数" & vbNewLine & _
                "From 检验仪器 D, 检验质控品 M, 检验质控品项目 Q" & vbNewLine & _
                "Where D.ID = M.仪器id And M.ID = Q.质控品id And Nvl(D.微生物, 0) <> 1 And" & vbNewLine & _
                "      D.使用小组id In (Select 部门id From 部门人员 Where 人员id = [1]) and d.使用小组id = [2] "
        gstrSql = "Select Distinct D.ID, D.编码, D.名称, D.质控水平数" & vbNewLine & _
                " From 检验仪器 D, 检验质控品 M, 检验质控品项目 Q" & vbNewLine & _
                " Where D.ID = M.仪器id And M.ID = Q.质控品id And Nvl(D.微生物, 0) <> 1 And D.使用小组id = [2] And" & vbNewLine & _
                "      D.ID In (Select Distinct D.ID" & vbNewLine & _
                "               From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
                "               Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [1] And C.仪器id = D.ID)"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(UserInfo.ID), CLng(Me.cbo科室.ItemData(Me.cbo科室.ListIndex)))
    End If
    
    With rsTemp
        Me.cbo仪器.Clear
        
        Do While Not .EOF
            Me.cbo仪器.AddItem !名称 & Space(200) & !质控水平数
            Me.cbo仪器.ItemData(Me.cbo仪器.NewIndex) = !ID
            If !ID = lngMachineID Then
                Me.cbo仪器.ListIndex = Me.cbo仪器.NewIndex
            End If
            .MoveNext
        Loop
'        If Me.cbo仪器.ListCount = 0 Then MsgBox "尚未完成仪器相关的质控设置！", vbInformation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo仪器.ListCount > 0 And cbo仪器.ListIndex = -1 Then
            Me.cbo仪器.ListIndex = 0
'            If Me.cbo仪器.ListCount = 1 Then Me.cbo仪器.Enabled = False
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------

Private Sub cbo仪器_Click()
    Dim lngItemID As Long   '项目ID

    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = Val(zlDatabase.GetPara("项目", glngSys, 1209, 0))
    
    mblnCusum = False
    If Me.cbo仪器.ListIndex = -1 Then Exit Sub
    Me.cbo仪器.Tag = Right(Me.cbo仪器.Text, 1)
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select Nvl(Count(*), 0) As 累积和" & vbNewLine & _
            "From 检验仪器规则 A, 检验质控规则 R" & vbNewLine & _
            "Where A.规则id = R.ID And A.性质 = '1' And R.种类 = 3 And A.仪器id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)))
    If rsTemp.RecordCount > 0 Then
        If rsTemp.Fields(0).Value > 0 Then mblnCusum = True
    End If
    
    gstrSql = "Select Distinct I.ID, I.编码, I.英文名, I.中文名" & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 Q, 诊治所见项目 I" & vbNewLine & _
            "Where M.ID = Q.质控品id And Q.项目id = I.ID And M.仪器id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)))
    
    If rsTemp.RecordCount <= 0 Then MsgBox "尚未完成仪器质控品设置！", vbInformation, gstrSysName: Unload Me: Exit Sub
    
    With Me.vfgItem
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        Set .DataSource = rsTemp
        .ColWidth(0) = 0
        .ColWidth(1) = 500
        .ColWidth(2) = 600
        .ColWidth(3) = 600
        .ColHidden(0) = True
        .AutoSize 1, 3
        .ColWidth(1) = 20
        .ExplorerBar = flexExSort
    End With
    Call vfgItem_RowColChange
        
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim rsTmp As New ADODB.Recordset
    Dim panThis As Pane
    On Error GoTo ErrHand
    '------------------------------------
    Select Case Control.ID
    
    Case conMenu_File_PrintSet
        Select Case Me.tbcCharts.Selected.Index
        Case mTab.FQ: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1209_" & mTab.FQ + 1, Me
        Case mTab.LJ: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1209_" & mTab.LJ + 1, Me
        Case mTab.ZS: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1209_" & mTab.ZS + 1, Me
        Case mTab.YD: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1209_" & mTab.YD + 1, Me
        Case mTab.CS: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1209_" & mTab.CS + 1, Me
        Case mTab.MN: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1209_" & mTab.MN + 1, Me
        Case mTab.GS: ReportPrintSet gcnOracle, glngSys, "ZL1_INSIDE_1209_" & mTab.GS + 1, Me
        End Select
    Case conMenu_File_Preview
        Select Case Me.tbcCharts.Selected.Index
        Case mTab.FQ: Call mfrmChartFQ.ChartPrint: Call PrintQC(mTab.FQ, False)
        Case mTab.LJ: Call mfrmChartLJ.ChartPrint: Call PrintQC(mTab.LJ, False, mfrmChartLJ.ChartPrint)
        Case mTab.ZS: Call mfrmChartZS.ChartPrint: Call PrintQC(mTab.ZS, False)
        Case mTab.YD: Call mfrmChartYD.ChartPrint: Call PrintQC(mTab.YD, False)
        Case mTab.CS: Call mfrmChartCS.ChartPrint: Call PrintQC(mTab.CS, False)
        Case mTab.MN: Call mfrmChartMN.ChartPrint: Call PrintQC(mTab.MN, False)
        
        Case mTab.GS: Call mfrmChartGS.ChartPrint: Call PrintQC(mTab.GS, False)
        Case mTab.Grubbs: Call mfrmGrubbs.ReportPrint(1)
        
        End Select
    Case conMenu_File_Print
        Select Case Me.tbcCharts.Selected.Index
        Case mTab.FQ: Call mfrmChartFQ.ChartPrint: Call PrintQC(mTab.FQ, True)
        Case mTab.LJ: Call mfrmChartLJ.ChartPrint: Call PrintQC(mTab.LJ, True, mfrmChartLJ.ChartPrint)
        Case mTab.ZS: Call mfrmChartZS.ChartPrint: Call PrintQC(mTab.ZS, True)
        Case mTab.YD: Call mfrmChartYD.ChartPrint: Call PrintQC(mTab.YD, True)
        Case mTab.CS: Call mfrmChartCS.ChartPrint: Call PrintQC(mTab.CS, True)
        Case mTab.MN: Call mfrmChartMN.ChartPrint: Call PrintQC(mTab.MN, True)
        Case mTab.GS: Call mfrmChartGS.ChartPrint: Call PrintQC(mTab.GS, True)
        
        Case mTab.Grubbs: Call mfrmGrubbs.ReportPrint(2)
        
        End Select
    Case conMenu_File_BatPrint: Call zlRptPrint(1)
    Case conMenu_Edit_Save
        If mEditMode = 0 Then
            Select Case Me.tbcCharts.Selected.Index
            Case mTab.FQ: Call mfrmChartFQ.ChartSaveAs
            Case mTab.LJ: Call mfrmChartLJ.ChartSaveAs
            Case mTab.ZS: Call mfrmChartZS.ChartSaveAs
            Case mTab.YD: Call mfrmChartYD.ChartSaveAs
            Case mTab.CS: Call mfrmChartCS.ChartSaveAs
            Case mTab.MN: Call mfrmChartMN.ChartSaveAs
            Case mTab.GS: Call mfrmChartGS.ChartSaveAs
            End Select
        Else
            Call mfrmRptTxt.zlEditSave
            mEditMode = 0
            cbrControl.Caption = "另存为"
            Me.cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_Untread                                   '取消
        mfrmRptTxt.zlEditCancel
        mEditMode = 0
        Set cbrControl = cbsThis.FindControl(, conMenu_Edit_Save, True, True)
        cbrControl.Caption = "另存为"
        Me.cbsThis.RecalcLayout
    Case conMenu_Edit_MarkMap
        Select Case Me.tbcCharts.Selected.Index
        Case mTab.FQ: Call mfrmChartFQ.ChartCopy
        Case mTab.LJ: Call mfrmChartLJ.ChartCopy
        Case mTab.ZS: Call mfrmChartZS.ChartCopy
        Case mTab.YD: Call mfrmChartYD.ChartCopy
        Case mTab.CS: Call mfrmChartCS.ChartCopy
        Case mTab.MN: Call mfrmChartMN.ChartCopy
        Case mTab.GS: Call mfrmChartGS.ChartCopy
        End Select
    Case conMenu_Edit_Adjust                                        '填写失控报告
        Set panThis = Me.dkpMan.FindPane(conPane_Report)
        panThis.Select
        Call mfrmRptTxt.ZlEditStart(Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)))
        mEditMode = 1
        Set cbrControl = cbsThis.FindControl(, conMenu_Edit_Save, True, True)
        cbrControl.Caption = "保存报告"
        Me.cbsThis.RecalcLayout
    Case conMenu_Edit_QCReport                                      '质控报告
        zlShowQCReport
    Case conMenu_Edit_Archive                                       '归档
        gstrSql = "select 归档人 from 检验质控报告 where 结果id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)))
        If rsTmp.EOF = False Then
            If Nvl(rsTmp("归档人")) = "" Then
                If MsgBox("真的要将当前失控报告归档吗？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                gstrSql = "Zl_检验质控报告_Archive(" & Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)) & ",0)"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                mstrPigeonhole = gstrDBUser
            Else
                If MsgBox("该失控报告已经归档，真的取消归档吗？", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                gstrSql = "Zl_检验质控报告_Archive(" & Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)) & ",1)"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                mstrPigeonhole = ""
            End If
        End If
        Call mfrmRptTxt.zlRefresh(Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)))
            
    Case conMenu_File_Exit: Unload Me
    
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Hide
        Me.vfgReagent.Visible = Not Me.vfgReagent.Visible
        Call picRecord_Resize
    Case conMenu_View_ShowAll
        mblnShowAll = Not mblnShowAll
        mlngItemID = -1 '强制刷新
        Call zlRefOthers
    Case conMenu_View_Append '是否补位显示LJ图
        If mintLJ图补位显示 = 0 Then
            mintLJ图补位显示 = 1
        Else
            mintLJ图补位显示 = 0
        End If
        mlngItemID = -1 '强制刷新
        Call zlRefOthers
    Case conMenu_View_Option

        Set panThis = Me.dkpMan.FindPane(conPane_Report)
        If panThis.Closed = False Then
            panThis.Close
        Else
            panThis.Select
            mlngItemID = -1 '强制刷新
            Call zlRefOthers
        End If
    Case conMenu_View_Refresh
        mLastItemID = 0
        Call RefreshData
    
    Case conMenu_Tool_Analyse
        Dim DateBegin As Date, dateEnd As Date
        If mlngItemID <= 0 Then Exit Sub
        With Me.vfgRecord
            If Not (.TextMatrix(.Rows - 2, .FixedCols) <> "" And IsDate(.TextMatrix(.Rows - 2, .FixedCols))) Then Exit Sub
            DateBegin = CDate(.TextMatrix(.Rows - 2, .FixedCols))
            dateEnd = CDate(.TextMatrix(.Rows - 2, .Cols - 1))
        End With
        If frmQCCompute.ShowME(Me, _
                Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex), mlngItemID, _
                DateBegin, _
                 CLng(Val("" & Me.vfgReagent.TextMatrix(Me.vfgReagent.Row, mColG.ID))), dateEnd) Then
            Call zlRefRecord
            mlngItemID = -1 '强制刷新
            Call zlRefOthers
        End If
    Case conMenu_Tool_Define
        If mlngItemID <= 0 Then Exit Sub
        If Me.vfgReagent.Rows - 1 > Me.vfgReagent.FixedRows Then
            If frmQCRedefine.ShowME(Me, _
                    Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex), mlngItemID, _
                    CDate(Format(Me.dtp日期(1).Value, "yyyy-MM-dd")), CLng(Val("" & Me.vfgReagent.TextMatrix(Me.vfgReagent.Row, mColG.ID)))) Then
                Call vfgItem_RowColChange
            End If
        Else
            If frmQCRedefine.ShowME(Me, Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex), mlngItemID, CDate(Format(Me.dtp日期(1).Value, "yyyy-MM-dd"))) Then
                Call vfgItem_RowColChange
            End If
        End If
        
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_Tool_SignNew
        '按日期输入
        If Me.cbo仪器.ListCount > 0 Then
            Call zlDatabase.SetPara("仪器", Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex), glngSys, 1209)
            frmQCAddData.ShowME mstrPrivs, Me
        End If

    Case conMenu_Tool_SignVerify
        '按次数输入
        If Me.cbo仪器.ListCount > 0 Then
            Call zlDatabase.SetPara("仪器", Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex), glngSys, 1209)
            frmQCAddData1.ShowME mstrPrivs, Me
        End If
        
    Case conMenu_Manage_Discard '弃用
        Call Discard_OR_Recall(mLastCell, 1)
    Case conMenu_Manage_Recall  '采信
        Call Discard_OR_Recall(mLastCell, 0)
    Case conMenu_Manage_Reset   '查看弃用原因
        Call Discard_OR_Recall(mLastCell, 2)
    Case conMenu_View_Jump      '显示失效数据
        If mint显示失效记录 = 0 Then
            mint显示失效记录 = 1
            Call zlRefRecord
        Else
            mint显示失效记录 = 0
            Call zlRefRecord
        End If
    Case conMenu_Tool_Reference_1
        '上
        Call ItemMoveUpDown(1)
    Case conMenu_Tool_Reference_2
        '下
        Call ItemMoveUpDown(2)
    Case Else
        If Control.ID < conMenu_ReportPopup * 100# + 1 Or Control.ID > conMenu_ReportPopup * 100# + 99 Then Exit Sub
        Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me)
    End Select
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    If Control.Type = xtpBarTypePopup Then
        Select Case Control.Index
        Case conMenu_EditPopup: Control.Visible = True
        End Select
    End If
    
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_BatPrint, conMenu_Edit_Save, conMenu_Edit_MarkMap: Control.Enabled = (Me.vfgRecord.Cols > Me.vfgRecord.FixedCols)
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Hide
        Control.Checked = Not Me.vfgReagent.Visible
    Case conMenu_View_Option
        Control.Checked = Me.dkpMan.FindPane(conPane_Report).Closed
    Case conMenu_Tool_Analyse
        Control.Enabled = (InStr(1, mstrPrivs, "计算") > 0 And Me.vfgRecord.Cols > Me.vfgRecord.FixedCols)
    Case conMenu_Tool_Define
        Control.Enabled = (InStr(1, mstrPrivs, "定值") > 0) And mlngItemID > 0
    Case conMenu_Tool_SignNew, conMenu_Tool_SignVerify
        Control.Enabled = (InStr(1, mstrPrivs, "质控记录输入") > 0)
    Case conMenu_View_ShowAll
        Control.Checked = mblnShowAll
    Case conMenu_Edit_Save
        If mEditMode = 1 Then
            Control.Caption = "保存报告"
        Else
            Control.Caption = "另存为"
        End If
    Case conMenu_Edit_Untread                           '取消
        Control.Enabled = (mEditMode = 1)
        Control.Visible = (mEditMode = 1)
    Case conMenu_Edit_Adjust                            '填写失控报告
        Control.Enabled = (mEditMode = 0 And mstrPigeonhole = "")
    Case conMenu_Edit_QCReport
        
    Case conMenu_Edit_Archive                           '归档
'        Control.Enabled = (mstrPigeonhole <> "")
    Case conMenu_Manage_Discard
        Control.Enabled = GetCellStat(mLastCell) = 1
    Case conMenu_Manage_Recall, conMenu_Manage_Reset
        Control.Enabled = GetCellStat(mLastCell) = 2
    Case conMenu_View_Jump
        Control.Checked = mint显示失效记录 = 1
    Case conMenu_View_Append
        Control.Checked = mintLJ图补位显示 = 1
    End Select
    
End Sub

Private Sub cmd刷新_Click()
    mLastItemID = 0
    Call RefreshData
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Record
        Item.Handle = Me.picRecord.hWnd
    Case conPane_Charts
        Item.Handle = Me.picCharts.hWnd
    Case conPane_Report
        Item.Handle = Me.picReport.hWnd
    Case conPane_Data
        Item.Handle = Me.picData.hWnd
    Case conPane_Calc
        Item.Handle = Me.picCalc.hWnd
    End Select
End Sub

Private Sub dkpMan_RClick(ByVal Pane As XtremeDockingPane.IPane)
    If Pane.ID = conPane_Data Then
        Me.picData.Visible = True
    ElseIf Pane.ID = conPane_Report Then
        Me.picReport.Visible = True
    End If
End Sub

Private Sub dkpSub_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Item.Handle = mfrmRptTxt.hWnd
End Sub

Private Sub dkpSub_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Err = 0: On Error Resume Next
    With Me.vfgReport
        .Left = Left + 60: .Width = Right - .Left
        .Top = Top + 60: .Height = Bottom - .Top * 2
    End With
End Sub

Private Sub RefreshData()
    Dim objControl As CommandBarControl
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset

    If mlngItemID = 0 Then Exit Sub
'    If Index = 0 Then
'        If Me.dtp日期(0).Value < DateAdd("m", -3, Me.dtp日期(1).Value) Then
'            '由于有需求提出三个的限期没有根据，超过三个月的查询情况很多暂时屏蔽
''            MsgBox "最大日期跨度不能超过三个月！", vbInformation, gstrSysName
''            Me.dtp日期(0).Value = DateAdd("m", -3, Me.dtp日期(1).Value)
'        End If
'    Else
'        If Me.dtp日期(0).Value > Me.dtp日期(1).Value Then Me.dtp日期(0).Value = Me.dtp日期(1).Value
'        If Me.dtp日期(1).Value > DateAdd("m", 3, Me.dtp日期(0).Value) Then
'            '由于有需求提出三个的限期没有根据，超过三个月的查询情况很多暂时屏蔽
''            MsgBox "最大日期跨度不能超过三个月！", vbInformation, gstrSysName
''            Me.dtp日期(1).Value = DateAdd("m", 3, Me.dtp日期(0).Value)
'        End If
'    End If
    Err = 0: On Error GoTo ErrHand
    
    If CDate(Format(Me.dtp日期(1).Value, "yyyy-MM-dd")) < CDate(Format(Me.dtp日期(0).Value, "yyyy-MM-dd")) Then
        MsgBox "结束日期不能大于开始日期！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSql = "Select M.ID, '' As 选择, M.批号 , M.名称 || ', 水平:' || M.水平 As 质控品, M.水平, to_Char(X.开始日期,'yy-MM-dd') as 开始日期,to_char(Nvl(X.结束日期, M.结束日期),'yy-MM-dd')  as 结束日期,输入均值,输入SD,输入cv," & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 I, 检验质控均值 X" & vbNewLine & _
            "Where M.ID = I.质控品id And I.质控品id = X.质控品id And I.项目id = X.项目id And M.仪器id = [1] And I.项目id = [2] And" & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.开始日期 And Nvl(X.结束日期, M.结束日期))" & vbNewLine & _
            "Order By M.开始日期, M.水平"
            
    gstrSql = "Select Id,选择,批号,质控品,Decode(substr(靶值,1,1),'.','0'||靶值,靶值) As 靶值,Decode(substr(SD,1,1),'.','0'||SD,SD) As SD,水平,min(开始日期) As 开始日期,Min(结束日期) As 结束日期,输入均值,输入SD,输入cv " & vbNewLine & _
            "From (" & vbNewLine & _
            "Select M.ID, '' As 选择, M.批号 , M.名称 || ', 水平:' || M.水平 As 质控品,I.靶值,I.SD, M.水平, to_Char(X.开始日期,'yy-MM-dd') as 开始日期,to_char(Nvl(X.结束日期, M.结束日期),'yy-MM-dd')  as 结束日期,x.均值 输入均值,x.sd 输入SD,x.cv 输入cv" & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 I, 检验质控均值 X" & vbNewLine & _
            "Where M.ID = I.质控品id And I.质控品id = X.质控品id And I.项目id = X.项目id And M.仪器id = [1] And I.项目id = [2] And" & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.开始日期 And Nvl(X.结束日期, M.结束日期))" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select M.ID, '' As 选择, M.批号 , M.名称 || ', 水平:' || M.水平 As 质控品,I.靶值,I.SD, M.水平, to_Char(X.开始日期,'yy-MM-dd') as 开始日期,to_char(Nvl(X.结束日期, M.结束日期),'yy-MM-dd')  as 结束日期,x.均值 输入均值,x.sd 输入SD,x.cv 输入cv" & vbNewLine & _
            "From 检验质控品 M, 检验质控品项目 I, 检验质控均值 X" & vbNewLine & _
            "Where M.ID = I.质控品id And I.质控品id = X.质控品id And I.项目id = X.项目id And M.仪器id = [1] And I.项目id = [2] And" & vbNewLine & _
            "        ( ( X.开始日期 Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd') )" & vbNewLine & _
            "         Or" & vbNewLine & _
            "          (nvl(X.结束日期,Sysdate) Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')+1-1/24*60*60)" & vbNewLine & _
            "         )" & vbNewLine & _
            "       )" & vbNewLine & _
            "Group By      Id,选择,批号,质控品,靶值,SD,水平,输入均值,输入SD,输入cv " & vbNewLine & _
            "Order By 质控品,水平"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, _
                CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex)), mlngItemID, _
                CStr(Format(Me.dtp日期(0), "yyyy-MM-dd")), CStr(Format(Me.dtp日期(1), "yyyy-MM-dd")))
    
    With Me.vfgReagent
        .FixedRows = 1
        Set .DataSource = rsTemp
        .ColWidth(mColG.选择) = 500
        .ColWidth(mColG.批号) = 900
        .ColWidth(mColG.ID) = 0: .ColWidth(mColG.水平) = 0
        .ColHidden(mColG.ID) = True: .ColHidden(mColG.水平) = True
        .ColHidden(mColG.靶值) = True: .ColHidden(mColG.SD) = True
        .ColHidden(mColG.输入均值) = True
        .ColHidden(mColG.输入SD) = True
        .ColHidden(mColG.输入cv) = True
        For lngCount = .FixedRows To .Rows - 1
'            If lngCount <= Val(Me.cbo仪器.Tag) Then
                .Cell(flexcpChecked, lngCount, mColG.选择) = flexChecked
'            Else
'                .Cell(flexcpChecked, lngCount, mColG.选择) = flexUnchecked
'            End If
        Next
        If .Rows <= .FixedRows Then
            '自动计算均值，SD，写入数据库,如为0，再提示
            MsgBox "还未定值,不能作图,请重新定值！", vbInformation, Me.Caption

        End If
    End With
    
    mLastStartDate = CDate(Format(dtp日期(0).Value, "yyyy-MM-dd"))
    mLastEndDate = CDate(Format(dtp日期(1).Value, "yyyy-MM-dd"))


    '刷新结果数据
    Call zlRefRecord
    Call zlRefCalc
    Call zlRefOthers
    Call picRecord_Resize
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim lngDeptID As Long  '科室ID
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gstrPrivs变化，导致控制无效
'
    '由于有技站要直接这个窗体所以重置一下脚本
    gstrPrivs = GetPrivFunc(100, 1209)
    mstrPrivs = gstrPrivs
    mlngListWidth = Me.picRecord.Width
    Me.picReport.BackColor = Me.BackColor
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    
    If Val(zlDatabase.GetPara("隐藏质控品", glngSys, 1209, 1)) = 1 Then
        Me.vfgReagent.Visible = False
    Else
        Me.vfgReagent.Visible = True
    End If
    
    lngDeptID = Val(zlDatabase.GetPara("科室", glngSys, 1209, 0))
    
    mblnShowAll = Val(zlDatabase.GetPara("显示所有失控项目", glngSys, 1209, 0)) = 1
    mint显示失效记录 = Val(zlDatabase.GetPara("显示失效数据", glngSys, 1209, 1))
    mintLJ图补位显示 = Val(zlDatabase.GetPara("LJ图补位显示", glngSys, 1209, 1))
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览控制图")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印控制图(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "打印质控结果(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "另存控制图(&S)...")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "复制控制图(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "显示失效数据(&V)"): cbrControl.BeginGroup = True
        'Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "隐藏质控报告(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Hide, "隐藏质控品选择(&H)"): cbrControl.Style = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "显示所有失控报告(&H)"):
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Append, "LJ图补位显示(&L)"):
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", -1, False)
    cbrMenuBar.ID = xtpControlPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "弃用结果(Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "采信结果(R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "查看(修改)弃用原因(V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "按日期输入质控记录(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "按次数输入质控记录(&C)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "失控计算(&Y)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "重新定值(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "失控报告"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCReport, "质控报告"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档")
        
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…"): cbrControl.BeginGroup = True
    End With
    
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "科室")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "科室")
    cbrCustom.Handle = Me.cbo科室.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "仪器")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "仪器")
    cbrCustom.Handle = Me.cbo仪器.hWnd: cbrCustom.Flags = xtpFlagRightAlign
'    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "项目")
'    cbrControl.Flags = xtpFlagRightAlign
'    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "项目")
'    cbrCustom.Handle = Me.cbo项目.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    
'    '右键菜单
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_MENU_MOUSE, "右键菜单", -1, False)
    cbrMenuBar.ID = ID_MENU_MOUSE
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "弃用结果(Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "采信结果(R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "查看(修改)弃用原因(V)"): cbrControl.BeginGroup = True

    End With
    cbrMenuBar.Visible = False
    
    '快键绑定
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    
        .Add 0, VK_UP, conMenu_Tool_Reference_1
        .Add 0, VK_DOWN, conMenu_Tool_Reference_2
    
    End With
    
    '设置不常用菜单
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_Edit_MarkMap
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbsThis, glngSys, glngModul, mstrPrivs)
    '-----------------------------------------------------
    '工具栏定义
    Set cbrToolBar = Me.cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "另存为")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "取消"): cbrControl.BeginGroup = True
        'Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "隐藏质控报告"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "失控计算"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "重新定值")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "失控报告"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCReport, "质控报告"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "归档")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '设置停靠窗格
    Dim panThis As Pane, panChild As Pane, panSub As Pane, panCalc As Pane
    
    With Me.dkpMan
        Set panThis = .CreatePane(conPane_Record, 200, 400, DockLeftOf, Nothing)
        panThis.Title = "质控结果表"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Set panThis = .CreatePane(conPane_Charts, 400, 500, DockRightOf, Nothing)
        panThis.Title = "质控统计图"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panChild = .CreatePane(conPane_Data, 400, 200, DockBottomOf, panThis)
        panChild.Title = "检验结果"
        panChild.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panSub = .CreatePane(conPane_Report, 400, 200, DockBottomOf)
        panSub.Title = "质控报告"
        panSub.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        panSub.AttachTo panChild
        
        Set panCalc = .CreatePane(conPane_Calc, 400, 200, DockBottomOf)
        panCalc.Title = "质控结果统计"
        panCalc.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        panCalc.AttachTo panChild
        
        panChild.Select
        
        .SetCommandBars Me.cbsThis
        .Options.ThemedFloatingFrames = True
        .Options.HideClient = True
    End With
    
    Set mfrmRptTxt = New frmQCTodayReport
    With Me.dkpSub
        Set panThis = .CreatePane(1, 400, 500, DockRightOf, Nothing)
        panThis.Title = "报告内容"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        .Options.ThemedFloatingFrames = True
        .Options.HideClient = False
    End With
    
    '-----------------------------------------------------
    '设置表格附加窗格
    Dim tbiThis As TabControlItem
    Set mfrmChartFQ = New frmQCChartFQ
    Set mfrmChartLJ = New frmQCChartLJ
    Set mfrmChartZS = New frmQCChartZS
    Set mfrmChartYD = New frmQCChartYD
    Set mfrmChartCS = New frmQCChartCS
    Set mfrmChartMN = New frmQCChartMN
    Set mfrmGrubbs = New frmQCGrubbs
    
    Set mfrmChartGS = New frmQCChartGS
    With Me.tbcCharts
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        Set tbiThis = .InsertItem(mTab.LJ, mfrmChartLJ.Caption, mfrmChartLJ.hWnd, 0)
        Set tbiThis = .InsertItem(mTab.FQ, mfrmChartFQ.Caption, mfrmChartFQ.hWnd, 0): tbiThis.Visible = False
        Set tbiThis = .InsertItem(mTab.ZS, mfrmChartZS.Caption, mfrmChartZS.hWnd, 0): tbiThis.Visible = False
        Set tbiThis = .InsertItem(mTab.YD, mfrmChartYD.Caption, mfrmChartYD.hWnd, 0): tbiThis.Visible = False
        Set tbiThis = .InsertItem(mTab.CS, mfrmChartCS.Caption, mfrmChartCS.hWnd, 0): tbiThis.Visible = False
        Set tbiThis = .InsertItem(mTab.MN, mfrmChartMN.Caption, mfrmChartMN.hWnd, 0): tbiThis.Visible = False
        Set tbiThis = .InsertItem(mTab.Grubbs, mfrmGrubbs.Caption, mfrmGrubbs.hWnd, 0): tbiThis.Visible = False
        
        Set tbiThis = .InsertItem(mTab.GS, mfrmChartGS.Caption, mfrmChartGS.hWnd, 0): tbiThis.Visible = False
    End With
    
    '-----------------------------------------------------
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    
    '-----------------------------------------------------
    '装入基本数据
    Dim rsTemp As New ADODB.Recordset
    
    Me.dtp日期(1).Value = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")): Me.dtp日期(0).Value = CDate(Format(Me.dtp日期(1).Value, "yyyy-MM") & "-01")
    Err = 0: On Error GoTo ErrHand
    
    If InStr(1, mstrPrivs, "所有科室") > 0 Then
        gstrSql = " Select Distinct b.Id, b.编码 , b.名称 As 科室 From 检验仪器 a ,部门表 b,检验质控品 c " & _
                  "Where a.使用小组ID = b.ID and a.id = c.仪器id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName)
        
    Else
        gstrSql = "Select Distinct B.ID, B.编码, B.名称 As 科室" & vbNewLine & _
                " From 检验仪器 A, 部门表 B, 检验质控品 C" & vbNewLine & _
                " Where A.使用小组id = B.ID And A.ID = C.仪器id And" & vbNewLine & _
                "      A.使用小组id In (Select Distinct D.使用小组id" & vbNewLine & _
                "                   From 检验小组成员 A, 检验小组 B, 检验小组仪器 C, 检验仪器 D" & vbNewLine & _
                "                   Where A.小组id = B.ID And B.ID = C.小组id　and 人员id = [1] And C.仪器id = D.ID)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, UserInfo.ID)
    End If
    
    Me.cbo科室.Clear
    Do Until rsTemp.EOF
        Me.cbo科室.AddItem rsTemp("编码") & "-" & rsTemp("科室")
        Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = rsTemp("Id")
        If rsTemp("ID") = lngDeptID Then
            Me.cbo科室.ListIndex = Me.cbo科室.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    If Me.cbo科室.ListCount = 0 Then MsgBox "尚未完成仪器使用小组的设置！", vbInformation, gstrSysName: Unload Me: Exit Sub
    If cbo科室.ListIndex = -1 Then
        Me.cbo科室.ListIndex = 0
    End If
    If Me.cbo科室.ListCount = 1 Then Me.cbo科室.Enabled = False
    
    mLastStartDate = CDate(0)
    mLastEndDate = CDate(0)
    
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim panThis As Pane
    If Me.WindowState = vbMinimized Then Exit Sub
    Set panThis = Me.dkpMan.FindPane(conPane_Record)
    panThis.MinTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
    Me.dkpMan.NormalizeSplitters
    panThis.MinTrackSize.SetSize mlngListWidth / Screen.TwipsPerPixelX, panThis.MinTrackSize.Height
    panThis.MaxTrackSize.SetSize Screen.Width / Screen.TwipsPerPixelX, panThis.MaxTrackSize.Height
    Me.dkpMan.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmRptTxt
    Unload mfrmChartFQ
    Unload mfrmChartLJ
    Unload mfrmChartZS
    Unload mfrmChartYD
    Unload mfrmChartCS
    Unload mfrmChartMN
    Set mfrmRptTxt = Nothing
    Set mfrmChartFQ = Nothing
    Set mfrmChartLJ = Nothing
    Set mfrmChartZS = Nothing
    Set mfrmChartYD = Nothing
    Set mfrmChartCS = Nothing
    Set mfrmChartMN = Nothing
    
    If Me.vfgReagent.Visible Then
        Call zlDatabase.SetPara("隐藏质控品", 0, glngSys, 1209)
    Else
        Call zlDatabase.SetPara("隐藏质控品", 1, glngSys, 1209)
    End If
    
    If Me.cbo科室.ListCount > 0 Then
        Call zlDatabase.SetPara("科室", Me.cbo科室.ItemData(Me.cbo科室.ListIndex), glngSys, 1209)
    End If
    If Me.cbo仪器.ListCount > 0 Then
        Call zlDatabase.SetPara("仪器", Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex), glngSys, 1209)
    End If
    If mlngItemID > 0 Then
        Call zlDatabase.SetPara("项目", mlngItemID, glngSys, 1209)
    End If
    Call zlDatabase.SetPara("显示所有失控项目", IIf(mblnShowAll, 1, 0), glngSys, 1209)
    Call zlDatabase.SetPara("显示失效数据", mint显示失效记录, glngSys, 1209)
    Call zlDatabase.SetPara("LJ图补位显示", mintLJ图补位显示, glngSys, 1209)
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picCalc_Resize()
    Err = 0: On Error Resume Next

    '数据列表
    With Me.vfgCalc
        .Left = Me.picCalc.ScaleLeft: .Width = Me.picCalc.ScaleWidth - .Left
        .Top = Me.picCalc.ScaleTop
        .Height = Me.picCalc.ScaleHeight - .Top
    End With
End Sub

Private Sub picCharts_Resize()
    Err = 0: On Error Resume Next
    With Me.tbcCharts
        .Left = Me.picCharts.ScaleLeft: .Width = Me.picCharts.ScaleWidth - .Left
        .Top = Me.picCharts.ScaleTop: .Height = Me.picCharts.ScaleHeight - .Top
    End With
End Sub

Private Sub picData_Resize()
    Err = 0: On Error Resume Next

    '数据列表
    With Me.vfgRecord
        .Left = Me.picData.ScaleLeft: .Width = Me.picData.ScaleWidth - .Left
        .Top = Me.picData.ScaleTop
        .Height = Me.picData.ScaleHeight - .Top
    End With
End Sub

Private Sub picRecord_Resize()
    Err = 0: On Error Resume Next
'    Me.cbo期间.Width = Me.picRecord.ScaleWidth - Me.cbo期间.Left
    Me.cmd刷新.Left = Me.picRecord.ScaleWidth - Me.cmd刷新.Width - 15
    Me.dtp日期(1).Width = Me.picRecord.ScaleWidth - Me.cmd刷新.Width - 15 - Me.dtp日期(1).Left - 15
    Me.dtp日期(0).Width = Me.dtp日期(1).Width

    
    With Me.vfgReagent
        .Left = Me.picRecord.ScaleLeft: .Width = Me.picRecord.ScaleWidth - .Left
        .Height = (.Rows + 1.5) * 300
        .Top = Me.picRecord.ScaleHeight - .Height
    End With
    
    '质控项目列表
    With Me.vfgItem
        .Left = Me.picRecord.ScaleLeft: .Width = Me.picRecord.ScaleWidth - .Left
        .Height = Me.picRecord.ScaleHeight - .Top - IIf(Me.vfgReagent.Visible, Me.vfgReagent.Height - 15, 0)
    End With
    
End Sub

Private Sub tbcCharts_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mlngItemID = -1 '强制刷新
    If Me.Visible Then Call zlRefOthers
End Sub

Private Sub vfgItem_RowColChange()

    
    If mLastStartDate <> CDate(0) And mLastEndDate <> CDate(0) Then
        Me.dtp日期(0) = CDate(Format(mLastStartDate, "yyyy-MM-dd"))
        Me.dtp日期(1) = CDate(Format(mLastEndDate, "yyyy-MM-dd"))
    
    Else
        Me.dtp日期(0) = CDate(Format(Now, "yyyy-MM-01"))
        Me.dtp日期(1) = CDate(Format(Now, "yyyy-MM-dd"))
    End If
    With Me.vfgItem
        If .Row >= .FixedRows Then
            If mlngItemID <> Val(.TextMatrix(.Row, 0)) Then
                mlngItemID = Val(.TextMatrix(.Row, 0))
                Call RefreshData
            End If
        End If
    End With
    
End Sub

Private Sub vfgReagent_DblClick()
    With Me.vfgReagent
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        If .Cell(flexcpChecked, .Row, mColG.选择) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, mColG.选择) = flexChecked
            'Me.vfgRecord.ColWidth(.Row + mColL.次数 + 1) = 900
            Me.vfgRecord.RowHidden(.Row + mColL.次数) = False
        Else
            .Cell(flexcpChecked, .Row, mColG.选择) = flexUnchecked
            'Me.vfgRecord.ColWidth(.Row + mColL.次数 + 1) = 0
            Me.vfgRecord.RowHidden(.Row + mColL.次数) = True
        End If
    End With
    Call zlRefOthers
End Sub

Private Sub vfgReagent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfgReagent_DblClick
End Sub

Private Sub vfgRecord_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow <= mColL.次数 Or NewRow >= Me.vfgRecord.Rows - 2 Then
        Cancel = True
    Else
        If NewRow - mColL.次数 - 1 >= 0 And NewRow - mColL.次数 <= Me.vfgReagent.Rows - 1 Then
            On Error Resume Next
            Me.vfgReagent.Row = NewRow - mColL.次数
        End If
    End If
End Sub

Private Sub vfgRecord_EnterCell()
    With vfgRecord
        mLastCell = .Row & "," & .Col
    End With
End Sub

Private Sub vfgRecord_LeaveCell()
    With vfgRecord
        mLastCell = .Row & "," & .Col
    End With
End Sub

Private Sub vfgRecord_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    On Error Resume Next
    If Button = 2 Then
        If vfgRecord.Cols <= 1 Then Exit Sub
        If GetCellStat(mLastCell) <> 0 Then
            Set objPopup = cbsThis.ActiveMenuBar.FindControl(, ID_MENU_MOUSE)
            If Not objPopup Is Nothing Then
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
End Sub

Private Sub vfgRecord_RowColChange()
    With vfgRecord
        mLastCell = .Row & "," & .Col
    End With
End Sub

Private Sub vfgReport_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mfrmRptTxt.zlRefresh(Val(Me.vfgReport.TextMatrix(NewRow, mColR.ID)))
End Sub

Private Sub PrintQC(intPrintType As Integer, blnPrintMode As Boolean, Optional ByVal ReportCount As Integer = 1)
    '打印或预览质控图
    '参数           intPrintMode =1 打印 =2 预览
    '               intPrintType 0=LJ 1=FQ 2=ZS 3=YD 4=CS 5=MN
    
    Dim rsTmp As New ADODB.Recordset
    Dim strPrintType As String                  '对应的单据
    Dim strQCID As String                       '质控品ID可能会是以","分隔的多个ID
    Dim lngQCID As Long                         '单个质控品ID
    Dim lngItemID As String                     '项目ID
    Dim lngMachine As Long                      '仪器ID
    Dim intLoop As Integer
    
    
    On Error GoTo errH
    
    strPrintType = "ZL1_INSIDE_1209_"
    strPrintType = strPrintType & intPrintType + 1
    
    gstrSql = "Select b.w, b.h " & vbNewLine & _
                " From Zlreports a, Zlrptitems b" & vbNewLine & _
                " Where a.Id = b.报表id And a.编号 = [1] And b.名称 = '质控图'"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strPrintType)
    '没有找到时退出
    If rsTmp.EOF Then
        MsgBox "在单据定义中没有定义<质控图>,请在单据中定义一个名为<质控图>的图像框!", vbQuestion, Me.Caption
        Exit Sub
    End If
    
    For intLoop = 0 To ReportCount - 1
        If Dir(App.path & "\QC_Tmp" & intLoop) <> "" Then
        With Me.chtCopy
            .Load App.path & "\QC_Tmp" & intLoop
            Kill App.path & "\QC_Tmp" & intLoop
            .Width = Nvl(rsTmp("w"), 1280 * Screen.TwipsPerPixelX)
            .Height = Nvl(rsTmp("h"), 500 * Screen.TwipsPerPixelY)
            .Header.Text = ""
            .ChartLabels.RemoveAll
            .ChartArea.Location.Top = -5
            .ChartArea.Location.Height = .ChartArea.Location.Height + 15
            If intPrintType = 3 Then
                .ChartArea.Location.Left = 30
            End If
            .SaveImageAsJpeg App.path & "\QC" & intLoop & ".jpg", 1000, False, False, False
        End With
        End If
    Next
    
    
    '得到质控品ID
    Select Case intPrintType
        Case mTab.LJ
            lngQCID = mfrmChartLJ.ZLGetLJ_QCID
            With Me.vfgReagent
                strQCID = ""
                For lngCount = 0 To .Rows - 1
                    If .Cell(flexcpChecked, lngCount, mColG.选择) = flexChecked Then
                        strQCID = strQCID & "," & .TextMatrix(lngCount, mColG.ID)
                    End If
                Next
                If strQCID <> "" Then strQCID = Mid(strQCID, 2)
            End With
        Case mTab.FQ
            With Me.vfgReagent
                strQCID = ""
                For lngCount = 0 To .Rows - 1
                    If .Cell(flexcpChecked, lngCount, mColG.选择) = flexChecked Then
                        strQCID = strQCID & "," & .TextMatrix(lngCount, mColG.ID)
                    End If
                Next
                If strQCID <> "" Then strQCID = Mid(strQCID, 2)
            End With
            lngQCID = mfrmChartFQ.ZLGetFQ_QCID
        Case mTab.ZS
'            lngQCID = mfrmChartZS.ZLGetzs_QCID
        Case mTab.MN
            lngQCID = mfrmChartMN.ZLGetMN_QCID
        Case mTab.CS
            lngQCID = mfrmChartCS.ZLGetCS_QCID
        Case mTab.GS
            lngQCID = mfrmChartGS.ZLGetGS_QCID
    End Select
    
    '得到项目ID
    If mlngItemID = 0 Then Exit Sub
    lngItemID = mlngItemID
    lngMachine = CLng(Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex))
    
    If Dir(App.path & "\QC0.jpg") <> "" Then
        Call ReportOpen(gcnOracle, glngSys, strPrintType, Me, "质控图=" & App.path & "\QC0.jpg", _
        "质控品ID=" & lngQCID, "项目ID=" & lngItemID, "开始日期=" & Format(dtp日期(0), "yyyy-MM-dd"), "结束日期=" & Format(dtp日期(1), "yyyy-MM-dd"), _
        "仪器ID=" & lngMachine, "质控品组=" & IIf(strQCID = "", "0", strQCID), _
        "质控图1=" & App.path & "\QC1.jpg", "质控图2=" & App.path & "\QC2.jpg", _
        IIf(blnPrintMode, 2, 1))
    End If
    
    If Dir(App.path & "\QC*.jpg") <> "" Then Kill App.path & "\QC*.jpg"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vfgReport_DblClick()
    Dim lngItemID As Long, intLoop As Integer
    lngItemID = Val(Me.vfgReport.TextMatrix(vfgReport.Row, mColR.检验项目id))
    
    If lngItemID <> mlngItemID Then
        For intLoop = Me.vfgItem.FixedRows To Me.vfgItem.Rows - 1
            If lngItemID = Val(Me.vfgItem.TextMatrix(intLoop, 0)) Then
                Me.vfgItem.Row = intLoop
                Call vfgItem_RowColChange
                Exit For
            End If
        Next
    End If
End Sub

Private Function Discard_OR_Recall(ByVal strCell As String, ByVal intType As Integer) As Boolean
    '弃用或采信
    'intType : 0-采信  1-弃用  2-查看弃用结果
    Dim lngRow As Long, lngCol As Long, lngID As Long, strTmp As String
    Dim str所有质控品 As String, str当前质控品 As String, lng_S As Long, lng_E As Long
    Dim strSQL As String, str原因 As String
    Dim frmDiscard As New frmQCDiscardEdit
    On Error GoTo errH
    
    If InStr(strCell, ",") > 0 Then
        lngRow = Val(Split(strCell, ",")(0))
        lngCol = Val(Split(strCell, ",")(1))
        
        With vfgRecord
            If Not (lngCol >= .FixedCols And lngCol < .Cols And lngRow > 3 And lngRow <= .Rows - 2) Then
                MsgBox "请选择一个数据单元格后再使用此功能！"
                Exit Function
            End If
            If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then
                str所有质控品 = .TextMatrix(mColL.ID, 0)
                str当前质控品 = .TextMatrix(lngRow, 0)
                lng_S = InStr(str所有质控品, "|" & str当前质控品 & "=")
                If lng_S > 0 Then
                    lng_E = InStr(lng_S + 1, str所有质控品, "|")
                    If lng_E > lng_S Then
                        str当前质控品 = Mid(str所有质控品, lng_S, lng_E - lng_S)
                    Else
                        str当前质控品 = Mid(str所有质控品, lng_S)
                    End If
                    str当前质控品 = Split(str当前质控品, "=")(1)
                End If
                
                strTmp = .TextMatrix(mColL.ID, lngCol)
                lng_S = InStr(strTmp, "|" & str当前质控品 & "=")
                If lng_S > 0 Then
                    lng_E = InStr(lng_S + 1, strTmp, "|")
                    
                    If lng_E > lng_S Then
                        str当前质控品 = Mid(strTmp, lng_S, lng_E - lng_S)
                    Else
                        str当前质控品 = Mid(strTmp, lng_S)
                    End If
                    lngID = Val(Split(str当前质控品, "=")(1))
                End If
                
                If lngID > 0 Then
                    If intType = 1 Then
                        If frmDiscard.ShowME(lngID, str原因, Me) Then
                            strSQL = "zl_检验普通结果_弃用(" & lngID & ",1,'" & str原因 & "')"
                            zlDatabase.ExecuteProcedure strSQL, Me.Caption
                            .Cell(flexcpBackColor, lngRow, lngCol) = &HC0C0C0
                            Discard_OR_Recall = True
                        End If
                    ElseIf intType = 0 Then
                        strSQL = "zl_检验普通结果_弃用(" & lngID & ",0)"
                        zlDatabase.ExecuteProcedure strSQL, Me.Caption
                        .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor
                        Discard_OR_Recall = True
                    Else
                        If frmDiscard.ShowME(lngID, str原因, Me) Then
                            strSQL = "zl_检验普通结果_弃用(" & lngID & ",2,'" & str原因 & "')"
                            zlDatabase.ExecuteProcedure strSQL, Me.Caption
                            Discard_OR_Recall = True
                        End If
                    End If
                    Set frmDiscard = Nothing
                    If Discard_OR_Recall Then
                        Call zlRefRecord
                        .Select lngRow, lngCol
                        mlngItemID = -1 '强制刷新
                        Call zlRefOthers
                    End If
                    Exit Function
                End If
            Else
                MsgBox "请选择一个非空的数据单元格！"
                Exit Function
            End If
        End With
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCellStat(ByVal strCell As String) As Integer
    '取单元格状态,用于控制菜单
    '返回值  0-禁用弃用相关功能 1-可以弃用 2-可以启用
    Dim lngRow As Long, lngCol As Long
    GetCellStat = -1
    If InStr(strCell, ",") > 0 Then
        lngRow = Val(Split(strCell, ",")(0))
        lngCol = Val(Split(strCell, ",")(1))
        
        With vfgRecord
            If Not (lngCol >= .FixedCols And lngCol < .Cols And lngRow > 3 And lngRow <= .Rows - 2) Then
                GetCellStat = 0
                Exit Function
            End If
            
            If .Cell(flexcpBackColor, lngRow, lngCol) = &HC0C0C0 Then
                GetCellStat = 2
            Else
                GetCellStat = 1
            End If
        End With
    End If
End Function

Private Sub ItemMoveUpDown(ByVal intUpDown As Integer)
    '上下键处理
    On Error Resume Next
    With Me.vfgItem
        If intUpDown = 1 Then
            If .Row - 1 > .FixedRows Then .Select .Row - 1, .Col
        Else
            If .Row + 1 < .Rows Then .Select .Row + 1, .Col
        End If
    End With
End Sub
