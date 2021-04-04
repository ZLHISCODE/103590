VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CODEJO~3.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDockInAdvice 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsfAdviceColor 
      Height          =   1215
      Left            =   240
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
      _cx             =   2990
      _cy             =   2143
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   11
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.Timer timHide 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3240
      Top             =   6285
   End
   Begin VB.PictureBox picAppend 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   3360
      ScaleHeight     =   4455
      ScaleWidth      =   4560
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2400
      Width           =   4560
      Begin VB.PictureBox picBlood 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   3225
         ScaleHeight     =   390
         ScaleWidth      =   1065
         TabIndex        =   23
         Top             =   1485
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Timer TimShow 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   285
            Top             =   30
         End
         Begin VB.Timer timBRefresh 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   720
            Top             =   15
         End
         Begin XtremeDockingPane.DockingPane DkpBlood 
            Left            =   45
            Top             =   15
            _Version        =   589884
            _ExtentX        =   450
            _ExtentY        =   423
            _StockProps     =   0
         End
      End
      Begin VB.Frame fraExecUD 
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   150
         MousePointer    =   7  'Size N S
         TabIndex        =   19
         Top             =   1920
         Visible         =   0   'False
         Width           =   6000
      End
      Begin VB.PictureBox picExec 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   90
         ScaleHeight     =   405
         ScaleWidth      =   7080
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1485
         Visible         =   0   'False
         Width           =   7080
         Begin XtremeCommandBars.CommandBars cbsExec 
            Left            =   120
            Top             =   30
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAppend 
         Bindings        =   "frmDockInAdvice.frx":0000
         Height          =   1395
         Left            =   75
         TabIndex        =   20
         Top             =   75
         Width           =   7110
         _cx             =   12541
         _cy             =   2461
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
      Begin VSFlex8Ctl.VSFlexGrid vsExec 
         Height          =   885
         Left            =   75
         TabIndex        =   21
         Top             =   1995
         Visible         =   0   'False
         Width           =   7125
         _cx             =   12568
         _cy             =   1561
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
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
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   2
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
   Begin VB.ComboBox cboTime 
      Height          =   300
      Left            =   6750
      Style           =   2  'Dropdown List
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   975
      Width           =   1305
   End
   Begin VB.PictureBox PicAdviceDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEFEF&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   75
      ScaleHeight     =   2745
      ScaleWidth      =   2775
      TabIndex        =   10
      Top             =   7545
      Visible         =   0   'False
      Width           =   2800
      Begin VSFlex8Ctl.VSFlexGrid vsfAdivceDetail 
         Height          =   2475
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   2745
         _cx             =   4851
         _cy             =   4366
         Appearance      =   2
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
         BackColor       =   16773103
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16773103
         ForeColorSel    =   -2147483640
         BackColorBkg    =   16773103
         BackColorAlternate=   16773103
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockInAdvice.frx":0028
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         WallPaper       =   "frmDockInAdvice.frx":0066
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   30
      ScaleHeight     =   5490
      ScaleWidth      =   7260
      TabIndex        =   0
      Top             =   135
      Width           =   7260
      Begin VB.Frame fraMore 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4785
         TabIndex        =   2
         Top             =   15
         Visible         =   0   'False
         Width           =   225
         Begin VB.Image imgMore 
            Height          =   225
            Left            =   0
            Picture         =   "frmDockInAdvice.frx":12998
            Top             =   0
            Width           =   225
         End
      End
      Begin VB.Frame fraColSel 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6090
         TabIndex        =   4
         Top             =   225
         Width           =   195
         Begin VB.Image imgColSel 
            Height          =   195
            Left            =   0
            Picture         =   "frmDockInAdvice.frx":12D99
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
         Height          =   3555
         Left            =   465
         TabIndex        =   7
         Top             =   165
         Width           =   5925
         _cx             =   10451
         _cy             =   6271
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
         MouseIcon       =   "frmDockInAdvice.frx":132E7
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16444122
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   250
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmDockInAdvice.frx":13BC1
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   -1  'True
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox pictmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   8
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame fraAdviceUD 
         BorderStyle     =   0  'None
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   3
         Top             =   3840
         Width           =   6000
      End
      Begin VB.Frame fraHide 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   75
         Left            =   6945
         TabIndex        =   1
         ToolTipText     =   "鼠标停留时,过滤条件栏会自动显示"
         Top             =   15
         Visible         =   0   'False
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsColumn 
         Height          =   3270
         Left            =   6090
         TabIndex        =   5
         Top             =   465
         Visible         =   0   'False
         Width           =   1470
         _cx             =   2593
         _cy             =   5768
         Appearance      =   0
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmDockInAdvice.frx":13C5C
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         Editable        =   2
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
      Begin XtremeSuiteControls.TabControl tbcAppend 
         Height          =   2775
         Left            =   30
         TabIndex        =   6
         Top             =   3915
         Width           =   1350
         _Version        =   589884
         _ExtentX        =   2381
         _ExtentY        =   2275
         _StockProps     =   64
      End
      Begin XtremeCommandBars.CommandBars cbsSub 
         Left            =   30
         Top             =   90
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
      End
   End
   Begin XtremeSuiteControls.TabControl tbcMain 
      Bindings        =   "frmDockInAdvice.frx":13CAA
      Height          =   435
      Left            =   7545
      TabIndex        =   9
      Top             =   75
      Width           =   390
      _Version        =   589884
      _ExtentX        =   688
      _ExtentY        =   767
      _StockProps     =   64
   End
   Begin RichTextLib.RichTextBox rtfInfo 
      Height          =   900
      Left            =   3600
      TabIndex        =   12
      Top             =   6795
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockInAdvice.frx":13CBE
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
   Begin RichTextLib.RichTextBox rtfAppend 
      Height          =   900
      Left            =   3180
      TabIndex        =   13
      Top             =   6795
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockInAdvice.frx":13D5B
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
   Begin RichTextLib.RichTextBox rtfSche 
      Height          =   900
      Left            =   4035
      TabIndex        =   15
      Top             =   6795
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockInAdvice.frx":13DF8
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
   Begin RichTextLib.RichTextBox rtfOther 
      Height          =   900
      Left            =   4515
      TabIndex        =   16
      Top             =   6795
      Width           =   200
      _ExtentX        =   344
      _ExtentY        =   1588
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDockInAdvice.frx":13E95
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
   Begin MSComctlLib.ImageList img16 
      Left            =   930
      Top             =   6705
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockInAdvice.frx":13F32
            Key             =   "屏蔽打印"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockInAdvice.frx":144CC
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockInAdvice.frx":14A66
            Key             =   "停嘱申请"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16dbl 
      Left            =   2220
      Top             =   6675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockInAdvice.frx":1B2C8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDockInAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Activate() '自已激活时
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) '要求主窗体刷新
Public Event StatusTextUpdate(ByVal Text As String) '要求更新主窗体状态栏文字
Public Event ViewEPRReport(ByVal 报告ID As Long, ByVal CanPrint As Boolean) '要求查看报告
Public Event PrintEPRReport(ByVal 报告ID As Long, ByVal Preview As Boolean) '要求打印报告
Public Event ViewPACSImage(ByVal 医嘱ID As Long) '要求进行观片
Public Event ExecLogNew(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, 完成 As Boolean) '执行情况登记
Public Event ExecLogModi(ByVal 医嘱ID As Long, ByVal 发送号 As Long, ByVal 科室ID As Long, ByVal 执行时间 As String, 完成 As Boolean) '执行情况修改
Public Event EditDiagnose(ParentForm As Object, ByVal 病人ID As Long, ByVal 主页ID As Long, ByVal 科室ID As Long, ByVal str类型 As String, Succeed As Boolean) '编辑住院诊断
Public Event SetEditState(ByVal blnEditState As Boolean)    '编辑状态时禁用菜单和可转移焦点的功能
Public Event DoByAdvice(ByVal lng医嘱ID As Long, ByVal lng相关ID As Long, ByVal lngWayID As Long, ByVal strTag As String)   '医嘱相关操作，lngWayID 功能ID。目前只支持  对医嘱计价,strTag 扩展参数

Private mint场合 As Integer  '调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
Private mMainPrivs As String '调用主界面所具有的权限,注意非内部模块权限
Private mcbsMain As Object
Private WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1
Private WithEvents mfrmEdit As frmInAdviceEdit '医嘱编辑窗体
Attribute mfrmEdit.VB_VarHelpID = -1
Private WithEvents mfrmEac As frmApplyConsultation    '会诊申请单窗体
Attribute mfrmEac.VB_VarHelpID = -1
Private WithEvents mfrmBilling As Form '记帐管理窗体
Attribute mfrmBilling.VB_VarHelpID = -1
Private WithEvents mfrmCompoundMedicine As frmCompoundMedicine  '输液配药记录
Attribute mfrmCompoundMedicine.VB_VarHelpID = -1
Private mobjPublicPACS As Object             'PACS业务封装公共部件

Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrBillPrint As String '当前打印的诊疗单据：报表编号、NO、记录性质

Private mintPState As TYPE_PATI_State '病人状态
Private mint执行状态 As Integer '医技站：执行状态
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng病人性质 As Long    '0-普通住院病人,1-门诊留观病人,2-住院留观病人
Private mbyt婴儿 As Byte
Private mint险类 As Integer

Private mlng病区ID As Long      '如果是转出病人，则为原病区ID
Private mlng科室ID As Long      '如果是转出病人，则为原科室ID
Private mlng前提ID As Long
Private mlng会诊医嘱ID As Long
Private mstr前提IDs As String
Private mlng界面科室ID As Long
Private mlng医护科室ID As Long
Private mstr姓名 As String
Private mstr性别 As String
Private mstr住院号 As String
Private mstr床号 As String
Private mdat重整 As Date '病案主页.医嘱重整时间
Private mblnBatch As Boolean '批量处理模式（固定弹出病人选择框）
Private mblnDirect As Boolean '是否直接调用功能（不显示医嘱清单的情况下）
Private mblnInsideTools As Boolean '内部工具条模式
Private mblnHaveAuditPriv As Boolean
Private mblnSignVisible As Boolean  '签名功能按钮可见性
Private mblnModalNew As Boolean '新开界面是否模态

Private mvInDate As Date '入院日期
Private mblnMoved As Boolean
Private mstr婴儿 As String
Private mstr住院医生 As String '病人的住院医师
Private mstr责任护士 As String '病人的责任护士
Private mint病案状态 As Integer
Private mlng路径状态 As Long    '-1-未导入，0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
Private mrsDefine As ADODB.Recordset    '医嘱内容定义
Private mobjVBA As Object
Private mobjScript As clsScript
Private mlng婴儿科室ID As Long
Private mlng婴儿病区ID As Long
Private mstr药品价格等级 As String '病人的药品价格等级
Private mstr卫材价格等级 As String '病人的卫材价格等级
Private mstr普通项目价格等级 As String '病人的普通项目价格等级

Private mblnFirst As Boolean '是否首次调用
Private mlngPlugInID As Long '自动执行的插件功能ID
Private mrsPlugInBar As ADODB.Recordset '菜单样式
Private mlngPromptRow As Long    '上一次，在鼠标移动图标列显示了提示信息的行

'Pass
Private mobjPassMap As Object  'PASS 窗体对象映射
Private mblnPass As Boolean  'PASS权限


'模块参数
Private mbln天数 As Boolean
Private mbln皮试验证 As Boolean
Private mbln护士签名 As Boolean
Private mblnShowExec As Boolean
Private mblnAutoRead As Boolean
Private mblnAutoReadEnabled As Boolean
Private mblnEditState As Boolean    '配药批次编辑状态
Private mblnNotEvaluete As Boolean  '未评估时允许添加医嘱到昨天
Private mlngBaby As Long
Private mblnFirstBaby As Boolean    '第一次按参数勾选
Private mlngBabyDept As Long      '上一次选择的婴儿选项
Private mintBillPrint As Integer   '0-选择医嘱清单打印诊疗单据（打印最后一次发送的诊疗单据），1-选择发送记录打印诊疗单据
Private mint申请单打印模式 As Integer  '1-发送时打印，2-新开时打印
Private mlngPrintType As Long '医嘱打印模式
Private mlngPrintPos As Long    '医嘱打印时，转科和出院医嘱打印在：0-长期医嘱单上，1-临时医嘱单上，2-两者都打印。
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln报告 As Boolean '是否是报告页签，保证在医嘱间切换时不刷新菜单
Private mstr检查入院诊断 As String
Private mstr自定义申请单IDs As String 'ID1,名称1|ID2,名称2···
Private mbln叮嘱发送执行 As Boolean
Private mobjFrmBlood As Object '血液执行窗体
Private mobjFrmBloodList As Object '血液明细窗体
Private mrs危急值 As ADODB.Recordset '当前病人的危急值信息
Private mbln危急值 As Boolean '处危急值的权限
Private mlng危急值ID As Long '当前处理的危急值记录ID
Private mbln确认会诊 As Boolean  '当前会诊记录确认医生是否到达
Private mbln医嘱定位最后 As Boolean  '医嘱光标默认定位到最后一行

'本地医嘱过滤条件
Private Enum CMD_FILTER
    ID_在用医嘱 = 1
    ID_所有医嘱 = 2
    ID_婴儿 = 3
    ID_重整 = 4
    ID_未记帐 = 5
    ID_科内 = 6
    ID_简洁 = 7
    ID_详细 = 8
    ID_全部 = 9
    ID_检查 = 10
    ID_检验 = 11
    ID_其他 = 12
    ID_时间 = 13
    ID_时间标签 = 14
    ID_是报告医嘱 = 15
    ID_非报告医嘱 = 16
    ID_未到终止时间 = 17
    ID_医嘱颜色示例 = 18
    ID_未出报告 = 19
    ID_已出报告 = 20
End Enum

Private Enum CMD_EXEC
    ID_显示执行 = 1
    ID_完成执行 = 2
    ID_取消完成 = 3
    ID_执行记录 = 4
    ID_执行调整 = 5
    ID_执行删除 = 6
    ID_核对 = 7
    ID_取消核对 = 8
End Enum

Private Type FilterCond
    婴儿 As Integer
    重整 As Boolean
    科内 As Boolean
    未记帐 As Boolean
    报告 As Integer     '0-全部，1－检查，2－检验，3－其他
    未出报告 As Boolean
    已出报告 As Boolean
    显示模式 As Integer '0-简洁，1－详细
    医嘱显示 As Integer '0-在用医嘱，1－所有医嘱
    过滤模式 As Integer '0-长嘱临嘱，1－长嘱，2－临嘱，3－报告
    开始时间 As Date
    结束时间 As Date
    是报告医嘱 As Boolean
    非报告医嘱 As Boolean
    未到终止时间 As Boolean '是否显示未到(执行终止时间)的医嘱
    医嘱ID As Long  '当前医嘱表格中选中行的医嘱ID
End Type
Private mvarCond As FilterCond
Private mblnHideFilter As Boolean
Private mintPreTime As Integer

'存放当前医嘱可回退列表
Private Type TYPE_AdviceRoll
    发送号 As Long
    操作类型 As Integer
    操作时间 As Date
    操作人员 As String
    操作内容 As String
End Type
Private marrRollList() As TYPE_AdviceRoll
Private mstr部门IDs As String '操作员人所属科室或病区
Private mblnAppend As Boolean '是否显示附加信息
Private mlngFontSize As Long  '字体大小

Private Enum COL医嘱清单
    '固定列
    COL_F标志 = 0
    COL_F报告 = 1
    '隐藏列
    COL_ID = 2
    COL_相关ID = COL_ID + 1
    COL_序号 = COL_ID + 2
    COL_婴儿ID = COL_ID + 3
    COL_医嘱状态 = COL_ID + 4   'flexcpData中存储审核状态
    COL_诊疗类别 = COL_ID + 5
    COL_操作类型 = COL_ID + 6
    COL_毒理分类 = COL_ID + 7
    COL_标志 = COL_ID + 8
    '可见列
    COL_警示 = COL_ID + 9 'Pass
    COL_期效 = COL_ID + 10
    COL_开始时间 = COL_ID + 11
    COL_并 = COL_ID + 12
    col_医嘱内容 = COL_ID + 13
    col_内容 = COL_ID + 14
    COL_皮试 = COL_ID + 15
    COL_总量 = COL_ID + 16
    COL_单量 = COL_ID + 17
    COL_天数 = COL_ID + 18
    COL_频率 = COL_ID + 19
    COL_用法 = COL_ID + 20
    COL_医生嘱托 = COL_ID + 21
    COL_执行时间 = COL_ID + 22
    COL_终止时间 = COL_ID + 23
    COL_执行科室 = COL_ID + 24
    COL_执行性质 = COL_ID + 25
    COL_上次执行 = COL_ID + 26
    COL_状态 = COL_ID + 27
    COL_开嘱医生 = COL_ID + 28
    COL_开嘱时间 = COL_ID + 29
    COL_校对护士 = COL_ID + 30
    COL_校对时间 = COL_ID + 31
    COL_停嘱医生 = COL_ID + 32
    COL_停嘱时间 = COL_ID + 33
    COL_停嘱护士 = COL_ID + 34
    COL_确认停嘱时间 = COL_ID + 35
    COL_基本药物 = COL_ID + 36
    COL_查阅状态 = COL_ID + 37
    COL_标本状态 = COL_ID + 38
    
    '隐藏列
    COL_诊疗项目ID = COL_ID + 39
    COL_试管编码 = COL_诊疗项目ID + 1
    COL_执行标记 = COL_诊疗项目ID + 2
    COL_屏蔽打印 = COL_诊疗项目ID + 3
    COL_前提ID = COL_诊疗项目ID + 4
    COL_签名否 = COL_诊疗项目ID + 5
    COL_文件ID = COL_诊疗项目ID + 6
    COL_报告项 = COL_诊疗项目ID + 7 '0-无报告，1-有报告并按编辑格式打印，2-有报告并按报表格式打印。
    COL_报告ID = COL_诊疗项目ID + 8
    COL_收费细目ID = COL_诊疗项目ID + 9
    COL_单量单位 = COL_诊疗项目ID + 10
    COL_开嘱科室ID = COL_诊疗项目ID + 11
    COL_审核状态 = COL_诊疗项目ID + 12
    COL_申请序号 = COL_诊疗项目ID + 13
    COL_审核标记 = COL_诊疗项目ID + 14
    COL_高危药品 = COL_诊疗项目ID + 15
    COL_标本部位 = COL_诊疗项目ID + 16   'PASS  药品名称
    COL_用药目的 = COL_诊疗项目ID + 17
    COL_检查报告ID = COL_诊疗项目ID + 18
    COL_处方审查状态 = COL_诊疗项目ID + 19
    COL_处方审查结果 = COL_诊疗项目ID + 20
    COL_RIS预约ID = COL_诊疗项目ID + 21
    COL_RIS报告ID = COL_诊疗项目ID + 22
    COL_LIS报告ID = COL_诊疗项目ID + 23
    COL_RIS预约状态 = COL_诊疗项目ID + 24
    col_诊疗项目名称 = COL_诊疗项目ID + 25
    COL_检查方法 = COL_诊疗项目ID + 26 '区分是备血医嘱还是用血医嘱
    COL_危急值ID = COL_诊疗项目ID + 27 '医嘱关和危急值关联
    COL_易跌倒 = COL_诊疗项目ID + 28 '药品至易跌倒
End Enum

Private COLPrice As New Collection
Private COLSend As New Collection
Private COLSign As New Collection
Private COLExec As New Collection

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByVal int场合 As Integer, _
                            ByVal blnInsideTools As Boolean, ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
    Dim lngTmp As Long
    
    mint场合 = int场合

    mblnInsideTools = blnInsideTools
    Set mfrmParent = frmParent
        mblnModalNew = blnModalNew

    If Not cbsMain Is Nothing Then

        '第一次调用时创建部件(不能放在Form_Load事件中，因为GetForm时触发该事件时还没有传mint场合)
        If Not mblnFirst Then
            mblnFirst = True

            Set mcbsMain = cbsMain
            Set cbsMain.Icons = zlCommFun.GetPubIcons
            Set gobjSquareCard = objSquareCard

            If mint场合 = 0 Then '医生站调用
                lngTmp = p住院医嘱下达
            ElseIf mint场合 = 1 Then '护士站调用
                lngTmp = p住院医嘱发送
            ElseIf mint场合 = 2 Then '医技站调用
                lngTmp = p住院医嘱下达
            End If
        
            '外挂程序对象初始化
            If gobjPlugIn Is Nothing Then
                On Error Resume Next
                Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
                err.Clear: On Error GoTo 0
            End If
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngTmp, mint场合)
                Call zlPlugInErrH(err, "Initialize")
                err.Clear: On Error GoTo 0
                Call GetPlugInBar(lngTmp, mint场合, mrsPlugInBar)
            End If

            'PASS接口初始化
            If gobjPass Is Nothing Then
                Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "合理用药监测", True)
                If Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassInit(gcnOracle, glngSys, PM_住院医嘱清单)
                    If gobjPass.PassType = 0 Then   '系统参数未启用合理用药监测
                        Set gobjPass = Nothing
                    Else
                        mblnPass = True
                    End If
                End If
            End If
           
        End If
        
        '不能放在Form_Load事件中，因为GetForm时触发该事件时还没有初始化合理用药检测)
        Call zlPASSMap
        If mblnPass Then
           'Pass
            Call gobjPass.zlPassAdviceColHidden(mobjPassMap) '警示列
        End If

        If mint场合 = 0 Then    '医生站调用
            Call DefCommandsInDoctor(cbsMain)
        ElseIf mint场合 = 1 Then    '护士站调用
            Call DefCommandsInNurse(cbsMain)
        ElseIf mint场合 = 2 Then    '医技站调用
            Call DefCommandsTechnic(cbsMain)
        End If

        If mint场合 <> 1 Then   '仅护士站才显示配药记录(Form_load时场合还没有传入)
            For lngTmp = 0 To tbcAppend.ItemCount - 1
                If tbcAppend(lngTmp).Tag = "配药" Then
                    Call tbcAppend.RemoveItem(lngTmp)
                    Exit For
                End If
            Next
        End If

        '外挂程序命令加载
        Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
        If mint场合 = 1 Then Call SetSendCommandBar '如果是护士站调用，重新添加发送按钮
    End If
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object, ByRef rsBar As ADODB.Recordset)
'功能：外挂部件菜单接入。
'说明：判断关键字  Auto  InTool 决定菜单样式
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim lngTmp As Long
    
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = 0
    If rsBar.RecordCount = 0 Then Exit Sub
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    '独立按钮
    rsBar.Filter = "IsInTool=1 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名)
                        objControl.IconId = rsBar!图标ID
                        objControl.Parameter = rsBar!功能名
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '下拉按钮，如果只有一个按钮，也当作独立按钮
    rsBar.Filter = "IsInTool=0 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        If Not objMenu Is Nothing Then
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能", , False)
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '工具栏按钮
    If mblnInsideTools Then
        Set objBar = cbsSub(2)
    Else
        Set objBar = cbsMain(2)
    End If
    Set objControl = objBar.FindControl(, conMenu_Help_Help)
    If Not objControl Is Nothing Then
        objControl.BeginGroup = True
        lngTmp = objControl.Index - 1
    Else
        lngTmp = -1
    End If
    rsBar.Filter = "IsInTool=1 and BarType=2"
    If Not rsBar.EOF Then
        With objBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名, lngTmp + 1)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End With
    End If
    
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "扩展功能", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名, lngTmp + 1)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    
    '自动执行的功能
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!功能ID
End Sub

Private Sub DefCommandsTechnic(ByVal cbsMain As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim lngIdx As Long
    Dim intTmp As Integer
    Dim strTmp As String
    Dim strName As String
    Dim lngID As Long
    Dim varArr As Variant
    Dim i As Long
    
    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "医嘱编辑(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "新开医嘱(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "修改医嘱(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "删除医嘱(&D)"
        End With
        
        intTmp = Val(Mid(gstrInUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",检查申请:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrInUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",检验申请:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrInUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",输血申请:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrInUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",手术申请:" & conMenu_Edit_OperationApply
        intTmp = Val(Mid(gstrInUseApp, 5, 1))
        If intTmp = 1 Then strTmp = strTmp & ",会诊申请:" & conMenu_Edit_ConsultationApply
                Get自定义申请单 2, mstr自定义申请单IDs
        If mstr自定义申请单IDs <> "" Then
            For i = 0 To UBound(Split(mstr自定义申请单IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(0)
            Next
        End If
        strTmp = Mid(strTmp, 2)
        
        If strTmp <> "" Then
            If InStr(strTmp, ",") = 0 Then
                strName = Split(strTmp, ":")(0)
                lngID = Val(Split(strTmp, ":")(1))
                Set objControl = .Add(xtpControlButton, lngID, strName)
                    objControl.IconId = conMenu_Manage_Request
                    objControl.ToolTipText = strName
                    objControl.BeginGroup = True
                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
            Else
                varArr = Split(strTmp, ",")
                For i = 0 To UBound(varArr)
                    strTmp = varArr(i)
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    
                    If i = 0 Then
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "下达申请"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "修改申请")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看申请")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "取消申请")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "检查预约")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "预约(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "调整预约(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "取消预约(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "医嘱作废(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "医嘱停止(&S)")
        If InStr(GetInsidePrivs(p住院医嘱下达), "发送门诊费用") = 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "临嘱发送(&G)"): objControl.BeginGroup = True
            objControl.IconId = conMenu_Edit_Send
        Else
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "临嘱发送"): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Edit_SendBilling, "住院记帐"    'update事件中根据当前是否留观病人再决定显示门诊记帐还是住院记帐
                .Add xtpControlButton, conMenu_Edit_SendCharge, "门诊收费"
            End With
        End If
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Untread, "医嘱回退(&L)")
        '医技工作站提供查看美康的药品说明书的菜单
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "关键图像")
       '2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "浏览检验结果(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "浏览检查结果(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "浏览检查图像和报告(&Y)")
                objControl.IconId = 237
        End If

        If gbln血库系统 Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_BloodInstant, "输血执行单")
            objControl.BeginGroup = True
        End If
    End With
    If Not objMenu Is Nothing Then
        With objMenu.CommandBar.Controls
            If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
            '子项放在最前面,反序加入
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印诊疗单据")
            
            '外挂菜单
            Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
        End With
    End If
    
    '报表菜单:主窗体可能没有,放在查看菜单前面
    '-----------------------------------------------------
    '工作站报表菜单自动显示报表是针对工作站的模块号统一发布
    '而这几张报表是医嘱虚拟模块中的，需要在该模块中单独处理
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
    End If
    
    If Not objMenu Is Nothing Then
        With objMenu.CommandBar.Controls
            If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
            '子项放在最前面,反序加入
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印诊疗单据", 1)
            objPopup.Visible = False '隐藏，只用于右键菜单处理
        End With
    End If

    '查看菜单
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '状态栏项后
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "附加信息(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "自动隐藏过滤条件栏(&H)", objControl.Index + 1)
    End With
        
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "电子签名(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "医嘱签名(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln启用影像信息系统预约 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "打印预约单")
                objControl.IconId = 103
        End If
        If gbln科室药房对照按本机参数设置 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "医嘱选项(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "成套方案定义(&S)"): objControl.BeginGroup = True
    End With

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Call AddToolBarInDoctor
    
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新开医嘱
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改医嘱
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete '删除医嘱
        .Add FCONTROL, vbKeyS, conMenu_Edit_Stop '停止医嘱
        .Add FCONTROL, vbKeyG, conMenu_Edit_SendBilling '医嘱发送
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread '医嘱回退
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend '查阅报告
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '观片处理
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '浏览检查图像和报告
       
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '自动隐藏过滤条件栏
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '浏览检验结果
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '浏览检查结果
        .Add 0, vbKeyF11, conMenu_Tool_Option '医嘱选项
    End With

    '设置不常用命令
    '-----------------------------------------------------
    With cbsMain.Options
    End With
End Sub

Private Sub DefCommandsInDoctor(ByVal cbsMain As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl, lngIdx As Long
    
    Dim varArr As Variant
    Dim strTmp As String
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    Dim i As Long
    
    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "医嘱编辑(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "新开医嘱(&A)"
            .Add xtpControlButton, conMenu_Edit_Audit, "新嘱审核(&T)"
            .Add xtpControlButton, conMenu_Edit_Modify, "修改医嘱(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "删除医嘱(&D)"
        End With
        
        intTmp = Val(Mid(gstrInUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",检查申请:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrInUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",检验申请:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrInUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",输血申请:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrInUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",手术申请:" & conMenu_Edit_OperationApply
        intTmp = Val(Mid(gstrInUseApp, 5, 1))
        If intTmp = 1 Then strTmp = strTmp & ",会诊申请:" & conMenu_Edit_ConsultationApply
        Get自定义申请单 2, mstr自定义申请单IDs
        If mstr自定义申请单IDs <> "" Then
            For i = 0 To UBound(Split(mstr自定义申请单IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(0)
            Next
        End If
        strTmp = Mid(strTmp, 2)
        
        If strTmp <> "" Then
            If InStr(strTmp, ",") = 0 Then
                strName = Split(strTmp, ":")(0)
                lngID = Val(Split(strTmp, ":")(1))
                Set objControl = .Add(xtpControlButton, lngID, strName)
                    objControl.IconId = conMenu_Manage_Request
                    objControl.ToolTipText = strName
                    objControl.BeginGroup = True
                                        If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
            Else
                varArr = Split(strTmp, ",")
                For i = 0 To UBound(varArr)
                    strTmp = varArr(i)
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    
                    If i = 0 Then
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "下达申请"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "修改申请")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看申请")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "取消申请")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "检查预约")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "预约(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "调整预约(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "取消预约(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "标记未用(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Sort, "调整顺序(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "医嘱作废(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "医嘱停止(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_StopAudit, "停嘱审核(&W)"): objControl.IconId = conMenu_Edit_Audit
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "医嘱暂停(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "医嘱启用(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "医嘱重整(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "屏蔽打印")
        If gbln血库系统 Then Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReaction, "输血反应"): objControl.BeginGroup = True: objControl.IconId = 4113
        If mbln危急值 Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_CriticalAdvice, "危急值医嘱")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "医嘱批量执行(&W)"): objControl.BeginGroup = True: objControl.IconId = 3587
        If InStr(GetInsidePrivs(p住院医嘱下达), "发送门诊费用") = 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "临嘱发送(&G)"): objControl.BeginGroup = True
            objControl.IconId = conMenu_Edit_Send
        Else
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "临嘱发送"): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Edit_SendBilling, "住院记帐"    'update事件中根据当前是否留观病人再决定显示门诊记帐还是住院记帐
                .Add xtpControlButton, conMenu_Edit_SendCharge, "门诊收费"
            End With
        End If
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Untread, "医嘱回退(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "销帐申请(&C)")
        objControl.IconId = conMenu_Edit_ChargeOff
                
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片处理(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "关键图像")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "浏览检验结果(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "浏览检查结果(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "浏览检查图像和报告(&Y)")
                objControl.IconId = 237
        End If

        Set objControl = .Add(xtpControlButton, conMenu_Manage_RecipeAuditView, "查看处方审查结果")
        objControl.IconId = 3205
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewDrugExplain, "查看药品说明书")
        objControl.IconId = 3205
        If gbln审方系统 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Refcom, "拒绝审查理由")
                objControl.IconId = 3205
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewRefcom, "查阅审核未通过信息")
                objControl.IconId = 3205
        End If
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
        '2012-02-16 by　陈东
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPrint, "批量打印检验报告(&J)"): objPopup.BeginGroup = True

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印诊疗单据(&1)")
        objPopup.BeginGroup = True
        '外挂菜单
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With
    
    '报表菜单:主窗体可能没有,放在查看菜单前面
    '-----------------------------------------------------
    '工作站报表菜单自动显示报表是针对工作站的模块号统一发布
    '而这几张报表是医嘱虚拟模块中的，需要在该模块中单独处理
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
    End If
    With objMenu.CommandBar.Controls
        If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
        '子项放在最前面,反序加入
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill3, "医嘱记录本(&3)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill1, "病人医嘱单(&2)", 1)
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印诊疗单据(&1)", 1)
    End With
    
    '查看菜单
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '状态栏项后
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "附加信息(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "自动隐藏过滤条件栏(&H)", objControl.Index + 1)
    End With
        
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "电子签名(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "医嘱签名(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln启用影像信息系统预约 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "打印预约单")
                objControl.IconId = 103
        End If
        If gbln科室药房对照按本机参数设置 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "医嘱选项(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "成套方案定义(&S)"): objControl.BeginGroup = True
    End With

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Call AddToolBarInDoctor

    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新开医嘱
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改医嘱
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete '删除医嘱
        .Add FCONTROL, vbKeyS, conMenu_Edit_Stop '停止医嘱
        .Add FCONTROL, vbKeyG, conMenu_Edit_SendBilling '医嘱发送
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread '医嘱回退
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend '查阅报告
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '观片处理
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '浏览检查图像和报告
        
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '自动隐藏过滤条件栏
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '浏览检验结果
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '浏览检查结果
        .Add 0, vbKeyF11, conMenu_Tool_Option '医嘱选项
    End With

End Sub

Private Sub DefCommandsInNurse(ByVal cbsMain As Object)
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim lngStart As Long
    Dim blnFirst As String
    Dim lngIdx As Long
  
    '医嘱菜单:放在管理菜单(主窗体可能没有)、文件菜单后面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "医嘱(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "医嘱编辑(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "新开医嘱(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "修改医嘱(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "删除医嘱(&D)"
        End With
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "检查预约")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "预约(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "调整预约(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "取消预约(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "标记未用(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Sort, "调整顺序(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "医嘱作废(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "医嘱停止(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "确认停止(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "医嘱暂停(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "医嘱启用(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "医嘱校对(&V)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Price, "计价调整(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "医嘱重整(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "屏蔽打印")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "医嘱发送(&G)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "医嘱批量执行(&W)"): objControl.BeginGroup = True: objControl.IconId = 3587
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "医嘱批量核对(&X)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "药品留存登记(&J)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Untread, "医嘱回退(&L)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePrice, "对医嘱记帐")
            objControl.IconId = conMenu_Edit_Price
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_ChargeOff, "费用销帐(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "超期发送收回(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Test, "皮试结果(&T)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatUnPack, "批量打包(&U)"): objControl.BeginGroup = True: objControl.IconId = 312
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MeetArrive, "确认会诊医生到场(&M)"): objControl.BeginGroup = True: objControl.IconId = 8122
        
        '护士站提供调阅阅美康的药品说明书的菜单
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "报告(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        ' 2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "浏览检验结果(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "浏览检查结果(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView
        
        '2012-02-16 by　陈东
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPrint, "批量打印检验报告(&J)"): objPopup.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印诊疗单据(&3)")
        objPopup.BeginGroup = True
        '2017-11-10 刘鹏飞
        If gbln血库系统 Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_BloodInstant, "输血执行单")
            objControl.BeginGroup = True
        End If
        '外挂菜单
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With

    '报表菜单:主窗体可能没有,放在查看菜单前面
    '-----------------------------------------------------
    '工作站报表菜单自动显示报表是针对工作站的模块号统一发布
    '而这几张报表是医嘱虚拟模块中的，需要在该模块中单独处理
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "报表(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '对xtpControlPopup类型的命令ID需重新赋值
    End If
    With objMenu.CommandBar.Controls
        If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
        '子项放在最前面,反序加入
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill3, "医嘱记录本(&5)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill1, "病人医嘱单(&4)", 1): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "打印诊疗单据(&3)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "打印执行单(&2)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_DrugQuery, "药疗收发查询(&1)", 1)
    End With

    '查看菜单
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '状态栏后
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "附加信息(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "自动隐藏过滤条件栏(&H)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_AdviceLost, "医嘱刷新时定位到最后(&L)", objControl.Index + 1)
        
        Set objControl = .Find(, conMenu_View_Refresh) '刷新项前(多个时注意反序)
        Set objControl = .Add(xtpControlButton, conMenu_View_Notify, "刷新提醒(&B)", objControl.Index)
        objControl.BeginGroup = True
    End With
    
    '工具菜单:主窗体可能没有,放在帮助菜单前面
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "工具(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "电子签名(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "医嘱签名(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "验证签名(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "取消签名(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln启用影像信息系统预约 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "打印预约单")
                objControl.IconId = 103
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrintBat, "批量打印预约单")
        End If
        If gbln科室药房对照按本机参数设置 Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "医嘱选项(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "成套方案定义(&S)"): objControl.BeginGroup = True
    End With

    '工具栏定义:从文件及管理菜单的命令按钮之后开始加入
    '-----------------------------------------------------
    Call AddToolBarInDoctor
     
    '命令的快键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '新开医嘱
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '修改医嘱
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete '删除医嘱
        .Add FCONTROL, vbKeyS, conMenu_Edit_Stop '停止医嘱
        .Add FCONTROL, vbKeyV, conMenu_Edit_Audit '医嘱校对
        .Add FCONTROL, vbKeyI, conMenu_Edit_Price '医嘱计价
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send '发送医嘱
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread '医嘱回退(xtpControlSplitButtonPopup方式时快捷键显示不到菜单上,但工具提示中有)
        .Add FCONTROL, vbKeyT, conMenu_Edit_Test '皮试结果
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend  '查阅报告
         
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '自动隐藏过滤条件栏
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '浏览检验结果
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '浏览检查结果
        .Add 0, vbKeyF2, conMenu_Edit_SendInfusion '发送输液药品医嘱 此菜单有无处于变化中，可能在SetSendCommandBar过程中被添加或者不添加
        .Add 0, vbKeyF9, conMenu_Report_AdviceBill1 '医嘱单打印
        .Add 0, vbKeyF10, conMenu_View_Notify '刷新医嘱提醒
        .Add 0, vbKeyF11, conMenu_Tool_Option '医嘱选项
    End With

End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim vRoll As TYPE_AdviceRoll, i As Long
    Dim arrTmp As Variant, strTmp As String
    Dim lng医嘱ID As Long
    Dim rsTmp As ADODB.Recordset
    
    If CommandBar.Parent Is Nothing Then Exit Sub

    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_CriticalAdvice
        If mbln危急值 And Not mrs危急值 Is Nothing Then
            mrs危急值.Filter = 0
            If Not mrs危急值.EOF Then
                Set rsTmp = GetCriticalAdvice(lng医嘱ID)
                With CommandBar.Controls
                    .DeleteAll
                    mrs危急值.MoveFirst
                    For i = 1 To mrs危急值.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_CriticalAdvice * 100# + i, mrs危急值!危急值描述 & "")
                            objControl.Parameter = mrs危急值!ID & "," & lng医嘱ID
                        rsTmp.Filter = "危急值ID=" & mrs危急值!ID
                        If Not rsTmp.EOF Then
                            objControl.Checked = True
                        End If
                        mrs危急值.MoveNext
                    Next
                    mrs危急值.MoveFirst
                End With
            End If
            mrs危急值.Filter = 0
        End If
    Case conMenu_Edit_Untread    '医嘱回退
        With CommandBar.Controls
            .DeleteAll
            For i = 1 To UBound(marrRollList)
                vRoll = marrRollList(i)
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread * 100# + i, vRoll.操作内容)
                If i = 1 Then
                    If Not RollFirstEnabled Then objControl.Enabled = False
                Else
                    If i = 2 Then
                        objControl.BeginGroup = True
                    End If
                    objControl.Enabled = False
                End If
                If i = 50 Then Exit For    '只加入50条显示
            Next
        End With
    Case conMenu_ReportPopup
        Set objControl = CommandBar.FindControl(, conMenu_Report_ClinicBill)
        If Not objControl Is Nothing Then
            objControl.Visible = False
        End If
    Case conMenu_Edit_ChargeOff    '费用销帐
        With CommandBar.Controls
            If .Count = 0 Then
                .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "销帐申请(&L)").BeginGroup = True
                .Add xtpControlButton, conMenu_Edit_ChargeDelAudit, "销帐审核(&U)"
                .Add(xtpControlButton, conMenu_Edit_ChargeOff * 10# + 1, "冲销当前选择的单据(&1)").BeginGroup = True
                .Add xtpControlButton, conMenu_Edit_ChargeOff * 10# + 2, "冲销当前医嘱该次发送的单据(&2)"
                .Add xtpControlButton, conMenu_Edit_ChargeOff * 10# + 3, "冲销该次发送的所有单据(&3)"
            End If
        End With
    Case conMenu_Edit_Compend    '报告
        With CommandBar.Controls
            If .Count = 0 Then
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "查阅报告(病历格式)"
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 6, "查阅报告(报表格式)"
                If gobjExchange Is Nothing Then
                    If mint场合 = 1 Then    '护士站
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)"
                    Else
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "打印报告(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)"

                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "我已查阅(&R)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "自动标记(&A)"
                    End If
                End If
            End If
        End With
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99
        'PASS药嘱审查
        If mblnPass Then
            Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, CommandBar, conMenu_Edit_MediAudit)
        End If
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim lngParValue As Long, objControl As CommandBarControl
    Dim lng待入住病人医嘱 As Long
    Dim strErr As String
    
    mblnBatch = False
    mblnDirect = False
 
    Select Case Control.ID
    Case conMenu_File_PrintSet '打印设置
        Call zlPrintSet
    Case conMenu_File_Preview '预览医嘱清单
        Call OutputList(2)
    Case conMenu_File_Print '打印医嘱清单
        Call OutputList(1)
    Case conMenu_File_Excel '输出医嘱清单
        Call OutputList(3)
    Case conMenu_Help_Help '帮助
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_View_AdviceLost '医嘱是否定位最后
        mbln医嘱定位最后 = Not mbln医嘱定位最后
        Call zlDatabase.SetPara("医嘱光标默认定位到最后一行", IIF(mbln医嘱定位最后, 1, 0), glngSys, p住院医嘱下达)
    Case conMenu_View_Append '附加信息
        mblnAppend = Not mblnAppend
        tbcAppend.Visible = Not tbcAppend.Visible
        fraAdviceUD.Visible = Not fraAdviceUD.Visible
        Call Form_Resize
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
        If vsAppend.Visible And vsAppend.Enabled Then
            vsAppend.SetFocus
        Else
            If vsAdvice.Visible And vsAdvice.Enabled Then vsAdvice.SetFocus
        End If
        Call cbsSub_Resize
    Case conMenu_View_Hide '自动隐藏过滤工具栏
        mblnHideFilter = Not mblnHideFilter
        cbsSub(2).Visible = Not mblnHideFilter And cbsSub(2).Controls.Count > 0
        cbsSub(3).Visible = Not mblnHideFilter
        fraHide.Visible = mblnHideFilter
        cboTime.Visible = Not mblnHideFilter
        cbsSub.RecalcLayout
    Case conMenu_Edit_NewItem, conMenu_Edit_NewItem * 10# + 1 '新开医嘱
        If Control.Parameter <> "" Then
            mlng危急值ID = Val(Control.Parameter)
            Call GetCriticalData
        Else
            mlng危急值ID = 0
        End If
        Call FuncAdviceAdd
    Case conMenu_Edit_Modify '修改医嘱
        Call FuncAdviceModi
    Case conMenu_Edit_Delete, conMenu_Edit_ApplyDel '删除医嘱'取消检验申请
        Call FuncAdviceDel
    Case conMenu_Edit_LISApply, conMenu_Edit_LISApply * 10 + 1   '检验申请
        Call FuncApplyLIS(0)
    Case conMenu_Edit_ApplyModi '修改申请
        Call FuncApplyModi
    Case conMenu_Edit_NewRisSch 'RIS预约
        Call FuncAdviceRISSch
    Case conMenu_Edit_NewRisDel '取消预约
        Call FuncAdviceRISDel
    Case conMenu_Edit_NewRisModi
        Call FuncAdviceRISModi
    Case conMenu_Tool_RisPrint, conMenu_Tool_RisPrintBat
        Call FuncAdviceRISPrintSch(Control.ID)
    Case conMenu_Edit_PacsApply, conMenu_Edit_PacsApply * 10 + 1 '检查申请
        Call FuncApplyPACS(0, 0)
    Case conMenu_Edit_BloodApply, conMenu_Edit_BloodApply * 10 + 1  '输血申请
        Call FuncApplyBlood(0)
    Case conMenu_Edit_OperationApply, conMenu_Edit_OperationApply * 10 + 1 '手术申请
        Call FuncApplyOperation(0)
    Case conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101#
        FuncApplyCustom 0, Control.Parameter
    Case conMenu_Edit_ConsultationApply, conMenu_Edit_ConsultationApply * 10 + 1 '会诊申请
        Call FuncApplyConsultation(0)
    Case conMenu_Edit_TraReaction  '输血反应
        Call FuncTraReaction(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), p住院医嘱下达, mblnMoved)
    Case conMenu_Edit_CriticalAdvice * 100# + 1 To conMenu_Edit_CriticalAdvice * 100# + 99
        Call FuncCriticalAdvice(Control.Parameter, Control.Checked)
    Case conMenu_Edit_ApplyView '查看申请
        Call FuncApplyView
    Case conMenu_Edit_UnUse '标记未用医嘱
        Call FuncAdviceUnUse
    Case conMenu_Edit_Sort   '调整顺序
        Call FuncAdviceSort
    Case conMenu_Edit_Audit '审核医嘱,校对医嘱
        If mint场合 = 1 Then
            Call FuncAdviceVerify
        Else
            Call FuncAdviceAudit
        End If
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99  '合理用药审查
        If mblnPass Then
            Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#)
        End If
    Case conMenu_Edit_Stop '停止医嘱
        Call FuncAdviceStop
    Case conMenu_Edit_StopAudit '停嘱审核
        Call FuncAdviceStopAudit
    Case conMenu_Edit_Blankoff '作废医嘱
        Call FuncAdviceRevoke
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6 '查阅、打印报告
        Call FuncEPRReport(Control.ID)
    Case conMenu_Edit_Compend * 10# + 4 '我是否已经查阅该报告
        Call FuncExecReportRead(Not Control.Checked)
    Case conMenu_Edit_Compend * 10# + 5 '自动标记查阅状态
        mblnAutoRead = Not mblnAutoRead
        Call zlDatabase.SetPara("自动标记报告查阅状态", IIF(mblnAutoRead, 1, 0), glngSys, p住院医嘱下达)
    Case conMenu_Edit_MarkMap '观片处理
        RaiseEvent ViewPACSImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    Case conMenu_Edit_MarkKeyMap '关键图像
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowStaticImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_ViewPacs '浏览检查图像和报告
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowPatientImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_Price '计价调整
        Call FuncAdvicePrice
    Case conMenu_Edit_ReStop '确认停止
        Call FuncAdviceConfirm(Control.Parameter = "医嘱提醒", Control)
    Case conMenu_Edit_Pause '医嘱暂停
        Call FuncAdvicePause
    Case conMenu_Edit_Reuse '医嘱启用
        Call FuncAdviceResume
    Case conMenu_Edit_ClearUp '医嘱重整
        Call FuncAdviceReform
    Case conMenu_Edit_NoPrint '屏蔽打印
        Call FuncAdviceNoPrint
    Case conMenu_Edit_ViewDrugExplain '查看药品说明书
        Call FuncViewDrugExplain(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_收费细目ID)), mfrmParent)
    Case conMenu_Edit_Refcom '拒绝审查理由
        Call FuncDrugRefcom '药品审查拒绝理由
    Case conMenu_Edit_ViewRefcom '查阅审核未通过信息
        If Not gobjPass Is Nothing And mlng病人ID <> 0 And mlng主页ID <> 0 Then Call gobjPass.ZLPharmReviewResultShow(Me, mlng病人ID, mlng主页ID)
    Case conMenu_Edit_Test '皮试结果
        Call FuncAdviceTest
    Case conMenu_Edit_Send, conMenu_Edit_SendInfusion '医嘱发送
        Call FuncAdviceSend(Control.Parameter = "医嘱提醒", Control)
    Case conMenu_Edit_SendCharge, conMenu_Edit_SendBilling '医生或医技，临嘱发送
        Call FuncAdviceSend(Control.Parameter = "医嘱提醒", Control)
    Case conMenu_Edit_BatExecute '医嘱批量执行
        '检查病人是否正在审核
        If Not CheckPatiIsAduit Then Exit Sub
        frmAdviceBatExecute.ShowMe 1, Me, mlng病区ID, mlng病人ID, mint场合, 0, mlng医护科室ID, mlng婴儿科室ID, mlng婴儿病区ID
    Case conMenu_Manage_ThingAudit '医嘱批量核对
        '检查病人是否正在审核
        If Not CheckPatiIsAduit Then Exit Sub
        frmAdviceBatExecute.ShowMe 1, Me, mlng病区ID, mlng病人ID, mint场合, 1, mlng医护科室ID, mlng婴儿科室ID, mlng婴儿病区ID
    Case conMenu_Edit_Surplus '药品留存登记
        Call frmDrugSurplus.ShowMe(mfrmParent, mlng病区ID)
    Case conMenu_Edit_SendBack '超期发送收回
        Call FuncAdviceSendBack
    Case conMenu_Edit_Untread, conMenu_Edit_Untread * 100# + 1 '医嘱回退(只能顺序回退)
        If Control.ID = conMenu_Edit_Untread Then
            '允许查看回退列表但尚未弹出
            If Not RollFirstEnabled Then Exit Sub
        End If
        Call FuncAdviceRoll
    Case conMenu_Edit_ChargeOff * 10# + 1 To conMenu_Edit_ChargeOff * 10# + 3 '直接冲销
        Call FuncAdviceChargeOff(Control.ID - conMenu_Edit_ChargeOff * 10# - 1)
    Case conMenu_Tool_SignNew '医嘱签名
        Call FuncAdviceSign
    Case conMenu_Tool_SignVerify '验证签名
        Call FuncAdviceSignVerify
    Case conMenu_Tool_SignEarse '取消签名
        Call FuncAdviceSignErase
    Case conMenu_Report_ClinicBill * 100# + 1 To conMenu_Report_ClinicBill * 100# + 99 '打印诊疗单据
        Call FuncBillPrint(Control)
    Case conMenu_Edit_AdvicePrice
        RaiseEvent DoByAdvice(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)), conMenu_Edit_AdvicePrice, "")
    Case conMenu_Edit_ChargeDelApply, conMenu_Edit_ChargeDelAudit '销帐申请审核
        Call FuncAdviceReCharge(Control.ID)
    Case conMenu_Report_DrugQuery '药疗收发查询
        Call FuncDrugSendQuery
    Case conMenu_Report_Reports '病区常用报表
        Call FuncAdviceReport
    Case conMenu_Report_AdviceBill1 '病人医嘱单
        Call frmAdvicePrint.ShowMe(mfrmParent, mlng病人ID, mlng主页ID)
    Case conMenu_Report_AdviceBill3 '医嘱记录本
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_3", mfrmParent, "病人科室=" & mlng科室ID)
    Case conMenu_Tool_Reference_2 '诊疗措施参考
        Call zlItemRef
    Case conMenu_Edit_BatUnPack '批量打包
        frmCompoundPack.ShowMe 1, Me, mlng病区ID, mlng病人ID, mlng医护科室ID, mlng婴儿科室ID, mlng婴儿病区ID
    Case conMenu_Tool_Option '医嘱选项
        frmInAdviceSetup.Show 1, mfrmParent
    Case conMenu_Tool_Define '成套方案定义
        Call FuncToolScheme
    Case conMenu_Manage_ReportLisView  '检验报告浏览
        Call FuncViewLisRpt
Case conMenu_Manage_ReportPacsView  '检查报告浏览
        Call FuncViewPacsRpt
    Case conMenu_Edit_MeetArrive
        Call Execute确认会诊(IIF(Control.Caption = "取消会诊医生到场(&K)", True, False))
    Case conMenu_Manage_RecipeAuditView '查看处方审查结果
        If InitObjRecipeAudit(p住院医嘱下达) Then
            Call gobjRecipeAudit.ShowResult(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)), mfrmParent)
        End If
    Case conMenu_Manage_ReportPrint
        Call PrintLisReport(mlng病区ID, mfrmParent)
    Case conMenu_Report_BloodInstant
        Call PrintBloodReport(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
        If CreatePlugInOK(p住院医嘱下达, mint场合) Then
            On Error Resume Next
            If PlugExeNew(Control.Parameter) = False Then
                Call gobjPlugIn.ExecuteFunc(glngSys, Decode(mint场合, 0, p住院医嘱下达, 1, p住院医嘱发送, 2, p住院医嘱下达), _
                    Control.Parameter, mlng病人ID, mlng主页ID, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mlng前提ID, mint场合)
                Call zlPlugInErrH(err, "ExecuteFunc")
                err.Clear: On Error GoTo 0
            End If
        End If
    End Select
End Sub


Private Function PlugExeNew(ByVal strName As String) As Boolean
'功能：向下兼容外挂部件的ExecuteFunc过程
    Dim lngID As Long
    Dim strXML As String
On Error GoTo errH
    If CreatePlugInOK(p住院医嘱下达, mint场合) Then
        With vsAdvice
            lngID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
            strXML = "<ROOT><诊疗项目名称>" & .TextMatrix(.Row, col_诊疗项目名称) & "</诊疗项目名称></ROOT>"
            Call gobjPlugIn.ExecuteFunc(glngSys, Decode(mint场合, 0, p住院医嘱下达, 1, p住院医嘱发送, 2, p住院医嘱下达), strName, mlng病人ID, mlng主页ID, lngID, mlng前提ID, mint场合, strXML)
        End With
    End If
   PlugExeNew = True
   Exit Function
errH:
    If err.Number = 450 Then
        PlugExeNew = False
        err.Clear
    Else
        PlugExeNew = True
        Call zlPlugInErrH(err, "ExecuteFunc")
        err.Clear: On Error GoTo 0
    End If
End Function


Public Sub zlExecuteCommandBarsDirect(ByVal Control As CommandBarControl, ByRef frmParent As Object, ByRef strPrivs As String, _
    ByVal bln批量 As Boolean, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal byt婴儿 As Byte, _
    ByVal lng病区ID As Long, ByVal lng科室id As Long, ByVal lng前提ID As Long, ByVal lng界面科室ID As Long, ByVal int场合 As Integer, _
    ParamArray arrPar() As Variant)
'功能：提供单独调用医嘱操作的接口
    Dim strErr As String
    
    Set mfrmParent = frmParent
    mMainPrivs = strPrivs
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mbyt婴儿 = byt婴儿
    mlng病区ID = lng病区ID
    mlng科室ID = lng科室id
    mlng前提ID = lng前提ID
    mlng界面科室ID = lng界面科室ID
    mlng医护科室ID = lng界面科室ID
    mblnSignVisible = True
    If mint场合 = 0 Then
        If CheckSign(1, 0, mlng界面科室ID, mlng科室ID, 2, False, gobjESign) = False Then
            mblnSignVisible = False '不同场合没有设置要使用签名
        End If
    ElseIf mint场合 = 2 Then
        If CheckSign(3, 0, mlng界面科室ID, mlng科室ID, 2, False, gobjESign) = False Then
            mblnSignVisible = False '不同场合没有设置要使用签名
        End If
    ElseIf mint场合 = 1 Then
        If CheckSign(2, mlng医护科室ID, , , , False, gobjESign) = False Then
            mblnSignVisible = False '不同场合没有设置要使用签名
        End If
    End If
    
    mint场合 = int场合
    mblnBatch = bln批量
    mblnDirect = True
    mblnInsideTools = False
    
    mblnMoved = CheckPatiDataMoved(lng病人ID, lng主页ID)
    '创建LIS部件
    If Control.ID = conMenu_Manage_ReportLisView Or Control.ID = conMenu_Edit_Send Then
       Call InitObjLis(p住院护士站)
    End If
    
    Select Case Control.ID
    Case conMenu_Edit_NewItem '新开医嘱
        Call FuncAdviceAdd
    Case conMenu_Edit_Audit '审核医嘱,校对医嘱
        If mint场合 = 1 Then
            Call FuncAdviceVerify
        Else
            Call FuncAdviceAudit
        End If
    Case conMenu_Edit_Price '计价调整
        Call FuncAdvicePrice
    Case conMenu_Edit_Send, conMenu_Edit_SendInfusion '医嘱发送
        Call FuncAdviceSend(Not bln批量, Control)
    Case conMenu_Edit_Stop '停止医嘱
        Call FuncAdviceStop
    Case conMenu_Edit_StopAudit '停嘱审核
        Call FuncAdviceStopAudit
    Case conMenu_Edit_ReStop '确认停止
        Call FuncAdviceConfirm(Not bln批量, Control)
    
    Case conMenu_Edit_BatExecute '医嘱批量执行
        frmAdviceBatExecute.ShowMe 1, frmParent, lng病区ID, lng病人ID, mint场合, 0, mlng医护科室ID, mlng婴儿科室ID, mlng婴儿病区ID
    Case conMenu_Manage_ThingAudit '医嘱批量核对
        frmAdviceBatExecute.ShowMe 1, frmParent, lng病区ID, lng病人ID, mint场合, 1, mlng医护科室ID, mlng婴儿科室ID, mlng婴儿病区ID
        
    Case conMenu_Edit_Blankoff '作废医嘱
        Call FuncAdviceRevoke
        
        
    Case conMenu_Edit_Pause '医嘱暂停
        Call FuncAdvicePause
    Case conMenu_Edit_Reuse '医嘱启用
        Call FuncAdviceResume
        
        
    Case conMenu_Edit_SendBack '超期发送收回
        Call FuncAdviceSendBack
    Case conMenu_Report_DrugQuery '药疗收发查询
        Call FuncDrugSendQuery
    
    Case conMenu_Manage_ReportLisView  '检验报告浏览
        If mlng病人ID <> 0 Then
            If Not gobjLIS Is Nothing And Sys.SystemShareWith(2500) Then
                gobjLIS.PatientSampleBrowse mfrmParent, mlng病人ID, mMainPrivs, mlng科室ID, mlng病区ID, 2, mlng主页ID
            Else
                frmLisView.ShowMe mlng病人ID, p住院医嘱下达, mfrmParent
            End If
        End If
    Case conMenu_Edit_Surplus '药品留存登记
        Call frmDrugSurplus.ShowMe(mfrmParent, mlng病区ID)
        
    Case conMenu_Report_Reports '病区常用报表
        Call FuncAdviceReport
        
    End Select
End Sub

Private Sub FuncAdviceReport()
'功能：调用病区常用报表
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '判断隐藏参数，如果有此参数行则使用老版执行单打印功能
    strSQL = "select 1 from zlParameters a where a.系统=[1] and a.模块=[2] and a.参数名=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys, p住院医嘱发送, "还原老版执行单打印功能")
    
    
    On Error Resume Next
    If rsTmp.EOF Then
        Call frmAdviceWardReport.ShowMe(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng病区ID, mlng病人ID)
    Else
        Call frmAdviceReport.ShowMe(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng病区ID, mlng病人ID, mblnDirect And Not mblnBatch Or mblnInsideTools, mlng医护科室ID, mlng婴儿病区ID)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSendBack()
'功能：超期发送收回
    Dim blnRoll As Boolean
    
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    '68081出院不允许操作医嘱产生费用变化
    If mintPState = ps预出 Or mintPState = ps出院 Then
        Call MsgBox("该病人已" & IIF(mintPState = ps预出, "预", "") & "出院，不允许进行医嘱超期收回！", vbInformation, gstrSysName)
        Exit Sub
    End If
    On Error Resume Next
    blnRoll = frmAdviceRollSend.ShowMe(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng病区ID, mlng病人ID, mlng主页ID, mblnDirect And Not mblnBatch Or mblnInsideTools, False, mlng医护科室ID, mlng婴儿病区ID)
    
    If blnRoll And mblnDirect = False Then
        RaiseEvent StatusTextUpdate("")
        Call LoadAdvice
    End If
End Sub

Public Sub zlCheckPrivs(ByVal Control As CommandBarControl, ByVal int场合 As Integer)
'功能：检查菜单或按钮的权限，并设置其可见性
    mint场合 = int场合
    Call SetControlVisible(Control)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
'功能：根据权限、当前病人或数据情况，设置功能或可见和可用性
'  1.无病人的情况
'  2.病人已出院的情况
'  3.无数据的情况
    Dim vRoll As TYPE_AdviceRoll
    Dim blnAdvice As Boolean, blnEnabled As Boolean, blnEdit As Boolean, bln补录 As Boolean
    Dim i As Integer
    
    tbcMain.Enabled = mlng病人ID <> 0
    For i = 0 To tbcMain.ItemCount - 1
        tbcMain.Item(i).Enabled = mlng病人ID <> 0
    Next
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    'Pass
    '如果此处不控制，当 control.Id 满足于[conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99]这个区间 时,下面医嘱操作部分和按钮可见状态中会改变的Pass
    'Enabled属性值。这样在独立部件中设置的enabled的值将会被覆盖。
    If Between(Control.ID, conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99) Then
        Control.Visible = IIF(Control.Category <> "", InStr(Control.Category, ";可见;") > 0, True)
        Control.Enabled = IIF(Control.Category <> "", InStr(Control.Category, ";可用;") > 0, True)
        Exit Sub
    End If
    
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    '医嘱操作部份
    '------------------------------------------------------------------------------
    '总的判断:无病人或已会诊病人不允许任何操作
    If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 998) _
        Or Between(Control.ID, conMenu_Edit_NewItem * 10#, (conMenu_Edit_NewItem + 998) * 10# + 9) Or Control.ID = conMenu_Manage_ThingAudit Then  '包含二级子菜单
        
        Control.Enabled = mlng病人ID <> 0 And mintPState <> ps已诊 And mintPState <> ps待转入 _
            And (InStr(",0,3,", mint执行状态) > 0 Or Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Or Control.ID = conMenu_Edit_MarkKeyMap Or Control.ID = conMenu_Edit_Compend _
                Or Between(Control.ID, conMenu_Edit_Compend * 10# + 1, conMenu_Edit_Compend * 10# + 6))
        If Not Control.Enabled Then Exit Sub
    End If
    
    blnAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
    blnEdit = (mintPState = ps在院 Or mintPState = ps待诊 Or mintPState = ps出院)
    bln补录 = (mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院)
    
    Select Case Control.ID
    Case conMenu_Edit_NewItem, conMenu_Edit_LISApply, conMenu_Edit_PacsApply, conMenu_Edit_BloodApply, conMenu_Edit_OperationApply, conMenu_Edit_ConsultationApply, conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101# '新开医嘱(本次住院正常退出路径的病人不能再新开和修改)
        Control.Enabled = (blnEdit Or bln补录)
        
    Case conMenu_Edit_Sort  '调整顺序
        Control.Enabled = blnEdit
    Case conMenu_Edit_Modify, conMenu_Edit_Delete '修改医嘱,删除医嘱
        With vsAdvice
            blnEnabled = blnAdvice
            If blnEnabled Then
                If Control.ID = conMenu_Edit_Modify Then
                    blnEnabled = (blnEdit Or bln补录)
                End If
            End If
            If blnEnabled Then
                If InStr(",1,2,", .TextMatrix(.Row, COL_医嘱状态)) = 0 Then blnEnabled = False
            End If
            If blnEnabled Then
                If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then blnEnabled = False
            End If
            If blnEnabled Then
                '临床和医技不能互相操作
                If mint场合 = 2 Then
                    blnEnabled = InStr("," & mstr前提IDs & ",", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) & ",") > 0
                Else
                    blnEnabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) = 0
                End If
                
            End If
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Manage_RecipeAuditView
        blnEnabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_处方审查状态)) <> 0
        Control.Enabled = blnEnabled
    '检验申请单，修改
    Case conMenu_Edit_ApplyModi
        With vsAdvice
            blnEnabled = blnAdvice
            If blnEnabled Then
                If Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E" Then
                    If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E") Then blnEnabled = False
                ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" And .TextMatrix(.Row, COL_操作类型) <> "病理" Then
                    If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And .TextMatrix(.Row, COL_诊疗类别) = "D") Then blnEnabled = False
                ElseIf .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                    If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And .TextMatrix(.Row, COL_诊疗类别) = "K") Then blnEnabled = False
                ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                    If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And .TextMatrix(.Row, COL_诊疗类别) = "F") Then blnEnabled = False
                ElseIf Val(.TextMatrix(.Row, COL_操作类型)) = 7 And .TextMatrix(.Row, COL_诊疗类别) = "Z" Then
                    If Not .TextMatrix(.Row, COL_医嘱状态) = "1" Then blnEnabled = False
                Else
                    blnEnabled = Val(.TextMatrix(.Row, COL_申请序号)) <> 0
                End If
                
            End If
            Control.Enabled = blnEnabled
        End With
    Case conMenu_Edit_NewRis
        blnEnabled = False
        With vsAdvice
            If InStr(",D,F,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_操作类型))) > 0 And .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_期效) = "临嘱" Then
                blnEnabled = True
            End If
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisSch
        blnEnabled = False
        If gbln启用影像信息系统预约 Then
            With vsAdvice
                If (InStr(",D,F,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_操作类型))) > 0 And .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_期效) = "临嘱") And Val(.TextMatrix(.Row, COL_RIS预约ID)) = 0 Then
                    blnEnabled = True
                End If
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisDel, conMenu_Tool_RisPrint
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS预约ID)) <> 0
    Case conMenu_Edit_NewRisModi
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS预约ID)) <> 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 8
    Case conMenu_Tool_RisPrintBat
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Edit_ViewDrugExplain '查看药品说明书
        Control.Enabled = blnAdvice And InStr(",5,6,7,", vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别)) > 0
    '申请单取消
    Case conMenu_Edit_ApplyDel
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (.TextMatrix(.Row, COL_医嘱状态) = "1" And _
                    (.TextMatrix(.Row, COL_诊疗类别) = "D" Or .TextMatrix(.Row, COL_诊疗类别) = "F" Or Val(.TextMatrix(.Row, COL_操作类型)) = 7 And .TextMatrix(.Row, COL_诊疗类别) = "Z" Or _
                        Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E" Or _
                        .TextMatrix(.Row, COL_诊疗类别) = "K")) Then
                    blnEnabled = Val(.TextMatrix(.Row, COL_申请序号)) <> 0
                End If
                '用血医嘱待审核不允许取消（新血库流程数据）
                If blnEnabled = True And .TextMatrix(.Row, COL_诊疗类别) = "K" And .TextMatrix(.Row, COL_医嘱状态) = "1" Then
                    If Val(.TextMatrix(.Row, COL_检查方法)) = 1 And Val(.TextMatrix(.Row, COL_审核状态)) = 1 Then blnEnabled = False
                End If
            End With
        End If
        Control.Enabled = blnEnabled
    '查看申请
    Case conMenu_Edit_ApplyView
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (InStr(",F,K,D,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Or Val(.TextMatrix(.Row, COL_操作类型)) = 7 And .TextMatrix(.Row, COL_诊疗类别) = "Z") Then blnEnabled = Val(.TextMatrix(.Row, COL_申请序号)) <> 0
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_UnUse '标记未用医嘱
        With vsAdvice
            Control.Checked = Val(.TextMatrix(.Row, COL_执行标记)) = -1
            If Val(.TextMatrix(.Row, COL_执行标记)) = -1 Then
                Control.Enabled = True
            Else
                blnEnabled = blnAdvice
                If blnEnabled Then
                    blnEnabled = mintPState = ps在院 Or mintPState = ps预出 Or mintPState = ps待诊
                End If
                If blnEnabled Then '未校对、已作废的医嘱不允许标记
                    If InStr(",1,2,4,", .TextMatrix(.Row, COL_医嘱状态)) > 0 Then blnEnabled = False
                End If
                If blnEnabled Then '已发送的长嘱不允许标记
                    If .TextMatrix(.Row, COL_期效) = "长嘱" And .TextMatrix(.Row, COL_上次执行) <> "" Then blnEnabled = False
                End If
                If blnEnabled Then '已有皮试结果的不允许标记(相当于执行了)
                    If .TextMatrix(.Row, COL_皮试) <> "" Then blnEnabled = False
                End If
                Control.Enabled = blnEnabled
            End If
        End With
    Case conMenu_Edit_Stop, conMenu_Edit_Blankoff '停止医嘱,医嘱作废
        If mint场合 = 2 Then '医技医生操作
            With vsAdvice
                blnEnabled = blnAdvice _
                    And InStr(",1,2,4,8,9,", Val(.TextMatrix(.Row, COL_医嘱状态))) = 0 _
                    And (Val(.TextMatrix(.Row, COL_签名否)) = 0 Or Not gobjESign Is Nothing) _
                    And InStr("," & mstr前提IDs & ",", "," & Val(.TextMatrix(.Row, COL_前提ID)) & ",") > 0 _
                    And .TextMatrix(.Row, COL_开嘱医生) = UserInfo.姓名
                
                If blnEnabled Then
                    If Control.ID = conMenu_Edit_Stop Then
                        '长嘱(不含中药配方)
                        blnEnabled = .TextMatrix(.Row, COL_期效) = "长嘱" And .TextMatrix(.Row, COL_总量) = ""
                    ElseIf Control.ID = conMenu_Edit_Blankoff Then
                        '未发送才可作废
                        blnEnabled = .TextMatrix(.Row, COL_上次执行) = ""
                    End If
                End If
            End With
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6 '查阅、打印报告
        If Not gobjExchange Is Nothing Then
            Control.Enabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) <> 0 And vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态) <> "4"
        Else
            Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_检查报告ID) <> "" Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS报告ID)) <> 0) And vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态) <> "4"
        End If
        If Control.ID = conMenu_Edit_Compend * 10# + 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        ElseIf Control.ID = conMenu_Edit_Compend * 10# + 6 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 2 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_Compend * 10# + 4 '我已经查阅该报告
        Control.Checked = Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_查阅状态)) = 1
        Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_检查报告ID) <> "")
    Case conMenu_Edit_Compend * 10# + 5 '自动标记查阅状态
        Control.Checked = mblnAutoRead
        Control.Enabled = mblnAutoReadEnabled
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs '观片处理
        blnEnabled = blnAdvice And InStr(",4,5,6,7,8,9,H,M,Z,", vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别)) = 0 ' And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID)) <> 0
        If blnEnabled Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) <> 8 Then
                blnEnabled = False
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Audit, conMenu_Edit_Price '医嘱校对(医嘱审核),计价调整
        Control.Enabled = (blnEdit Or bln补录)
    Case conMenu_Edit_Pause, conMenu_Edit_Reuse '医嘱暂停,医嘱启用
        Control.Enabled = (mintPState = ps在院 Or mintPState = ps待诊)
    
    Case conMenu_Edit_Untread '医嘱回退(子项在弹出时已设置可用状态)
        Control.Enabled = UBound(marrRollList) >= 1 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_执行标记)) <> -1
        If Control.Enabled And Not RollFirstEnabled Then
            Control.IconId = conMenu_Edit_Untread * 100# + 99 '表示有但不可以回退
        Else
            Control.IconId = conMenu_Edit_Untread
        End If
    Case conMenu_Edit_AdvicePrice
        Control.Enabled = (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) > 4 Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 3)
    Case conMenu_Edit_ChargeDelApply, conMenu_Edit_ChargeDelAudit '销帐申请审核
        Control.Enabled = mlng病区ID <> 0
    Case conMenu_Edit_ChargeOff * 10# + 1 To conMenu_Edit_ChargeOff * 10# + 3 '直接冲销
        blnEnabled = False
        If tbcAppend.Selected.Tag = "发送" And mblnAppend Then
            If Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("发送号"))) <> 0 Then
                If Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("记录性质"))) = 2 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnAdvice And blnEnabled
    Case conMenu_Edit_Test '皮试结果:发送后才能标注
        With vsAdvice
            Control.Enabled = blnAdvice And Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 _
                And .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "1"
        End With
    Case conMenu_Edit_ReStop, conMenu_Edit_ClearUp '确认停止,医嘱重整
    Case conMenu_Edit_NoPrint '屏蔽打印
        Control.Enabled = blnAdvice And Control.Visible
        If Control.Enabled Then
            Control.Checked = Val(vsAdvice.ValueMatrix(vsAdvice.Row, COL_屏蔽打印)) = 1
        End If
    Case conMenu_Edit_Send  '发送
        If mint场合 <> 1 Then '医生医技临嘱发送
            Control.Enabled = (blnEdit Or bln补录)
        End If
    Case conMenu_Edit_SendBilling
        If InStr(GetInsidePrivs(p住院医嘱下达), "发送门诊费用") > 0 Then
            If mlng病人性质 = 1 Then    '门诊留观病人，只能发送为门诊记帐单
                Control.Caption = "门诊记帐"
            Else
                Control.Caption = "住院记帐"
            End If
        End If
    Case conMenu_Edit_TraReaction
        With vsAdvice
            Control.Enabled = (.TextMatrix(.Row, COL_诊疗类别) = "K") And Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 And gbln血库系统
        End With
    Case conMenu_Edit_CriticalAdvice
        blnEnabled = False
        If Not mrs危急值 Is Nothing Then
            If Not mrs危急值.EOF Then
                blnEnabled = True
            End If
        End If
        If blnEnabled Then
            blnEnabled = (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) <> 4 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0)
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_MediAudit '药嘱审查(中药不显示)
        If mblnPass Then
            Call gobjPass.zlPassCommandBarUpdate(mobjPassMap, Control, blnAdvice)
        End If
    Case conMenu_Edit_MeetArrive
        Control.Caption = IIF(mbln确认会诊, "取消会诊医生到场(&K)", "确认会诊医生到场(&M)")
        Control.Enabled = True
    End Select
    
    '病人范围医嘱下达检查
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        '新开医嘱
        If Control.Enabled Then Control.Enabled = PatiCanAdvice
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        '修改医嘱,删除医嘱
        If Control.Enabled And mint场合 = 2 Then
            Control.Enabled = PatiCanAdvice
        ElseIf Control.Enabled Then
            Control.Enabled = PatiCanAdvice
        End If
    Case conMenu_Edit_ClearUp, conMenu_Edit_Untread
        '医嘱重整,医嘱回退
        If mint场合 = 0 Then
            If Control.Enabled Then Control.Enabled = PatiCanAdvice
        End If
    End Select
            
    '医嘱报表部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Report_ClinicBill '打印诊疗单据
        Control.Enabled = Control.CommandBar.Controls.Count > 0
    Case conMenu_Report_Reports, conMenu_Report_DrugQuery, conMenu_Report_MultiBill '其它固定报表
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Report_AdviceBill1 To conMenu_Report_AdviceBill3 '长期医嘱单,临时医嘱单,病人医嘱本
        Control.Enabled = mlng病人ID <> 0
    End Select
    
    '电子签名部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_SignNew '医嘱签名
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Tool_SignVerify '验证签名
        blnEnabled = mlng病人ID <> 0 And blnAdvice And tbcAppend.Selected.Tag = "签名" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        If blnEnabled Then blnEnabled = vsAppend.RowData(vsAppend.Row) <> 0
        Control.Enabled = blnEnabled
    Case conMenu_Tool_SignEarse '取消签名
        blnEnabled = mlng病人ID <> 0 And blnAdvice And tbcAppend.Selected.Tag = "签名" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        If blnEnabled Then blnEnabled = vsAppend.RowData(vsAppend.Row) <> 0 And vsAppend.Cell(flexcpData, vsAppend.Row, 0) <> 3
        Control.Enabled = blnEnabled
    End Select
    
    '其它部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = blnAdvice
    Case conMenu_View_Append '附加信息
        Control.Checked = tbcAppend.Visible
    Case conMenu_View_AdviceLost '医嘱是否定位最后
        Control.Checked = mbln医嘱定位最后
    Case conMenu_View_Hide '自动隐藏过滤工具栏
        Control.Checked = mblnHideFilter
    Case conMenu_Manage_ReportLisView, conMenu_Manage_ReportPacsView  '检验报告浏览
        Control.Enabled = mlng病人ID <> 0
    Case conMenu_Manage_ReportPrint  '检验报告批量打印
        Control.Enabled = mlng病区ID <> 0
    Case conMenu_Report_BloodInstant  '执行单打印
        Control.Visible = InStr(GetInsidePrivs(9005, , 2200), ";输血执行打印;") <> 0
        Control.Enabled = vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "K" And Control.Visible
    End Select
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strItem As String, blnSendPriv As Boolean

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" And Control.ID <> conMenu_Edit_SendBilling And Control.ID <> conMenu_Edit_Audit And Control.ID <> conMenu_Edit_MeetArrive Then Exit Sub

    blnVisible = True
    
    '身份权限判断
    '------------------------------------------------------------------------------
    If mint场合 = 0 And InStr(UserInfo.性质, "医生") = 0 _
        Or mint场合 = 1 And InStr(UserInfo.性质, "护士") = 0 Then
        If Control.ID = conMenu_EditPopup Then blnVisible = False
        If Control.ID = conMenu_ReportPopup Then blnVisible = False
        If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 999) Then blnVisible = False
    End If
    
    '医嘱操作部份
    '------------------------------------------------------------------------------
    If mint场合 = 0 Or mint场合 = 2 Then
        Select Case Control.ID
        Case conMenu_Edit_Untread
            '医嘱回退
            If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱下达;") = 0 Then blnVisible = False
        'Case conMenu_Edit_Send  '没有临嘱发送权限时（也就没有发送门诊收费权限），是创建的conMenu_Edit_SendBilling
         
        Case conMenu_Edit_SendBilling
            If mlng病人性质 = 1 Then  '发送住院记帐没有单独控制权限
                If InStr(GetInsidePrivs(p住院医嘱下达), ";发送门诊记帐;") = 0 Then blnVisible = False
            Else
                If InStr(GetInsidePrivs(p住院医嘱下达), ";临嘱发送;") = 0 Then blnVisible = False
            End If
        Case conMenu_Edit_SendCharge
            If InStr(GetInsidePrivs(p住院医嘱下达), ";发送门诊费用;") = 0 Then blnVisible = False
        Case conMenu_Edit_BatExecute
            '医嘱批量执行
            If InStr(GetInsidePrivs(p住院医嘱下达), ";批量执行登记;") = 0 Then blnVisible = False
        Case conMenu_Edit_NoPrint
            If InStr(GetInsidePrivs(p住院医嘱下达), ";屏蔽打印;") = 0 Then
                blnVisible = False
            Else
                blnVisible = True
            End If
            Control.Enabled = blnVisible
        Case conMenu_Edit_TraReaction  '输血反应登记
            If gbln血库系统 Then '老血库系统默认显示
                If InStr(GetInsidePrivs(9005, , 2200), ";输血反应登记;") = 0 Then
                    blnVisible = False
                Else
                    blnVisible = True
                End If
                Control.Enabled = blnVisible
            End If
        End Select

    ElseIf mint场合 = 1 Then
        strItem = GetInsidePrivs(p住院医嘱发送)
        blnSendPriv = InStr(strItem, ";发送药疗临嘱;") > 0 Or InStr(strItem, ";发送药疗长嘱;") > 0 _
                        Or InStr(strItem, ";发送其他临嘱;") > 0 Or InStr(strItem, ";发送其他长嘱;") > 0
                
        Select Case Control.ID
        Case conMenu_Edit_Untread
            '医嘱回退
            If InStr(strItem, ";医嘱状态回退;") = 0 Then blnVisible = False
        Case conMenu_Edit_Send
            '医嘱发送
            If Not blnSendPriv Then blnVisible = False
        Case conMenu_Edit_BatExecute, conMenu_Manage_ThingAudit
            '医嘱批量执行
            If InStr(strItem, ";批量执行登记;") = 0 Then blnVisible = False
            If blnVisible And Control.ID = conMenu_Manage_ThingAudit Then
                If Val(gstr医嘱核对) = 0 Then blnVisible = False
            End If
        Case conMenu_Edit_MeetArrive
            With vsAdvice
                blnVisible = Val(.TextMatrix(.Row, COL_申请序号)) <> 0 And Val(.TextMatrix(.Row, COL_操作类型)) = 7 And .TextMatrix(.Row, COL_诊疗类别) = "Z" And .TextMatrix(.Row, COL_状态) = "停止"
            End With
        Case conMenu_Edit_NoPrint
            If InStr(strItem, ";屏蔽打印;") = 0 Then
                blnVisible = False
            Else
                blnVisible = True
            End If
            Control.Enabled = blnVisible
        End Select
    End If
    
    Select Case Control.ID
    Case conMenu_Edit_ClearUp
        '医嘱重整
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱重整;") = 0 Then blnVisible = False
    Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Sort, conMenu_Edit_LISApply, conMenu_Edit_ApplyDel, conMenu_Edit_ApplyModi, conMenu_Edit_Apply, conMenu_Edit_ApplyDel, conMenu_Edit_ApplyView
        '新开医嘱,修改医嘱,删除医嘱 ,调整顺序,检验申请、修改、删除
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱下达;") = 0 Then blnVisible = False
        
    Case conMenu_Edit_UnUse '未用医嘱
        If InStr(GetInsidePrivs(p住院医嘱下达), ";标记未用医嘱;") = 0 Then blnVisible = False
    Case conMenu_Edit_Pause, conMenu_Edit_Reuse
        '医嘱暂停,医嘱启用
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱暂停;") = 0 Then blnVisible = False
    Case conMenu_Edit_Stop
        '停止医嘱
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱停止;") = 0 Then blnVisible = False
    Case conMenu_Edit_Blankoff
        '医嘱作废
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱作废;") = 0 Then blnVisible = False
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1, conMenu_Edit_Compend * 10# + 6
        '报告弹出(含打印),查阅报告
        If InStr(GetInsidePrivs(p住院医嘱下达), ";报告查阅;") = 0 Then blnVisible = False
    Case conMenu_Edit_Compend * 10# + 2, conMenu_Edit_Compend * 10# + 3
        '打印报告
        If InStr(GetInsidePrivs(p住院医嘱下达), ";报告打印;") = 0 Then blnVisible = False
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs
        '观片处理
        If GetInsidePrivs(pXWPACS观片) <> "" And InStr(GetInsidePrivs(p住院医嘱下达), ";观片处理;") <> 0 Then
            blnVisible = True
        Else
            If Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Then
                If InStr(GetInsidePrivs(p住院医嘱下达), ";观片处理;") = 0 Or GetInsidePrivs(p观片工具管理) = "" Then
                    blnVisible = False
                End If
            Else
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_Audit
        If mint场合 = 1 Then
            '医嘱校对
            If Val(zlDatabase.GetPara("发送前自动校对", glngSys, p住院医嘱发送, 0)) = 1 Then
                If InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱校对处理;") > 0 And Not blnSendPriv Then
                    '门诊留观病人的长嘱（膳食，营养等），只校对不发送
                    blnVisible = True
                Else
                    blnVisible = False
                End If
                Control.Enabled = blnVisible
            Else
                If InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱校对处理;") = 0 Then blnVisible = False
            End If
        Else
            '医嘱审核:无权限或不具有资格时
            If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱审核;") = 0 Or Not mblnHaveAuditPriv Then blnVisible = False
        End If
    Case conMenu_Edit_StopAudit
        '停嘱审核和新嘱审核公用权限
        If InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱审核;") = 0 Or Not mblnHaveAuditPriv Then blnVisible = False
    Case conMenu_Edit_Price
        '计价调整
        If InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱校对处理;") = 0 Then blnVisible = False
    Case conMenu_Edit_ReStop
        '确认停止
        If InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱确认停止;") = 0 Then blnVisible = False
    Case conMenu_Edit_Test
        '皮试结果
        If InStr(GetInsidePrivs(p住院医嘱发送), ";皮试医嘱结果;") = 0 Then blnVisible = False
    Case conMenu_Edit_SendBack
        '超期限送收回
        If InStr(GetInsidePrivs(p住院医嘱发送), ";超期发送收回;") = 0 Then blnVisible = False
    Case conMenu_Edit_ViewDrugExplain '查看药品说明书
        If gobjDrugExplain Is Nothing Or InStr(GetInsidePrivs(p住院医嘱下达), ";药品说明书;") = 0 Then blnVisible = False
    Case conMenu_Edit_ChargeDelApply
        '销帐申请
        '55380
        strItem = GetInsidePrivs(p住院记帐操作)
        If InStr(strItem, ";药品销帐申请;") = 0 _
            And InStr(strItem, ";卫材销帐申请;") = 0 _
            And InStr(strItem, ";诊疗销帐申请;") = 0 Then blnVisible = False
    Case conMenu_Edit_ChargeDelAudit
        '销帐审核
        strItem = GetInsidePrivs(p住院记帐操作)
        If InStr(strItem, ";销帐审核;") = 0 Then blnVisible = False
    Case conMenu_Edit_Surplus
        '药品留存登记
        strItem = GetInsidePrivs(p住院医嘱发送)
        If InStr(strItem, ";药品留存登记;") = 0 Then blnVisible = False
    Case conMenu_Edit_MediAudit
        '合理用药审查
        strItem = GetInsidePrivs(p住院医嘱下达)
        If InStr(strItem, "合理用药监测") = 0 Then blnVisible = False
    End Select
    '医嘱报表部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Report_DrugQuery '药疗收发查询
        If InStr(GetInsidePrivs(p住院医嘱发送), ";药疗收发查询;") = 0 Then blnVisible = False
    Case conMenu_Report_AdviceBill1 '长期医嘱单,临时医嘱单
        blnVisible = False
        If InStr(UserInfo.性质, "医生") > 0 Then
            If InStr(GetInsidePrivs(p住院医嘱下达), "长期医嘱单") > 0 Or InStr(GetInsidePrivs(p住院医嘱下达), "临时医嘱单") > 0 Then
                blnVisible = True
            End If
        End If
        If Not blnVisible Then
            If InStr(UserInfo.性质, "护士") > 0 Then
                If InStr(GetInsidePrivs(p住院医嘱发送), "长期医嘱单") > 0 Or InStr(GetInsidePrivs(p住院医嘱发送), "临时医嘱单") > 0 Then
                    blnVisible = True
                End If
            End If
        End If
    Case conMenu_Report_AdviceBill3 '病人医嘱本
        blnVisible = False
        If InStr(UserInfo.性质, "医生") > 0 Then
            If InStr(GetInsidePrivs(p住院医嘱下达), "病人医嘱本") > 0 Then
                blnVisible = True
            End If
        End If
        If Not blnVisible Then
            If InStr(UserInfo.性质, "护士") > 0 Then
                If InStr(GetInsidePrivs(p住院医嘱发送), "病人医嘱本") > 0 Then
                    blnVisible = True
                End If
            End If
        End If
    End Select
    
    Control.Category = "已判断"
    '电子签名部份
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_Sign, conMenu_Tool_SignNew '电子签名,医嘱签名
        If gobjESign Is Nothing Or (InStr(GetInsidePrivs(p住院医嘱下达), ";医嘱下达;") = 0 And InStr(GetInsidePrivs(p住院医嘱发送), ";医嘱校对处理;") = 0) _
            Or Not mblnHaveAuditPriv Then
            blnVisible = False
        ElseIf mblnSignVisible = False Then
            blnVisible = False '不同场合没有设置要使用签名
        End If
        Control.Category = ""  '签名按钮动态判断可见性
    End Select

    Control.Enabled = blnVisible
    Control.Visible = blnVisible
End Sub

Private Function PatiCanAdvice() As Boolean
'功能：检查对当前病人是否可以下达医嘱
'说明：主要是有下达的权限时,再检查本科和全院病人范围
    Dim strPriv As String, bln下达 As Boolean
    
    strPriv = GetInsidePrivs(p住院医嘱下达)
    If mlng病人ID <> 0 Then
        If mintPState = ps待诊 Then
            bln下达 = True '待会诊病人是允许的
        ElseIf mint场合 = 0 Then
            If mstr住院医生 = UserInfo.姓名 Then
                bln下达 = True '当前医生经治病人
            ElseIf InStr(strPriv, ";全院医嘱下达;") > 0 Then
                bln下达 = True '有全院病人医嘱下达权限
            ElseIf InStr(strPriv, ";本科医嘱下达;") > 0 _
                And InStr("," & mstr部门IDs & ",", "," & mlng科室ID & ",") > 0 Then
                bln下达 = True '有本科病人医嘱下达权限
            End If
        ElseIf mint场合 = 1 Then
            If mstr责任护士 = UserInfo.姓名 Then
                bln下达 = True '当前护士责任护理
            ElseIf InStr(strPriv, ";全院医嘱下达;") > 0 Then
                bln下达 = True '有全院病人医嘱下达权限
            ElseIf InStr(strPriv, ";本科医嘱下达;") > 0 _
                And InStr("," & mstr部门IDs & ",", "," & mlng病区ID & ",") > 0 Then
                bln下达 = True '有本科病人医嘱下达权限
            End If
        Else
            bln下达 = True '其它场合暂不限制
        End If
    Else
        bln下达 = True '当作可以
    End If
    PatiCanAdvice = bln下达
End Function

Public Sub zlRefresh(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng病区ID As Long, ByVal lng科室id As Long, _
    ByVal int状态 As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal lng前提ID As Long, _
    Optional ByVal int执行状态 As Integer, Optional ByVal lng界面科室ID As Long, Optional ByVal lng路径状态 As Long = -1, _
    Optional ByVal lng医护科室ID As Long, Optional ByRef objMip As Object, Optional ByVal int婴儿 As Integer = -1, Optional ByVal lng前提科室ID As Long, Optional ByVal lng会诊医嘱ID As Long)
'功能：刷新住院医嘱数据
'参数：int类型=病人的不同类型
'      lng前提ID=当由医技站调用时传入
'      lng病区ID，lng科室ID=当“5-最近转科病人”时为病人原病区或原科室
'      lng界面科室ID=如果当前医生站是会诊病人，则为会诊科室ID；如果是医技站调用，则为医技科室ID
'      int状态=当由医技站调用时传入,项目的执行状态
'      lng路径状态=-1:未导入,0-不符合导入条件，1-执行中，2-正常结束，3-变异结束
'      blnMoved=该病人的数据是否已转出
'      lng医护科室ID=医护站界面科室ID
'      lng前提科室ID= lng前提ID这条医嘱对应的执行科室ID；当医技站调用且满这个条件时：lng界面科室ID<>lng前提科室ID  lng前提科室ID参数必须传入
'      lng会诊医嘱ID 住院医生工作站处理会诊医嘱时，选中的会诊医嘱ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objControl As CommandBarControl
    Dim lngPre病人ID As Long
    Dim lngPre科室ID As Long, lngPre病区ID As Long
    Dim lngPre界面科室ID As Long
    Dim strPrivs As String
    
    lngPre病人ID = mlng病人ID
    lngPre科室ID = mlng科室ID
    lngPre界面科室ID = mlng界面科室ID
    lngPre病区ID = mlng病区ID
    mlng会诊医嘱ID = lng会诊医嘱ID
    
    mintPState = int状态: mblnMoved = blnMoved
    mlng病人ID = lng病人ID: mlng主页ID = lng主页ID
    mlng病区ID = lng病区ID: mlng科室ID = lng科室id
    mlng前提ID = lng前提ID: mint执行状态 = int执行状态
    mlng界面科室ID = lng界面科室ID
    mlng医护科室ID = lng医护科室ID
    mlng路径状态 = lng路径状态
    mbyt婴儿 = 0
    
    If InitObjPublicExpense Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng病人ID, mlng主页ID, "", mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
    End If
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    If Visible Or mblnInsideTools Then
        mblnSignVisible = True
        If mint场合 = 0 Then
            If CheckSign(1, 0, mlng界面科室ID, mlng科室ID, 2, False, gobjESign) = False Then
                mblnSignVisible = False '不同场合没有设置要使用签名
            End If
        ElseIf mint场合 = 2 Then
            If CheckSign(3, 0, mlng界面科室ID, mlng科室ID, 2, False, gobjESign) = False Then
                mblnSignVisible = False '不同场合没有设置要使用签名
            End If
        ElseIf mint场合 = 1 Then
            If CheckSign(2, mlng医护科室ID, , , , False, gobjESign) = False Then
                mblnSignVisible = False '不同场合没有设置要使用签名
            End If
        End If
    End If
    
    '读取一些额外的信息
    If mlng病人ID <> 0 And lngPre病人ID <> mlng病人ID Then
        On Error GoTo errH
        strSQL = "Select a.入院日期, a.住院医师, a.责任护士, a.病案状态, a.病人性质, a.险类, a.婴儿科室id, a.婴儿病区id, a.住院号, b.姓名,b.当前床号,b.性别,a.医嘱重整时间 as 重整" & _
            " From 病案主页 A, 病人信息 B Where a.病人id = b.病人id And a.病人id = [1] And a.主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mlng病人ID, mlng主页ID)
        mvInDate = rsTmp!入院日期
        mstr住院医生 = NVL(rsTmp!住院医师)
        mstr责任护士 = NVL(rsTmp!责任护士)
        mint病案状态 = NVL(rsTmp!病案状态, 0)
        mlng病人性质 = Val("" & rsTmp!病人性质)
        mint险类 = Val("" & rsTmp!险类)
        mlng婴儿科室ID = NVL(rsTmp!婴儿科室ID, 0)
        mlng婴儿病区ID = NVL(rsTmp!婴儿病区ID, 0)
        mstr姓名 = rsTmp!姓名 & ""
        mstr住院号 = rsTmp!住院号 & ""
        mstr床号 = rsTmp!当前床号 & ""
        mstr性别 = rsTmp!性别 & ""
        mdat重整 = NVL(rsTmp!重整, CDate("1900-01-01"))
        
        '读取婴儿信息
        mstr婴儿 = GetBabyRegList(lng病人ID, lng主页ID)
        If mstr婴儿 <> "" Then
            '读取最近缺省值：-1=所有,0=病人,1-婴儿1
            mvarCond.婴儿 = Val(zlDatabase.GetPara("病人婴儿过滤", glngSys, p住院医嘱下达, "0"))
            If mvarCond.婴儿 > UBound(Split(mstr婴儿, "<Split>")) + 1 Then mvarCond.婴儿 = 0
            If mvarCond.婴儿 <> -1 Then mbyt婴儿 = mvarCond.婴儿
        End If
        Call GetCriticalData
        On Error GoTo 0
    ElseIf mlng病人ID = 0 Then
        mvInDate = CDate(0)
        mstr住院医生 = ""
        mstr责任护士 = ""
        mint病案状态 = 0
        mlng病人性质 = 0
        mstr婴儿 = ""
        mlng婴儿科室ID = 0
        mlng婴儿病区ID = 0
    End If
    
    If int婴儿 <> -1 And mstr婴儿 <> "" Then
        mbyt婴儿 = int婴儿
        mvarCond.婴儿 = mbyt婴儿
        mlngBaby = mbyt婴儿
    End If
    
    If (lngPre界面科室ID <> mlng界面科室ID Or lngPre病人ID <> mlng病人ID) And mlng前提ID <> 0 Then
        mstr前提IDs = Get医技科室医嘱IDs(mlng病人ID, mlng主页ID, IIF(0 = lng前提科室ID, mlng界面科室ID, lng前提科室ID), True, mlng前提ID)
    ElseIf mlng前提ID = 0 Then
        mstr前提IDs = ""
    End If
    
    If mstr部门IDs = "" Then
        If mint场合 = 0 Then
            mstr部门IDs = GetUser科室IDs(True)
        ElseIf mint场合 = 1 Then
            mstr部门IDs = GetUser病区IDs
        End If
    End If
    'PASS 合理用药检测 病人信息发生变动
    If lngPre病人ID <> mlng病人ID Then
        If mblnPass Then
            Call zlPASSPati
            On Error Resume Next
            Call gobjPass.zlPassClearLight(mobjPassMap)
            On Error GoTo 0
        End If
    End If
    
    '修改发送菜单
    If mint场合 = 1 And gstr输液配置中心 <> "" And lngPre病区ID <> mlng病区ID Then
        strPrivs = GetInsidePrivs(p住院医嘱发送)
        If Not (InStr(";" & strPrivs & ";", ";发送药疗临嘱;") = 0 Or InStr(";" & strPrivs & ";", ";发送药疗长嘱;") = 0) Then
            Call SetSendCommandBar
        End If
    End If
    
    If Not grsTube Is Nothing Then
        If grsTube.State = 1 Then grsTube.Close
        Set grsTube = Nothing
    End If
    
    '刷新数据
    Call RefreshData
    
    '执行自动插件功能：病人ID=0也调用，以实现如关闭界面
    If mlngPlugInID <> 0 And lngPre病人ID <> mlng病人ID Then
        If mblnInsideTools Then
            Set objControl = cbsSub.FindControl(, mlngPlugInID, , True)
        Else
            Set objControl = mcbsMain.FindControl(, mlngPlugInID, , True)
        End If
        If Not objControl Is Nothing Then
            objControl.Execute
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlInitMip(ByRef objMip As Object)
'功能：消息对象
'参数：objMip zl9ComLib.clsMipModule
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
End Sub

Public Sub zlItemRef()
'功能：调用诊疗参考
    Dim lng诊疗项目ID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_诊疗类别) = "E" And (RowIs配方行(.Row) Or RowIs检验行(.Row)) Then
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng诊疗项目ID = Get诊疗项目ID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    'ToDo:显示诊疗参考
    
End Sub

Public Function zlSeekAndViewEPRReport(ByVal lng报告ID As Long) As Boolean
'功能：定位到报告对应的医嘱，并打开报告查看
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngRow As Long
    
    strSQL = "Select 医嘱ID From 病人医嘱报告 Where 病历ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng报告ID)
    If Not rsTmp.EOF Then
        lngRow = vsAdvice.FindRow(CStr(rsTmp!医嘱ID), , COL_ID)
        If lngRow <> -1 Then vsAdvice.Row = lngRow
        
        '有权限则弹行打开，不管是否定位到及病人状态
        If InStr(GetInsidePrivs(p住院医嘱下达), ";报告查阅;") > 0 Then
            Select Case CheckEPRReport(rsTmp!医嘱ID, lng报告ID)
            Case 0
                MsgBox "该医嘱的报告没有书写！", vbInformation, gstrSysName
                Exit Function
            Case 2
                If InStr(GetInsidePrivs(p住院医嘱下达), "查阅未完成报告") > 0 Then
                    MsgBox "注意：该医嘱的报告还没有正式签名！", vbInformation, gstrSysName
                Else
                    MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，你没有权限操作！", vbInformation, gstrSysName
                    Exit Function
                End If
            End Select
            
            RaiseEvent ViewEPRReport(lng报告ID, False)
        End If
        
        zlSeekAndViewEPRReport = True
    Else
        MsgBox "没有找到报告对应的医嘱记录。", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingAudit() As Boolean
'功能：核对
    Dim bln输血皮试 As Boolean
    Dim strSQL As String
    Dim str核对人 As String
    Dim bln输血 As Boolean  '输血医嘱和输血途径
    Dim lngRow As Long
    Dim strCheckTime As String
    Dim blnDo As Boolean
    Dim strXML As String
    
    With vsAppend
        bln输血 = (.TextMatrix(.Row, COLSend("诊疗类别")) = "K" Or .TextMatrix(.Row, COLSend("诊疗类别")) = "E" And .TextMatrix(.Row, COLSend("操作类型")) = "8") And Mid(gstr医嘱核对, 1, 1) = "1"
        bln输血皮试 = (bln输血 Or .TextMatrix(.Row, COLSend("诊疗类别")) = "E" And .TextMatrix(.Row, COLSend("操作类型")) = "1" And Mid(gstr医嘱核对, 2, 1) = "1")
        If Not bln输血皮试 Then
            If Val(gstr医嘱核对) = 1 Then
                MsgBox "只能核对皮试医嘱。", vbInformation, gstrSysName
            ElseIf Val(gstr医嘱核对) = 10 Then
                MsgBox "只能核对输血医嘱。", vbInformation, gstrSysName
            Else
                MsgBox "只能核对输血或是皮试医嘱。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.Row, COLExec("核对人")) <> "" Then
            MsgBox "该医嘱已经核对，不能再次核对。", vbInformation, gstrSysName
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.Row, vsExec.FixedCols) = "" Then
            MsgBox "该医嘱还未进行执行情况登记，不能核对。", vbInformation, gstrSysName
            Exit Function
        End If
        str核对人 = zlDatabase.UserIdentifyByUser(Me, "在核对执行情况前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "执行情况登记", , True)
        If str核对人 = "" Then Exit Function
        
        If str核对人 = vsExec.TextMatrix(vsExec.Row, COLExec("执行人")) Then
            MsgBox "执行人不能和审核人相同，不能核对。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '获取核对时间
        strCheckTime = frmAdviceStopTime.ShowMe(mfrmParent, Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("医嘱ID"))), mlng科室ID, 1, Format(vsExec.TextMatrix(vsExec.Row, COLExec("登记时间")), "yyyy-MM-dd HH:mm"))
        
        If Not IsDate(strCheckTime) Then
            Exit Function
        End If
        
    End With
    With vsExec
        On Error GoTo errH
        lngRow = vsExec.Row
        
        '调用核对前外挂接口
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            blnDo = gobjPlugIn.AdvcieBeforToReview(glngSys, IIF(mint场合 = 0, p住院医生站, p住院护士站), mlng病人ID, mlng主页ID, Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("医嘱ID"))), Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("发送号"))), str核对人, strCheckTime, vsExec.TextMatrix(vsExec.Row, COLExec("执行人")) & "", strXML)
            Call zlPlugInErrH(err, "AdvcieBeforToReview")
            If 0 = err.Number Then '接口没有出错的情况下再判断接口的返回值
                If blnDo Then
                    strSQL = "Zl_病人医嘱核对_Insert(" & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("医嘱ID"))) & "," & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("发送号"))) & ",'" & str核对人 & "'" & _
                    IIF(bln输血, ",To_Date('" & Format(vsExec.Cell(flexcpData, vsExec.Row, COLExec("执行时间")), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", ",Null") & _
                    ",To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'))"
                Else
                    Exit Function
                End If
            End If
            If err.Number <> 0 Then err.Clear: Exit Function
            On Error GoTo 0
        Else
            strSQL = "Zl_病人医嘱核对_Insert(" & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("医嘱ID"))) & "," & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("发送号"))) & ",'" & str核对人 & "'" & _
            IIF(bln输血, ",To_Date('" & Format(vsExec.Cell(flexcpData, vsExec.Row, COLExec("执行时间")), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", ",Null") & _
            ",To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'))"
        End If
        
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "医嘱核对")
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '要更新执行状态'可能要更新执行状态
        vsExec.Row = lngRow
        FuncThingAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingDelAudit() As Boolean
'功能：取消核对
    Dim bln输血皮试 As Boolean
    Dim strSQL As String
    Dim str核对人 As String
    Dim bln输血 As Boolean '输血医嘱和输血途径
    Dim lngRow As Long
    
    With vsAppend
        bln输血 = (.TextMatrix(.Row, COLSend("诊疗类别")) = "K" Or .TextMatrix(.Row, COLSend("诊疗类别")) = "E" And .TextMatrix(.Row, COLSend("操作类型")) = "8") And Mid(gstr医嘱核对, 1, 1) = "1"
        bln输血皮试 = (bln输血 Or .TextMatrix(.Row, COLSend("诊疗类别")) = "E" And .TextMatrix(.Row, COLSend("操作类型")) = "1" And Mid(gstr医嘱核对, 2, 1) = "1")
        If Not bln输血皮试 Then
            If Val(gstr医嘱核对) = 1 Then
                MsgBox "只能取消核对皮试医嘱。", vbInformation, gstrSysName
            ElseIf Val(gstr医嘱核对) = 10 Then
                MsgBox "只能取消核对输血医嘱。", vbInformation, gstrSysName
            Else
                MsgBox "只能取消核对输血或是皮试医嘱。", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.Row, COLExec("核对人")) = "" Then
            MsgBox "该医嘱还未进行核对，不能取消。", vbInformation, gstrSysName
            Exit Function
        End If
        

    End With
    With vsExec
        If vsExec.TextMatrix(vsExec.Row, COLExec("核对人")) <> UserInfo.姓名 Then
            str核对人 = zlDatabase.UserIdentifyByUser(Me, "在取消核对前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "执行情况登记", , True)
            If str核对人 = "" Then Exit Function
            If str核对人 <> vsExec.TextMatrix(vsExec.Row, COLExec("核对人")) Then
                MsgBox "只能取消自己核对的医嘱，当前医嘱核对人是""" & vsExec.TextMatrix(vsExec.Row, COLExec("核对人")) & """", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If MsgBox("你确定要取消核对吗？", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
        End If
        On Error GoTo errH
        lngRow = vsExec.Row
        strSQL = "Zl_病人医嘱核对_Delete(" & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("医嘱ID"))) & "," & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("发送号"))) & _
            IIF(bln输血, ",To_Date('" & Format(vsExec.Cell(flexcpData, vsExec.Row, COLExec("执行时间")), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))", ")")
        Call zlDatabase.ExecuteProcedure(strSQL, "取消医嘱核对")
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '要更新执行状态'可能要更新执行状态
        vsExec.Row = lngRow
        FuncThingDelAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboTime_Click()
    Dim curDate As Date
    
    If cboTime.ListIndex = mintPreTime And mintPreTime <> 7 Then Exit Sub
    
    curDate = zlDatabase.Currentdate
    
    Select Case cboTime.Text
    Case "所有"
        mvarCond.开始时间 = CDate(0)
        mvarCond.结束时间 = CDate(0)
    Case "今天"
        mvarCond.开始时间 = Format(curDate, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "昨天"
        mvarCond.开始时间 = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
    Case "最近三天"
        mvarCond.开始时间 = Format(curDate - 2, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近一周"
        mvarCond.开始时间 = Format(curDate - 7, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近两周"
        mvarCond.开始时间 = Format(curDate - 14, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "最近一月"
        mvarCond.开始时间 = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
        mvarCond.结束时间 = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "[指定..]"
        If Not frmSelectTime.ShowMe(Me, mvarCond.开始时间, mvarCond.结束时间, cboTime, 1) Then
            '取消时恢复原来的选择
            Call zlControl.CboSetIndex(cboTime.hwnd, mintPreTime)
            If vsAdvice.Enabled Then vsAdvice.SetFocus
            Exit Sub
        Else
            If vsAdvice.Enabled Then vsAdvice.SetFocus
        End If
    End Select
        
    If mvarCond.开始时间 = CDate(0) Or mvarCond.结束时间 = CDate(0) Then
        cboTime.ToolTipText = ""
    Else
        cboTime.ToolTipText = "范围：" & Format(mvarCond.开始时间, "yyyy-MM-dd HH:mm:ss") & " 至 " & Format(mvarCond.结束时间, "yyyy-MM-dd HH:mm:ss")
    End If
    mintPreTime = cboTime.ListIndex
    Me.Refresh
    
    Call LoadAdvice
End Sub

Private Sub cbsExec_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Control.ID <> 0 Then
        If cbsExec.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
        Case ID_显示执行
            mblnShowExec = Not mblnShowExec
            Call SetExecShow(True, mblnShowExec)
            Call vsAppend_AfterRowColChange(-1, -1, vsAppend.Row, vsAppend.Col)
        Case ID_完成执行
            Call FuncExecFinish
        Case ID_取消完成
            Call FuncExecCancel
        Case ID_执行记录
            Call FuncThingNew
        Case ID_执行调整
            Call FuncThingModi
        Case ID_执行删除
            Call FuncThingDel
        Case ID_核对
            Call FuncThingAudit
        Case ID_取消核对
            Call FuncThingDelAudit
    End Select
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim int执行状态 As Integer, blnSelect As Boolean
    Dim str完成人 As String, int取消执行完成 As Integer
    
    If Not tbcAppend.Selected.Tag = "发送" Or Not picExec.Visible Then Exit Sub
    
    With vsAppend
        blnSelect = Val(.TextMatrix(.Row, COLSend("医嘱ID"))) <> 0
        If blnSelect Then '0-未执行,1-已执行,2-拒绝执行,3-正在执行
            int执行状态 = Val(.Cell(flexcpData, .Row, COLSend("执行状态")))
            str完成人 = .TextMatrix(.Row, COLSend("执行人"))
        End If
    End With
    
    Select Case Control.ID
        Case ID_显示执行
            Control.Checked = mblnShowExec
        Case ID_完成执行
            If InStr(GetInsidePrivs(p住院医嘱发送), "确认执行完成") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = blnSelect And (int执行状态 = 0 Or int执行状态 = 3)
            End If
        Case ID_取消完成
            int取消执行完成 = IIF(InStr(GetInsidePrivs(p住院医嘱发送), "取消执行完成") = 0, 0, 1) + IIF(InStr(GetInsidePrivs(p住院医嘱发送), "取消他人执行完成") = 0, 0, 2)
            
            If int取消执行完成 = 0 Then
                Control.Visible = False
            ElseIf int取消执行完成 = 1 Then
                Control.Enabled = blnSelect And int执行状态 = 1 And str完成人 = UserInfo.姓名
            ElseIf int取消执行完成 = 2 Then
                Control.Enabled = blnSelect And int执行状态 = 1 And str完成人 <> UserInfo.姓名
            ElseIf int取消执行完成 = 3 Then
                 Control.Enabled = blnSelect And int执行状态 = 1
            End If
        Case ID_执行记录
            If InStr(GetInsidePrivs(p住院医嘱发送), "执行情况登记") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnShowExec And blnSelect And (int执行状态 = 0 Or int执行状态 = 3)
            End If
        Case ID_执行调整, ID_执行删除
            If InStr(GetInsidePrivs(p住院医嘱发送), "执行情况登记") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnShowExec And blnSelect And (int执行状态 = 0 Or int执行状态 = 3) _
                    And vsExec.TextMatrix(vsExec.Row, vsExec.FixedCols) <> "" And vsExec.Row = vsExec.FixedRows
            End If
        Case ID_核对, ID_取消核对
            If InStr(GetInsidePrivs(p住院医嘱发送), "执行情况登记") = 0 Or Val(gstr医嘱核对) = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnShowExec And blnSelect And (int执行状态 = 0 Or int执行状态 = 3)
                If mblnShowExec And (int执行状态 = 0 Or int执行状态 = 3) Then
                    If vsExec.TextMatrix(vsExec.Row, COLExec("核对人")) = "" Then
                        If Control.ID = ID_核对 Then Control.Enabled = True
                        If Control.ID = ID_取消核对 Then Control.Enabled = False
                    Else
                        If Control.ID = ID_核对 Then Control.Enabled = False
                        If Control.ID = ID_取消核对 Then Control.Enabled = True
                    End If
                End If
            End If
    End Select
End Sub

Private Sub cbsSub_ControlSelected(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Control Is Nothing Then
        Select Case Control.ID
            Case ID_医嘱颜色示例
                vsfAdviceColor.Visible = True
                If vsfAdviceColor.Row <= 0 Then
                    With vsfAdviceColor
                        .ColWidth(0) = 3400
                        .Width = 3400
                        .Height = 3300
                        .TextMatrix(0, 0) = "新开"
                        .Cell(flexcpForeColor, 0, 0, 0, 0) = vbBlack
                        .RowHeight(0) = 300

                        .TextMatrix(1, 0) = "校对疑问"
                        .Cell(flexcpForeColor, 1, 0, 1, 0) = &H80&
                        .RowHeight(1) = 300

                        .TextMatrix(2, 0) = "已校对/已重整/已启用"
                        .Cell(flexcpForeColor, 2, 0, 2, 0) = &HC00000
                        .RowHeight(2) = 300
                        
                        .TextMatrix(3, 0) = "已停止/已确认停止/未用医嘱"
                        .Cell(flexcpForeColor, 3, 0, 3, 0) = &H808080
                        .RowHeight(3) = 300
                        
                        .TextMatrix(4, 0) = "已暂停"
                        .Cell(flexcpForeColor, 4, 0, 4, 0) = &H8000&
                        .RowHeight(4) = 300
                        
                        .TextMatrix(5, 0) = "已作废"
                        .Cell(flexcpForeColor, 5, 0, 5, 0) = &H808080
                        .Cell(flexcpFontStrikethru, 5, 0, 5, 0) = True
                        .RowHeight(5) = 300

                        .TextMatrix(6, 0) = "停止、暂停后时间未到"
                        .Cell(flexcpForeColor, 6, 0, 6, 0) = &HFF8080
                        .RowHeight(6) = 300

                        .TextMatrix(7, 0) = "启用后时间未到"
                        .Cell(flexcpForeColor, 7, 0, 7, 0) = &H4AAD00
                        .RowHeight(7) = 300

                        .TextMatrix(8, 0) = "术后医嘱校对后，转科医嘱发送后"
                        .Cell(flexcpForeColor, 8, 0, 8, 0) = vbRed
                        .RowHeight(8) = 300

                        .TextMatrix(9, 0) = "(仅医嘱内容列)毒麻精神特殊药品"
                        .Cell(flexcpFontBold, 9, 0, 9, 0) = True
                        .RowHeight(9) = 300
                        
                        .TextMatrix(10, 0) = "当天已发送的(长嘱可能发送到将来)"
                        .Cell(flexcpForeColor, 10, 0, 10, 0) = &HA08000
                        .RowHeight(7) = 300

                        .Top = vsAdvice.Top + 300
                        .Left = Me.Width - 3500
                        .Row = -1
                    End With
                End If
        End Select
    Else
        vsfAdviceColor.Visible = False
    End If
End Sub

Private Sub cbsSub_InitCommandsPopup(ByVal CommandBar As XtremeCommandBars.ICommandBar)
    Dim objControl As CommandBarControl
    Dim arrBaby As Variant, i As Long
    Dim strTmp As String
    If CommandBar.Parent Is Nothing Then Exit Sub
    
    Select Case CommandBar.Parent.ID
    Case ID_婴儿
        strTmp = IIF(mvarCond.过滤模式 = 3, "报告", "医嘱")
        arrBaby = Split(mstr婴儿, "<Split>")
        With CommandBar.Controls
            .DeleteAll
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100#, "所有" & strTmp)
            Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + 1, "病人" & strTmp): objControl.BeginGroup = True
            For i = 0 To UBound(arrBaby)
                Set objControl = .Add(xtpControlButton, ID_婴儿 * 100# + i + 2, "婴儿 " & i + 1 & IIF(arrBaby(i) <> "", "：" & arrBaby(i), ""))
                If i = 0 Then objControl.BeginGroup = True
            Next
        End With
    Case Else
        Call zlPopupCommandBars(CommandBar)
    End Select
End Sub

Private Sub DkpBlood_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If picBlood.Tag <> "可见" Then Exit Sub
    If Item.ID = 1 Then
        If InitObjBlood = True And Not Item.Tag = 1 Then
            If mobjFrmBlood Is Nothing Then
                Set mobjFrmBlood = gobjPublicBlood.zlGetBloodExec
            End If
            Item.Handle = mobjFrmBlood.hwnd
            Item.Tag = 1
        End If
    End If
End Sub

Private Sub fraExecUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAppend.Height + Y < 700 Or vsExec.Height - Y < 700 Then Exit Sub
        fraExecUD.Top = fraExecUD.Top + Y
        vsAppend.Height = vsAppend.Height + Y
        picExec.Top = picExec.Top + Y
        vsExec.Top = vsExec.Top + Y
        vsExec.Height = vsExec.Height - Y
        '输血执行
        picBlood.Top = picBlood.Top + Y
        picBlood.Height = picBlood.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub fraHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timHide.Enabled = True
End Sub

Private Sub mfrmCompoundMedicine_SetEditState(ByVal blnEditState As Boolean)
'功能：根据当前是否修改状态，设置是否可转移焦点
    RaiseEvent SetEditState(blnEditState)
    mblnEditState = blnEditState
    vsAdvice.Enabled = Not blnEditState
End Sub

Private Sub mfrmCompoundMedicine_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mfrmEdit_EditDiagnose(ParentForm As Object, ByVal 病人ID As Long, ByVal 主页ID As Long, ByVal 科室ID As Long, ByVal str类型 As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, 病人ID, 主页ID, 科室ID, str类型, Succeed)
End Sub

Private Sub mfrmEdit_FormUnload(Cancel As Integer)
    If mlng危急值ID <> 0 Then
        Call GetCriticalData
    End If
    mlng危急值ID = 0
    If Not Cancel Then
        If mfrmEdit.mblnOK Then Call LoadAdvice(True)
        Set mfrmEdit = Nothing
        
        If Me.Visible Then
            Call BringWindowToTop(Me.hwnd)
        End If
        
         '处理路径清单的刷新（医嘱新开界面可能删除，新增等）
        If mlng路径状态 = 1 And Not gobjPath Is Nothing Then
            If GetInsidePrivs(p临床路径应用) <> "" Then
                Call gobjPath.zlRefresh(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, mintPState, mblnMoved, True)
            End If
        End If
    End If
    RaiseEvent Activate
End Sub

Private Sub mfrmEac_FormUnload(Cancel As Integer)
    Set mfrmEac = Nothing
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
    Dim strSQL As String
    
    '申请单据打印之后的处理
    If mstrBillPrint <> "" Then
        If Split(mstrBillPrint, ",")(0) = ReportNum Then
            strSQL = "Zl_诊疗单据打印_Insert('" & Split(mstrBillPrint, ",")(1) & "'," & Val(Split(mstrBillPrint, ",")(2)) & ",1,'" & UserInfo.姓名 & "')"
        End If
    End If
    
    On Error GoTo errH
    If strSQL <> "" Then
        zlDatabase.ExecuteProcedure strSQL, Me.Name
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub picAppend_Resize()
    On Error Resume Next

    If picAppend.Tag = "不执行" Then Exit Sub
    
    vsAppend.Left = 0
    vsAppend.Top = 0
    vsAppend.Width = picAppend.Width
        
     '输血执行登记和医嘱执行登记始终只能显示一个，或都不显示
    If picBlood.Tag = "可见" Then
        vsAppend.Height = picAppend.Height - picBlood.Height - IIF(DkpBlood.Tag = "可见", fraExecUD.Height, 0)
        vsAppend.TopRow = vsAppend.Row
                
        fraExecUD.Left = 0
        fraExecUD.Width = picAppend.Width
        fraExecUD.Top = vsAppend.Top + vsAppend.Height
        
        '输血执行相关
        With picBlood
            .Left = 0
            If DkpBlood.Tag = "可见" Then
                .Top = fraExecUD.Top + fraExecUD.Height
            Else
                .Top = vsAppend.Top + vsAppend.Height
            End If
            .Width = picAppend.Width
        End With
    Else
        vsAppend.Height = picAppend.Height - IIF(picExec.Tag = "可见", picExec.Height, 0) - IIF(vsExec.Tag = "可见", fraExecUD.Height + vsExec.Height, 0)
    
        vsAppend.TopRow = vsAppend.Row
                
        fraExecUD.Left = 0
        fraExecUD.Width = picAppend.Width
        fraExecUD.Top = vsAppend.Top + vsAppend.Height
        
        picExec.Left = 0
        picExec.Width = picAppend.Width
        If vsExec.Tag = "可见" Then
            picExec.Top = fraExecUD.Top + fraExecUD.Height
        Else
            picExec.Top = vsAppend.Top + vsAppend.Height
        End If
        
        vsExec.Left = 0
        vsExec.Width = picAppend.Width
        vsExec.Top = picExec.Top + picExec.Height
    End If
End Sub

Private Sub cbsSub_Resize()
    Dim BarHideH As Long, PriceH As Long
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsSub.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If cbsSub.Count >= 3 Then
        If Not cbsSub(3).Visible Then BarHideH = fraHide.Height
    End If
    
    On Error Resume Next
    If fraMore.Visible Then
        fraMore.Tag = ""
        fraMore.Visible = False
    End If
    
    PriceH = IIF(tbcAppend.Visible, fraAdviceUD.Height + tbcAppend.Height, 0)
    
    fraHide.Left = lngLeft
    fraHide.Top = lngTop
    fraHide.Width = lngRight - lngLeft
    
    vsAdvice.Left = lngLeft
    vsAdvice.Top = lngTop + BarHideH
    vsAdvice.Width = lngRight - lngLeft
    vsAdvice.Height = lngBottom - lngTop - PriceH - BarHideH
    
    '列选择器
    With vsAdvice
        fraColSel.Left = .Left + (.ColWidth(COL_F标志) + .ColWidth(COL_F报告) - fraColSel.Width) / 2 + 30
        fraColSel.Top = .Top + (225 - fraColSel.Height) / 2 + 30
    End With
    
    fraAdviceUD.Left = lngLeft
    fraAdviceUD.Top = vsAdvice.Top + vsAdvice.Height
    fraAdviceUD.Width = vsAdvice.Width
    
    tbcAppend.Left = lngLeft
    tbcAppend.Top = fraAdviceUD.Top + fraAdviceUD.Height
    tbcAppend.Width = vsAdvice.Width
End Sub

Private Sub tbcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim blnTmp As Boolean
    If Not Me.Visible Then Exit Sub
    blnTmp = mbln报告
    mbln报告 = False
    Select Case Item.Tag
    Case "长嘱和临嘱"
        mvarCond.过滤模式 = 0
    Case "长嘱"
        mvarCond.过滤模式 = 1
    Case "临嘱"
        mvarCond.过滤模式 = 2
    Case "报告"
        mvarCond.过滤模式 = 3
        mbln报告 = True
    End Select
    
    If Item.Tag <> "" And mlng病人ID <> 0 Then
        If blnTmp <> mbln报告 Then
            Call AddToolBarInDoctor
            Call DefInSidePlugInBar(mrsPlugInBar)
            cbsSub.RecalcLayout
        End If
        Call RefreshData
    End If
End Sub

Private Sub timBRefresh_Timer()
    '供血库输血执行窗体填写完执行内容后，医嘱对应内容的刷新
    Dim intState As Integer
    timBRefresh.Enabled = False
    If Not mobjFrmBlood Is Nothing Then
        On Error Resume Next
        intState = mobjFrmBlood.AdviceExecState
        If err <> 0 Then
            err.Clear
        Else
            mobjFrmBlood.ExecFresh = True
            Select Case intState
                Case 1, 2 '记录执行或调整执行，删除执行
                    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '要更新执行状态'可能要更新执行状态
                Case 3, 4 '执行完成,取消完成
                    Call LoadAdvice
                Case 5, 6 '执行核对,取消核对
                    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '要更新执行状态'可能要更新执行状态
            End Select
            mobjFrmBlood.ExecFresh = False
            mobjFrmBlood.AdviceExecState = 0
        End If
    End If
End Sub

Private Sub timHide_Timer()
'功能：处理过滤工具栏的自动显示和隐藏
    Dim vPos As PointAPI, vRect As RECT
    Static sngBegin As Single
    
    If Not mblnHideFilter Then
        timHide.Enabled = False
        sngBegin = 0: Exit Sub
    End If
    
    If sngBegin = 0 Then sngBegin = Timer
    GetCursorPos vPos
    
    If fraHide.Visible Then
        ScreenToClient Me.hwnd, vPos
        If vPos.X * Screen.TwipsPerPixelX >= fraHide.Left And vPos.X * Screen.TwipsPerPixelX <= fraHide.Left + fraHide.Width _
            And vPos.Y * Screen.TwipsPerPixelY >= fraHide.Top And vPos.Y * Screen.TwipsPerPixelY <= picMain.Top + fraHide.Top + fraHide.Height Then
            fraHide.BackColor = cbsSub.GetSpecialColor(XPCOLOR_SEPARATOR)
            If Timer - sngBegin >= 0.35 Then
                fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
                fraHide.Visible = False: cbsSub(2).Visible = True And cbsSub(2).Controls.Count > 0: cbsSub(3).Visible = True
                cboTime.Visible = True
                sngBegin = 0: cbsSub.RecalcLayout
            End If
        Else
            fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
            sngBegin = 0: timHide.Enabled = False
        End If
    ElseIf cbsSub(3).Visible Then
        cbsSub(3).GetWindowRect vRect.Left, vRect.Top, vRect.Right, vRect.Bottom
        If Not (vPos.X >= vRect.Left / Screen.TwipsPerPixelX And vPos.X <= vRect.Right / Screen.TwipsPerPixelX _
            And vPos.Y >= vRect.Top / Screen.TwipsPerPixelY And vPos.Y <= vRect.Bottom / Screen.TwipsPerPixelY) Then
            If Timer - sngBegin >= 1 Then
                sngBegin = 0: timHide.Enabled = False
                fraHide.Visible = True: cbsSub(2).Visible = False: cbsSub(3).Visible = False
                cboTime.Visible = False
                cbsSub.RecalcLayout
            End If
        Else
            sngBegin = 0
        End If
    End If
End Sub

Private Sub cbsSub_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnUnrefresh As Boolean
    Dim bln报告 As Boolean
    
    If Control.ID <> 0 Then
        If cbsSub.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
        Case ID_在用医嘱, ID_所有医嘱
            mvarCond.医嘱显示 = IIF(mvarCond.医嘱显示 = 1, 0, 1)
        Case ID_全部
            mvarCond.报告 = 0
        Case ID_检查
            mvarCond.报告 = 1
        Case ID_检验
            mvarCond.报告 = 2
        Case ID_其他
            mvarCond.报告 = 3
        Case ID_未出报告
            If mvarCond.未出报告 Then
                If mvarCond.已出报告 Then
                    mvarCond.未出报告 = Not mvarCond.未出报告
                End If
            Else
                mvarCond.未出报告 = Not mvarCond.未出报告
            End If
        Case ID_已出报告
            If mvarCond.已出报告 Then
                If mvarCond.未出报告 Then
                    mvarCond.已出报告 = Not mvarCond.已出报告
                End If
            Else
                mvarCond.已出报告 = Not mvarCond.已出报告
            End If
        Case ID_婴儿 * 100# '所有医嘱
            If mvarCond.婴儿 = -1 Then Exit Sub
            mvarCond.婴儿 = -1
            mbyt婴儿 = 0
            Call zlDatabase.SetPara("病人婴儿过滤", mvarCond.婴儿, glngSys, p住院医嘱下达)
        Case ID_婴儿 * 100# + 1 To ID_婴儿 * 100# + 6 '病人、婴儿医嘱
            If mvarCond.婴儿 = Control.ID - ID_婴儿 * 100# - 1 Then Exit Sub
            mvarCond.婴儿 = Control.ID - ID_婴儿 * 100# - 1
            mbyt婴儿 = mvarCond.婴儿
            Call zlDatabase.SetPara("病人婴儿过滤", mvarCond.婴儿, glngSys, p住院医嘱下达)
        Case ID_重整
            mvarCond.重整 = Not mvarCond.重整
        Case ID_未到终止时间
            mvarCond.未到终止时间 = Not mvarCond.未到终止时间
        Case ID_未记帐
            mvarCond.未记帐 = Not mvarCond.未记帐
        Case ID_科内
            mvarCond.科内 = Not mvarCond.科内
        Case ID_是报告医嘱
            If mvarCond.是报告医嘱 Then
                If mvarCond.非报告医嘱 Then
                    mvarCond.是报告医嘱 = Not mvarCond.是报告医嘱
                End If
            Else
                mvarCond.是报告医嘱 = Not mvarCond.是报告医嘱
            End If
        Case ID_非报告医嘱
            If mvarCond.非报告医嘱 Then
                If mvarCond.是报告医嘱 Then
                    mvarCond.非报告医嘱 = Not mvarCond.非报告医嘱
                End If
            Else
                mvarCond.非报告医嘱 = Not mvarCond.非报告医嘱
            End If
        Case ID_简洁
            mvarCond.显示模式 = 0
        Case ID_详细
            mvarCond.显示模式 = 1
        Case Else
            Call zlExecuteCommandBars(Control)
            blnUnrefresh = True
    End Select
    
    bln报告 = InStr("," & ID_未出报告 & "," & "," & ID_已出报告 & "," & "," & ID_全部 & "," & ID_检查 & "," & ID_检验 & "," & ID_其他 & ",", "," & Control.ID & ",") > 0
    
    If Not blnUnrefresh Then cbsSub.RecalcLayout
    
    If Not bln报告 And blnUnrefresh = False Then
        Call RefreshData
    ElseIf bln报告 Then
        Call Refresh报告
    End If
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Between(Control.ID, conMenu_Edit_Untread * 100# + 1, conMenu_Edit_Untread * 100# + 99) Then Control.Enabled = mlng病人ID <> 0 And mblnEditState = False
    If Not Control.Enabled Then Exit Sub
    Select Case Control.ID
        Case ID_时间, ID_时间标签
            If mvarCond.过滤模式 <> 3 And mvarCond.医嘱显示 = 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
        Case ID_未到终止时间
            Control.Checked = mvarCond.未到终止时间
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            If mvarCond.医嘱显示 = 0 And (mvarCond.过滤模式 = 0 Or mvarCond.过滤模式 = 1) Then
                Control.Visible = mvarCond.过滤模式 <> 3
            Else
                Control.Visible = False
            End If
        Case ID_在用医嘱
            Control.Checked = mvarCond.医嘱显示 = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 <> 3
        Case ID_所有医嘱
            Control.Checked = mvarCond.医嘱显示 = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 <> 3
        Case ID_全部
            Control.Checked = mvarCond.报告 = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_检查
            Control.Checked = mvarCond.报告 = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_检验
            Control.Checked = mvarCond.报告 = 2
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_其他
            Control.Checked = mvarCond.报告 = 3
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_未出报告
            Control.Checked = mvarCond.未出报告
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_已出报告
            Control.Checked = mvarCond.已出报告
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.过滤模式 = 3
        Case ID_婴儿 '婴儿医嘱条件
            If mstr婴儿 <> "" Then
                Control.Visible = False
                If mvarCond.婴儿 = -1 Then
                    Control.Caption = IIF(mvarCond.过滤模式 = 3, "所有报告", "所有医嘱")
                ElseIf mvarCond.婴儿 = 0 Then
                    Control.Caption = IIF(mvarCond.过滤模式 = 3, "病人报告", "病人医嘱")
                Else
                    Control.Caption = "婴儿 " & mvarCond.婴儿
                End If
                Control.Visible = True
            Else
                If mvarCond.婴儿 <> -1 Or Control.Visible Then
                    mvarCond.婴儿 = -1
                    mbyt婴儿 = 0
                    Control.Visible = False
                    Call zlDatabase.SetPara("病人婴儿过滤", mvarCond.婴儿, glngSys, p住院医嘱下达)
                End If
            End If
        Case ID_婴儿 * 100# '所有医嘱
            Control.Checked = mvarCond.婴儿 = -1
        Case ID_婴儿 * 100# + 1 To ID_婴儿 * 100# + 6 '病人、婴儿医嘱
            Control.Checked = mvarCond.婴儿 = Control.ID - ID_婴儿 * 100# - 1
        Case ID_重整
            Control.Checked = mvarCond.重整
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
        Case ID_未记帐
            If mint场合 <> 1 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Checked = mvarCond.未记帐
                Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            End If
        Case ID_科内
            If mint场合 <> 2 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Checked = mvarCond.科内
                Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            End If
        Case ID_是报告医嘱
            Control.Checked = mvarCond.是报告医嘱
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.过滤模式 <> 3
        Case ID_非报告医嘱
            Control.Checked = mvarCond.非报告医嘱
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.过滤模式 <> 3
        Case ID_简洁
            Control.Checked = mvarCond.显示模式 = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 <> 3
        Case ID_详细
            Control.Checked = mvarCond.显示模式 = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.过滤模式 <> 3
        Case Else
            Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If vsColumn.Visible Then
        vsColumn.SetFocus '列选择器
    Else
        If Me.Visible And vsAdvice.Enabled Then vsAdvice.SetFocus
    End If
    RaiseEvent Activate
End Sub

Private Sub Form_Deactivate()
    vsColumn.Visible = False '列选择器
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ActiveHotKey(KeyCode, Shift)
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsColumn
            If .Visible Then
                .Visible = False
                If vsAdvice.Enabled Then vsAdvice.SetFocus
            Else
                For i = .FixedRows To .Rows - 1
                    If vsAdvice.ColHidden(.RowData(i)) Or vsAdvice.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                vsColumn.Height = vsColumn.RowHeightMin * vsColumn.Rows + 150
                If vsColumn.Top + vsColumn.Height > Me.ScaleHeight Then
                    vsColumn.Height = Me.ScaleHeight - vsColumn.Top
                    vsColumn.Width = 1750
                Else
                    vsColumn.Width = 1470
                End If
                
                .Left = fraColSel.Left
                .Top = fraColSel.Top + fraColSel.Height
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub mfrmBilling_Unload(Cancel As Integer)
    '刷新医嘱发送明细
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Set mfrmBilling = Nothing
End Sub

Private Function CheckWindow() As Boolean
'功能：检查医嘱编辑窗口是否已经打开
    If Not mfrmEdit Is Nothing Then
        '当前窗口打开了
        MsgBox "医嘱编辑窗口已经打开，请先完成当前操作后再执行。", vbInformation, gstrSysName
        '定位到当前的窗口
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        If mfrmEdit.Visible Then mfrmEdit.SetFocus
        Exit Function
    Else
        '其它窗口打开了
        If Not CheckAdviceWindow("住院医嘱编辑") Then Exit Function
    End If
 
    '检查会诊申请单窗体是否已经打开
    If Not mfrmEac Is Nothing Then
        '当前窗口打开了
        MsgBox "会诊申请窗口已经打开，请先完成当前操作后再执行。", vbInformation, gstrSysName
        '定位到当前的窗口
        If mfrmEac.WindowState = vbMinimized Then mfrmEac.WindowState = vbNormal
        If mfrmEac.Visible Then mfrmEac.SetFocus
        Exit Function
    End If
 
    CheckWindow = True
End Function

Private Sub FuncBillPrint(objControl As CommandBarControl)
'功能：打印诊疗单据
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strNO As String, lng记录性质 As Long
    Dim strParameter As String
    Dim lng相关ID As Long
    Dim strErr As String
    Dim blnDo As Boolean
    Dim strBillName As String '诊疗单据的名称  病历文件列表.名称
    
    If objControl.Parameter = "" Then Exit Sub
    strParameter = objControl.Parameter
    If InStr(objControl.Parameter, "|") > 0 Then strParameter = Split(objControl.Parameter, "|")(0): strNO = Split(objControl.Parameter, "|")(1)
    
    strBillName = objControl.Caption
    strBillName = Replace("<Tab>" & strBillName, "<Tab>打印:", "")
    If InStr(strBillName, "(&") > 0 Then
        strBillName = Mid(strBillName, 1, InStr(strBillName, "(&") - 1)
    End If
    
    '出院病人不允许打印
    If mintPState = ps出院 Then
        MsgBox "该病人已经出院,不能打印:" & strBillName & "。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsAdvice
        '打印次数提示
        On Error GoTo errH
        lng相关ID = Decode(Val(.TextMatrix(.Row, COL_相关ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_相关ID)))
        If .TextMatrix(.Row, COL_诊疗类别) = "E" And Val(.TextMatrix(.Row, COL_操作类型)) = 6 Then
            If Not gobjLIS Is Nothing Then    '打印检验申请单据
                 blnDo = gobjLIS.CheckAcceptance(CStr(lng相关ID), strErr)
                 If Not blnDo Then
                    MsgBox "该标本已经被检验科核收，不能打印:" & strBillName & "。", vbInformation, gstrSysName
                    Exit Sub
                 End If
            End If
        End If
        If mintBillPrint = 0 Then
            If strNO <> "" Then
                strSQL = "Select A.NO,A.记录性质 from 病人医嘱发送 A,病人医嘱记录 B Where a.医嘱ID=b.id And a.NO=[2] And (b.ID=[3] Or b.相关ID=[3])"
            Else
                strSQL = "Select NO,记录性质 from 病人医嘱发送 Where 医嘱ID=[1] order By 发送时间 Desc"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, COL_ID)), strNO, lng相关ID)
            If rsTmp.RecordCount > 0 Then
                strNO = rsTmp!NO & ""
                lng记录性质 = Val(rsTmp!记录性质 & "")
            End If
        Else
            strNO = vsAppend.TextMatrix(vsAppend.Row, COLSend("单据号"))
            lng记录性质 = Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("记录性质")))
        End If
        strSQL = "Select 打印人,打印时间 From 诊疗单据打印 Where NO=[1] And 记录性质=[2] And 打印内容=1 Order by 打印时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strNO, lng记录性质)
        If Not rsTmp.EOF Then
            If MsgBox("该[" & strBillName & "]已经打印了 " & rsTmp.RecordCount & " 次，最近一次由""" & _
                rsTmp!打印人 & """在""" & Format(rsTmp!打印时间, "yyyy-MM-dd HH:mm") & """打印。" & vbCrLf & vbCrLf & "要继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
         '输血医嘱打印申请单调用相关函数进行检查
        If InStr(1, ",ZL1_INSIDE_1254_17_1,ZL1_INSIDE_1254_17_2,", "," & strParameter & ",") <> 0 Then
            If BloodApplyPrintCheck(Val(.TextMatrix(.Row, COL_ID)), 2, IIF(strParameter = "ZL1_INSIDE_1254_17_1", 1, 2), 1) = False Then Exit Sub
        End If
        On Error GoTo 0
        
        '调用打印
        If mobjReport.ReportPrintSet(gcnOracle, glngSys, strParameter, mfrmParent) Then
            mstrBillPrint = strParameter & "," & strNO & "," & lng记录性质
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strParameter, mfrmParent, "NO=" & strNO, "性质=" & lng记录性质, "医嘱ID=" & lng相关ID, 2)
            mstrBillPrint = ""
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDataMoved() As Boolean
    If mblnMoved Then
        MsgBox "病人的本次住院数据已经转出到后备数据库，不允许操作。" & vbCrLf & _
            "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
        CheckDataMoved = True
    End If
End Function

Private Sub FuncAdviceAdd()
'功能：新增医嘱
    Dim datTurn As Date, int婴儿 As Integer

    On Error GoTo errH
    
    If Not CheckWindow Then Exit Sub
    If CheckAdviceAddModi(0, 0, datTurn) = False Then Exit Sub
    
    If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
        If CheckPatiTurnLimit(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, datTurn, mintPState) = False Then Exit Sub
    End If

    If Not FuncPathAdd() Then Exit Sub
    '-1表示病人和婴儿
    If mvarCond.婴儿 >= 0 Then int婴儿 = mvarCond.婴儿
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mint场合, mMainPrivs, mlng病人ID, mlng主页ID, mlng前提ID, _
            int婴儿, , mblnModalNew, mlng界面科室ID, , , mintPState, mlng病区ID, mlng科室ID, datTurn, mlng医护科室ID, mstr前提IDs, mclsMipModule, mlng危急值ID, mlng会诊医嘱ID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceConfirm(ByVal blnOnePati As Boolean, ByVal Control As XtremeCommandBars.ICommandBarControl)
'功能：确认停止医嘱
    Dim lng医嘱ID As Long
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lng医嘱ID = 0
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 2, mlng病人ID, mlng主页ID, mlng病区ID, lng医嘱ID, mint场合 = 1, , , , , mblnDirect And Not mblnBatch Or mblnInsideTools Or blnOnePati, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdviceAudit()
'功能：审核医嘱
    Dim datTurn As Date
        
    If Not CheckWindow Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    If Not mblnHaveAuditPriv Then
        MsgBox "你不具有审核医嘱的资格！", vbInformation, gstrSysName
        Exit Sub
    End If
    If CheckDataMoved Then Exit Sub
    
    '审核时不控制时限，因为新开和修改时可能没到时限，但审核时到了，会导致产生的医嘱无法后续处理
'    If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
'        If CheckPatiTurnLimit(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, datTurn, mintPState) = False Then Exit Sub
'    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mint场合, mMainPrivs, mlng病人ID, mlng主页ID, _
            mlng前提ID, , , , mlng界面科室ID, True, , mintPState, mlng病区ID, mlng科室ID, datTurn, mlng医护科室ID, mstr前提IDs, mclsMipModule)
End Sub

Private Sub FuncAdviceDel()
'删除：删除当前医嘱
'说明：在主界面删除,对检查组合,手术组合,中药配方,是整个删除,一并给药只删除当前药品
    Dim strSQL As String, lng医嘱ID As Long
    Dim blnGroup As Boolean, i As Long, blnBat As Boolean, blnTrans As Boolean
    Dim lngRow As Long, arrSQL As Variant, lng申请序号 As Long
    Dim strDelIDs As String, arrDelID() As String
    Dim strDelDrugIDs As String              '记录删除的药品医嘱,用于传入合理用药监测
    Dim lngBabyEdit As Long, int期效 As Integer
    Dim strMsg As String
    Dim lng组ID As Long
    Dim blnRIS预约 As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strWhere As String
    Dim bln输血 As Boolean, strErr As String
    Dim bln疑问医嘱 As Boolean
    
    With vsAdvice
        '检查是否可以删除
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub

        '检查病人是否正在审核
        If Not CheckPatiIsAduit Then Exit Sub
        lngBabyEdit = CheckBabyEdit(Val(.TextMatrix(.Row, COL_婴儿ID)))
        If lngBabyEdit = 1 Then
            MsgBox "当前病人不在本科室，不允许删除病人医嘱。", vbInformation, gstrSysName
            Exit Sub
        ElseIf lngBabyEdit = 2 Then
            MsgBox "当前病人的婴儿不在本科室，不允许删除婴儿医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '医技和医生下达的医嘱不能互删，医技站只能删本科室下达的医嘱
        If mint场合 = 2 Then
            If InStr("," & mstr前提IDs & ",", "," & .TextMatrix(.Row, COL_前提ID) & ",") = 0 Then
                MsgBox "该医嘱不为当前医技科室下达，不能删除该医嘱。", vbInformation, gstrSysName
                Exit Sub
            ElseIf Val(.TextMatrix(.Row, COL_前提ID)) = 0 Then
                MsgBox "该医嘱不是医技科室下达，不能删除该医嘱。", vbInformation, gstrSysName
                Exit Sub
            End If
        ElseIf Val(.TextMatrix(.Row, COL_前提ID)) <> 0 Then
            MsgBox "该医嘱为医技科室下达，不能删除该医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '转科病人
        If CheckOtherDeptPatiOpt = False Then Exit Sub

        If InStr(",1,2,", .TextMatrix(.Row, COL_医嘱状态)) = 0 Then
            MsgBox "当前选择的医嘱已经过校对，不能删除。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '已签名的医嘱不能删除
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            MsgBox "当前选择的医嘱已经签名，不能删除。请先取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If

        If mint场合 = 1 Then
            '护士对于已经过审核的医嘱，不允许修改删除。
            If .TextMatrix(.Row, COL_开嘱医生) Like "*/*" Then
                MsgBox "当前选择的医嘱已经过医生审核，不能删除。", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            '无执业资格的医生只能删除修改未审核的医嘱。
            If Not mblnHaveAuditPriv Then
                If HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_开嘱医生))) Then
                    MsgBox "你没有资格删除当前选择的医嘱，或者当前选择的医嘱已经过审核，不能删除。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        '已经执行登记的医嘱项目不能删除
        If mlng路径状态 = 1 Then
            If CheckPathAdviceIsExe(lng医嘱ID) Then
                MsgBox "该医嘱对应的项目已经执行。" & vbCrLf & "请取消执行登记后再进行删除操作。", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        
        '启用血库系统输血医嘱删除限制，进入血库审核阶段的新开医嘱不能删
        bln输血 = gbln血库系统 And .TextMatrix(.Row, COL_诊疗类别) = "K"
        If gbln血库系统 And .TextMatrix(.Row, COL_诊疗类别) = "K" And InStr("5,2", Val(.TextMatrix(.Row, COL_审核状态))) > 0 Then
            MsgBox "该输血医嘱已被血库接收" & IIF(Val(.TextMatrix(.Row, COL_审核状态)) = 5, "正在配血", "并且已完成配血") & "，不能删除，若需删除请与输血科联系。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        bln疑问医嘱 = Val(.TextMatrix(.Row, COL_医嘱状态)) = 2
        
        'PASS
        If InStr(",5,6,", "," & .TextMatrix(.Row, COL_诊疗类别) & ",") > 0 Then
            strDelDrugIDs = "【西药】" & lng医嘱ID & "|" & .TextMatrix(.Row, COL_相关ID)
        ElseIf .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "4" Then
            strDelDrugIDs = "【中药】" & .Cell(flexcpData, .Row, COL_相关ID) & "|" & .TextMatrix(.Row, COL_ID)
        End If
        
        arrSQL = Array()

        If InStr(",5,6,", .TextMatrix(.Row, COL_诊疗类别)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_相关ID)) = Val(.TextMatrix(.Row, COL_相关ID)) Then blnGroup = True
            End If
            If blnGroup Then
                lng组ID = Val(.TextMatrix(.Row, COL_相关ID))
                If MsgBox("医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """与其它药品一并给药,确实要删除吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        ElseIf .TextMatrix(.Row, COL_申请序号) <> "" Then
            If .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                If MsgBox("确实要取消输血申请""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("要将""" & .TextMatrix(.Row, col_医嘱内容) & """同时申请的其他项目一起取消吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnBat = True
                End If
            End If
        Else
            If MsgBox("确实要删除医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        If .TextMatrix(.Row, COL_诊疗类别) = "D" Then
            If HaveRIS And gbln启用影像信息系统预约 Then
                blnRIS预约 = True
            End If
        End If
        
        Call CreatePlugInOK(p住院医嘱下达, mint场合)
        If blnBat Then
            lng申请序号 = Val(.TextMatrix(.Row, COL_申请序号))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, COL_医嘱状态) = "1" And Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 Then
                    '调用删除前外挂接口
                    On Error Resume Next
                    If Not gobjPlugIn Is Nothing Then
                        If gobjPlugIn.AdviceDeletBefor(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(.TextMatrix(i, COL_ID)), mint场合) = False Then
                            If err.Number = 0 Then Exit Sub
                        End If
                        Call zlPlugInErrH(err, "AdviceDeletBefor")
                    End If
                    If Not CheckDelAdivceOfPathItem(Val(.TextMatrix(i, COL_ID))) Then Exit Sub

                    If err.Number <> 0 Then err.Clear
                    On Error GoTo 0
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & .TextMatrix(i, COL_ID) & ",1)"
                    strDelIDs = strDelIDs & "," & .TextMatrix(i, COL_ID)
                End If
            Next
        Else
            '调用删除前外挂接口
            On Error Resume Next
            If Not gobjPlugIn Is Nothing Then
                If gobjPlugIn.AdviceDeletBefor(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, lng医嘱ID, mint场合) = False Then
                    If err.Number = 0 Then Exit Sub
                End If
                Call zlPlugInErrH(err, "AdviceDeletBefor")
            End If
            If Not CheckDelAdivceOfPathItem(lng医嘱ID) Then Exit Sub    '必须生成的路径医嘱检查
            If err.Number <> 0 Then err.Clear
            On Error GoTo 0
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & lng医嘱ID & ",1)"
            strDelIDs = strDelIDs & "," & lng医嘱ID
        End If
        
        '医嘱打印判断
        strDelIDs = Mid(strDelIDs, 2)
        If blnGroup Then
            For i = .Row To .FixedRows - 1 Step -1
                If .TextMatrix(i, COL_期效) <> "" Then
                    int期效 = IIF(.TextMatrix(i, COL_期效) = "长嘱", 0, 1)
                    Exit For
                End If
            Next
        Else
            int期效 = IIF(.TextMatrix(.Row, COL_期效) = "长嘱", 0, 1)
        End If
        strSQL = Get病人打印记录DelSQL(4, mlng病人ID, mlng主页ID, Val(.TextMatrix(.Row, COL_婴儿ID)), int期效, , strDelIDs, Val(.TextMatrix(.Row, COL_婴儿ID)) <> 0, strMsg)
        
        If strMsg <> "" Then
            If MsgBox("您删除的医嘱或之后的医嘱已经打印，需清除重打。" & vbCrLf & strMsg & vbCrLf & "是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        If strSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    End With
    If blnRIS预约 Then
        Set rsTmp = GetDataRIS预约(strDelIDs)
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!预约id & "")) Then '医嘱删除
                MsgBox "当前启用了影像信息系统接口，但由于影像信息系统接口(HISSchedulingEx)取消息预约未调用成功，请与系统管理员联系！", vbInformation, gstrSysName
            End If
            rsTmp.MoveNext
        Next
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    If bln输血 = True Then
        If InitObjBlood(True) = True Then
            If gobjPublicBlood.AdviceOperation(p住院医嘱下达, lng医嘱ID, 2, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "血库公共部件调用失败，详细信息：" & strErr, vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "血库公共部件创建失败，请检查！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0

    '处理新开消息
    '如果有删除的医嘱要同步一下新开消息
    If gblnKSSStrict Or gbln手术分级管理 Or gbln输血分级管理 Or gbln血库系统 Then
        strWhere = strWhere & " And (Nvl(A.审核状态,0) Not in(1,3,7" & IIF(gbln血库系统 = True, "", ",4,5") & ") or a.医嘱期效=0 and a.审核状态=1 and a.紧急标志=1 and (instr(',5,6,',A.诊疗类别)>0 or A.诊疗类别='E' and B.操作类型='2'))"
    End If
    strSQL = "select 1 from 病人医嘱记录 a,诊疗项目目录 b where a.诊疗项目id=b.id(+) and A.医嘱状态=1 and a.病人id=[1] and a.主页id=[2]" & strWhere & _
            " And Exists ( Select M.姓名 From 人员表 M,执业类别 N" & _
            " Where M.姓名=Decode(A.审核标记,1,Substr(A.开嘱医生,1,Instr(A.开嘱医生,'/')-1),Substr(A.开嘱医生,Instr(A.开嘱医生,'/')+1))" & _
            " And M.执业类别=N.编码 And N.分类 IN('执业医师','执业助理医师')) And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3 and Rownum<2"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If rsTmp.EOF Then '无数据则将消息设置为已阅
         strSQL = "Zl_业务消息清单_Read(" & mlng病人ID & "," & mlng主页ID & ",'ZLHIS_CIS_001',3,'" & UserInfo.姓名 & "'," & mlng病区ID & ")"
         Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    '如果是疑问医嘱消息处理
    If bln疑问医嘱 Then
        strSQL = "select 1 from 病人医嘱记录 a where A.医嘱状态=2 and a.病人id=[1] and a.主页id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        If rsTmp.EOF Then '无数据则将消息设置为已阅
             strSQL = "Zl_业务消息清单_Read(" & mlng病人ID & "," & mlng主页ID & ",'ZLHIS_CIS_035',2,'" & UserInfo.姓名 & "'," & mlng科室ID & ")"
             Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If

    With vsAdvice
        '界面上直接删除
        .Redraw = False

        '删除一并给药第一行时的显示处理
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_相关ID)) = Val(.TextMatrix(.Row + 1, COL_相关ID)) Then
                If .TextMatrix(.Row, COL_开始时间) <> "" And .TextMatrix(.Row + 1, COL_开始时间) = "" Then
                    .TextMatrix(.Row + 1, COL_期效) = .TextMatrix(.Row, COL_期效)
                    .TextMatrix(.Row + 1, COL_开始时间) = .TextMatrix(.Row, COL_开始时间)
                    .TextMatrix(.Row + 1, COL_频率) = .TextMatrix(.Row, COL_频率)
                    .TextMatrix(.Row + 1, COL_用法) = .TextMatrix(.Row, COL_用法)
                End If
            End If
        End If

        lngRow = .Row
        If blnBat Then
            For i = .Rows - 1 To 1 Step -1
                If .TextMatrix(i, COL_医嘱状态) = "1" And Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 Then
                    .RemoveItem i
                End If
            Next
        Else
            .RemoveItem .Row
        End If

        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        If lngRow <= .Rows - 1 Then
            .Row = lngRow
        Else
            .Row = .Rows - 1
        End If
        If lng组ID <> 0 Then
            i = .FindRow(CStr(lng组ID), , COL_相关ID)
            If i <> -1 Then
                 .TextMatrix(i, COL_并) = ""
                Call SetTag一并给药(i)
            End If
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = True
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)    '颜色及附表更新
        
        
        '自动刷新医嘱提醒区域
        RaiseEvent RequestRefresh(True)

        '调用删除后外挂接口
        On Error Resume Next
        arrDelID = Split(strDelIDs, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                If Not gobjPlugIn Is Nothing Then
                    Call gobjPlugIn.AdviceDeleted(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(arrDelID(i)), mint场合)
                    Call zlPlugInErrH(err, "AdviceDeleted")
                End If
            End If
        Next
        If err.Number <> 0 Then err.Clear
        On Error GoTo errH
    End With
    If mlng路径状态 = 1 And Not gobjPath Is Nothing Then
        If GetInsidePrivs(p临床路径应用) <> "" Then
            Call gobjPath.zlRefresh(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, mintPState, mblnMoved, True)
        End If
    End If
    'PASS医嘱删除后自动调用审查功能
    If mblnPass And mint场合 = 0 Then
        Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 4, strDelDrugIDs)
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckOtherDeptPatiOpt() As Boolean
'功能：检查转科病人的当前医嘱是否允许操作
    
     '转科病人
    If mintPState = ps最近转出 Then
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_开嘱科室ID)) <> mlng科室ID Then
            MsgBox "不允许操作其他科室下达的病人医嘱。", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckOtherDeptPatiOpt = True
End Function

Private Sub FuncAdviceModi()
'功能：修改当前医嘱
    Dim lng医嘱ID As Long
    Dim datTurn As Date
    
    If Not CheckWindow Then Exit Sub
    With vsAdvice
        If CheckAdviceAddModi(1, lng医嘱ID, datTurn) = False Then Exit Sub
        Set mfrmEdit = frmInAdviceEdit
        Call frmInAdviceEdit.ShowMe(mfrmParent, mint场合, mMainPrivs, mlng病人ID, mlng主页ID, mlng前提ID, Val(.TextMatrix(.Row, COL_婴儿ID)), lng医嘱ID, , mlng界面科室ID, , , mintPState, _
                 mlng病区ID, mlng科室ID, datTurn, mlng医护科室ID, mstr前提IDs, mclsMipModule, , mlng会诊医嘱ID)
    End With
End Sub

Private Sub FuncAdviceSort()
'功能：调整医嘱顺序
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mint场合, mMainPrivs, mlng病人ID, mlng主页ID, mlng前提ID, _
            , , , mlng界面科室ID, False, 3, mintPState, mlng科室ID, mlng病区ID, , mlng医护科室ID)

End Sub

Private Sub FuncAdviceUnUse()
'功能：标记未用医嘱
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String, strSQL As String
    Dim lng医嘱ID As Long
    Dim i As Long, strTab As String
    Dim strErr As String
    Dim bln标记 As Boolean
    Dim blnFallback As Boolean
    Dim bln输血 As Boolean
    Dim str原因 As String
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If CheckOtherDeptPatiOpt = False Then Exit Sub
    
    On Error GoTo errH
    
    With vsAdvice
        lng医嘱ID = IIF(Val(.TextMatrix(.Row, COL_相关ID)) <> 0, Val(.TextMatrix(.Row, COL_相关ID)), Val(.TextMatrix(.Row, COL_ID)))
        bln输血 = gbln血库系统 And .TextMatrix(.Row, COL_诊疗类别) = "K"
        If Val(.TextMatrix(.Row, COL_执行标记)) = -1 Then
            strMsg = strMsg & "确实要将当前" & IIF(RowIn一并给药(.Row, 0, 0), "一并给药的", "") & "医嘱取消标记为未用吗？" & _
                IIF(.TextMatrix(.Row, COL_期效) = "临嘱", vbCrLf & vbCrLf & "你可能需要重新发送医嘱以产生费用和执行。", "")
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            bln标记 = False
        ElseIf Val(.TextMatrix(.Row, COL_执行标记)) <> -1 Then
            bln输血 = gbln血库系统 And .TextMatrix(.Row, COL_诊疗类别) = "K"
            If bln输血 Then
                '填写标记未用的原因
                Call zlCommFun.ShowMsgBox("输血医嘱标记未用", "请录入原因", "确定(&O),取消?(&C)", Me, , , , "2", , , "原因", 200, str原因)
                If str原因 = "" Then
                    MsgBox "操作失败：未录入输血原因。", vbInformation, gstrSysName
                    Exit Sub
                ElseIf Len(str原因) > 200 Then
                    MsgBox "操作失败：未录入输血原因超长，只能录入200个字符或者100个汉字。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            If .TextMatrix(.Row, COL_期效) = "临嘱" And Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 Then
                
                If GetAdviceFeeKind(lng医嘱ID) = 2 Then  '住院医生站的临嘱可发送到门诊
                    strTab = "住院费用记录"
                Else
                    strTab = "门诊费用记录"
                End If
            
                '检查特殊医嘱，发送表示已执行，不允许标记：3-转科;5-出院;6-转院,11-死亡
                If .TextMatrix(.Row, COL_诊疗类别) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(.Row, COL_操作类型))) > 0 Then
                    MsgBox "该特殊医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """已发送执行，不能标记为未用。", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                '检查已发送医嘱对应的执行状态和收费状态
                strSQL = "Select B.执行状态 as 医嘱执行,C.执行状态 as 费用执行,C.记录性质,C.记录状态" & _
                    " From 病人医嘱记录 A,病人医嘱发送 B," & strTab & " C" & _
                    " Where A.ID=[1] And A.ID=B.医嘱ID" & _
                    " And B.NO=C.NO(+) And B.记录性质=C.记录性质(+) And B.医嘱ID=C.医嘱序号(+)" & _
                    " And (B.执行状态 IN(1,3) Or C.执行状态 IN(1,2) Or (C.记录性质=1 And C.记录状态=1))" & _
                    " Union ALL " & _
                    " Select B.执行状态 as 医嘱执行,C.执行状态 as 费用执行,C.记录性质,C.记录状态" & _
                    " From 病人医嘱记录 A,病人医嘱发送 B," & strTab & " C" & _
                    " Where A.相关ID=[1] And A.ID=B.医嘱ID" & _
                    " And B.NO=C.NO(+) And B.记录性质=C.记录性质(+) And B.医嘱ID=C.医嘱序号(+)" & _
                    " And (B.执行状态 IN(1,3) Or C.执行状态 IN(1,2) Or (C.记录性质=1 And C.记录状态=1))"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng医嘱ID)
                If Not rsTmp.EOF Then
                    If InStr(",1,3,", NVL(rsTmp!医嘱执行, 0)) > 0 Then
                        strMsg = "医嘱已经执行或正在执行。"
                    ElseIf InStr(",1,2,", NVL(rsTmp!费用执行, 0)) > 0 Then
                        strMsg = "医嘱关联的费用已经执行或部分执行。"
                    ElseIf NVL(rsTmp!记录性质, 0) = 1 And NVL(rsTmp!记录状态, 0) = 1 Then
                        strMsg = "医嘱关联的费用已经在门诊收费。"
                    End If
                    MsgBox "当前" & IIF(RowIn一并给药(.Row, 0, 0), "一并给药的", "") & "医嘱不能标记为未用，" & strMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strMsg = "该医嘱已经发送，标记为未用将取消相关联的费用和发送状态。" & vbCrLf & vbCrLf
            End If
            strMsg = strMsg & "确实要将当前" & IIF(RowIn一并给药(.Row, 0, 0), "一并给药的", "") & "医嘱标记为未用吗？"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            bln标记 = True
        End If
    End With
    
    If bln输血 Then
        If InitObjBlood(True) Then
            If gobjPublicBlood.AdviceTermination(p住院医生站, lng医嘱ID, bln标记, False, strErr, blnFallback) = False Then
                MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    strSQL = "Zl_病人医嘱记录_未用(" & lng医嘱ID
    If bln标记 Then
        strSQL = strSQL & ",-1"
    Else
        strSQL = strSQL & ",0"
    End If
    If bln输血 Then
        strSQL = strSQL & ",1," & IIF(blnFallback, "null,", "1,") & IIF(str原因 = "", "null,", "'" & str原因 & "',") & "'" & UserInfo.姓名 & "')"
    Else
        strSQL = strSQL & ")"
    End If
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "6" Then
        '删除”检验申请组合”中的记录
        Call InitObjLis(p住院医生站)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(CStr(lng医嘱ID), strErr) = False Then
                MsgBox "删除检验申请失败：" & strErr, vbInformation, gstrSysName
            End If
        End If

    End If
    Call LoadAdvice
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdvicePause()
'功能：暂停医嘱
    Dim lng医嘱ID As Long, blnOnePati As Boolean
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Or mblnBatch Then
        lng医嘱ID = 0
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If mblnDirect And mblnBatch Then
        blnOnePati = False
    ElseIf mblnDirect And Not mblnBatch Or mblnInsideTools Then
        blnOnePati = True
    Else
        If mint场合 = 1 Then
            blnOnePati = Val(zlDatabase.GetPara("批量医嘱启停", glngSys, p住院医嘱发送)) = 0
        Else
            blnOnePati = True
        End If
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 5, mlng病人ID, mlng主页ID, mlng病区ID, lng医嘱ID, mint场合 = 1, , , , , blnOnePati, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdvicePrice()
'功能：调整病人的医嘱计价项目
    Dim lng医嘱ID As Long
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lng医嘱ID = 0
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 4, mlng病人ID, mlng主页ID, mlng病区ID, lng医嘱ID, mint场合 = 1, , , , , True, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub FuncAdviceReform()
'功能：重整医嘱
    Dim strSQL As String
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If MsgBox("要重整该病人的医嘱吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSQL = "ZL_病人医嘱记录_重整(" & mlng病人ID & "," & mlng主页ID & ",'" & UserInfo.姓名 & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    On Error GoTo 0

        '获取重整时间
    mdat重整 = GetRsRedoDate(mlng病人ID, mlng主页ID)

    If mblnDirect = False Then Call LoadAdvice
    
    If Val(zlDatabase.GetPara("自动进入医嘱打印", glngSys, p住院医嘱发送)) = 1 Then
        Call frmAdvicePrint.ShowMe(Me, mlng病人ID, mlng主页ID)
    Else
        MsgBox "病人医嘱重整完毕。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceNoPrint()
'功能：屏蔽打印
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strClear As String
    Dim lng医嘱ID As Long, lng页号 As Long
    Dim blnTran As Boolean, i As Long
    Dim datPrint As Date
    Dim int期效 As Integer
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If CheckOtherDeptPatiOpt = False Then Exit Sub
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        
        int期效 = IIF(.TextMatrix(.Row, COL_期效) = "长嘱", 0, 1)
        
        On Error GoTo errH
        
        If Val(.TextMatrix(.Row, COL_屏蔽打印)) = 0 Then
            strSQL = "Select Min(页号) as 页号,Min(打印时间) as 打印时间,Min(LPad(页号,4,'0')||LPad(行号,3,'0')) As 位置" & _
                    " From 病人医嘱打印 Where 医嘱id In (Select ID From 病人医嘱记录 Where ID = [1] Or 相关id = [1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncAdviceNoPrint", lng医嘱ID)
            If Not rsTmp.EOF Then
                lng页号 = NVL(rsTmp!页号, 0)
                datPrint = NVL(rsTmp!打印时间, CDate("1900-01-01"))
            End If
            
            If lng页号 > 0 Then
                '如果是在重整之前打印的，则不允许再屏蔽打印
                If datPrint < mdat重整 And datPrint <> CDate("1900-01-01") And int期效 = 0 Then
                    MsgBox "该医嘱在最近次重整之前已经打印过了，不能再屏蔽打印。", vbInformation, gstrSysName
                    Exit Sub
                Else
                    lng页号 = GetAdvicePrintPage(mlng病人ID, mlng主页ID, Val(.TextMatrix(.Row, COL_婴儿ID)), int期效, lng页号)
                    If datPrint <> CDate("1900-01-01") Then
                        If MsgBox("该医嘱已经打印过了，不能再屏蔽打印。" & vbCrLf & _
                            vbCrLf & "如果你确实想屏蔽打印，则将清除" & .TextMatrix(.Row, COL_期效) & "医嘱单在第 " & lng页号 & " 页开始的打印内容，这些页需要重新打印。" & _
                            vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        strClear = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & Val(.TextMatrix(.Row, COL_婴儿ID)) & "," & int期效 & "," & lng页号 & ")"
                    Else
                        strClear = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & Val(.TextMatrix(.Row, COL_婴儿ID)) & "," & int期效 & "," & lng页号 & "," & rsTmp!位置 & ")"
                    End If
                End If
            End If
        Else
            If int期效 = 0 And mdat重整 <> CDate("1900-01-01") Then
                '如果取消屏蔽后应在重整前打印，则不允许再取消屏蔽
                strSQL = "Select Count(*) From 病人医嘱状态 A,病人医嘱记录 B" & _
                    " Where A.操作时间+0<[2] And A.操作类型 Not In(1,2) And A.医嘱ID=B.ID And (B.ID=[1] Or B.相关ID=[1]) and b.医嘱期效=0"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncAdviceNoPrint", lng医嘱ID, mdat重整)
                If Not rsTmp.EOF Then
                    If NVL(rsTmp.Fields(0).value, 0) > 0 Then
                        MsgBox "该医嘱是在最近次重整之前屏蔽打印的，不能再取消屏蔽打印。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Else
                '取消屏蔽时：检查后面的医嘱是否已打印过了
                strSQL = "Select Min(页号) as 页号,Min(打印时间) as 打印时间,Min(LPad(页号,4,'0')||LPad(行号,3,'0')) As 位置 From 病人医嘱打印" & _
                    " Where 医嘱id In (" & _
                        " Select ID From 病人医嘱记录 A" & _
                        " Where a.病人id = [2] And a.主页id = [3] And Nvl(a.婴儿, 0) = [4] And a.医嘱期效 = [5] " & _
                        " And 序号 > (Select Max(序号) From 病人医嘱记录 Where ID = [1] Or 相关id = [1]))"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncAdviceNoPrint", lng医嘱ID, mlng病人ID, mlng主页ID, Val(.TextMatrix(.Row, COL_婴儿ID)), int期效)
                If Not rsTmp.EOF Then
                    lng页号 = NVL(rsTmp!页号, 0)
                    datPrint = NVL(rsTmp!打印时间, CDate("1900-01-01"))
                End If
                If lng页号 > 0 Then
                    lng页号 = GetAdvicePrintPage(mlng病人ID, mlng主页ID, Val(.TextMatrix(.Row, COL_婴儿ID)), int期效, lng页号)
                    If datPrint <> CDate("1900-01-01") Then
                        If MsgBox("该医嘱之后的医嘱已经打印过了，不能再取消屏蔽打印。" & vbCrLf & _
                            vbCrLf & "如果你确实想取消屏蔽打印，则将清除" & .TextMatrix(.Row, COL_期效) & "医嘱单在第 " & lng页号 & " 页开始的打印内容，这些页需要重新打印。" & _
                            vbCrLf & "要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        strClear = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & Val(.TextMatrix(.Row, COL_婴儿ID)) & "," & int期效 & "," & lng页号 & ")"
                    Else
                        strClear = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & Val(.TextMatrix(.Row, COL_婴儿ID)) & "," & int期效 & "," & lng页号 & "," & rsTmp!位置 & ")"
                    End If
                End If
            End If
        End If
    End With
    
    '执行
    gcnOracle.BeginTrans: blnTran = True
    If strClear <> "" Then
        zlDatabase.ExecuteProcedure strClear, Me.Name
    End If
    strSQL = "Zl_病人医嘱记录_屏蔽打印(" & lng医嘱ID & "," & IIF(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_屏蔽打印)) = 0, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans: blnTran = False
    
    '不读数据库刷新
    With vsAdvice
        .TextMatrix(.Row, COL_屏蔽打印) = IIF(Val(.TextMatrix(.Row, COL_屏蔽打印)) = 0, 1, 0)
        Call SetAdviceIcon(.Row)
        For i = .Row - 1 To .FixedRows Step -1
            If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) = _
                IIF(Val(.TextMatrix(.Row, COL_相关ID)) <> 0, Val(.TextMatrix(.Row, COL_相关ID)), Val(.TextMatrix(.Row, COL_ID))) Then
                .TextMatrix(i, COL_屏蔽打印) = .TextMatrix(.Row, COL_屏蔽打印)
                Call SetAdviceIcon(i)
            Else
                Exit For
            End If
        Next
        For i = .Row + 1 To .Rows - 1
            If IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID))) = _
                IIF(Val(.TextMatrix(.Row, COL_相关ID)) <> 0, Val(.TextMatrix(.Row, COL_相关ID)), Val(.TextMatrix(.Row, COL_ID))) Then
                .TextMatrix(i, COL_屏蔽打印) = .TextMatrix(.Row, COL_屏蔽打印)
                Call SetAdviceIcon(i)
            Else
                Exit For
            End If
        Next
    End With
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceResume()
'功能：启用医嘱
    Dim lng医嘱ID As Long, blnOnePati As Boolean
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Or mblnBatch Then
        lng医嘱ID = 0
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If mblnDirect And mblnBatch Then
        blnOnePati = False
    ElseIf mblnDirect And Not mblnBatch Or mblnInsideTools Then
        blnOnePati = True
    Else
        If mint场合 = 1 Then
            blnOnePati = Val(zlDatabase.GetPara("批量医嘱启停", glngSys, p住院医嘱发送)) = 0
        Else
            blnOnePati = True
        End If
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 6, mlng病人ID, mlng主页ID, mlng病区ID, lng医嘱ID, mint场合 = 1, , , , , blnOnePati, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdviceRevoke()
'功能：医嘱作废
    Dim lng医嘱ID As Long
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
        
    If mlng前提ID = 0 Or mblnDirect Then
        '用于医生站,护士站
        
        '转科病人
        If CheckOtherDeptPatiOpt = False Then Exit Sub
        
        If mblnDirect Then
            lng医嘱ID = 0
        Else
            lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        End If

        If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 0, mlng病人ID, mlng主页ID, mlng病区ID, lng医嘱ID, mint场合 = 1, , , , , True, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
            If mblnDirect = False Then
                Call LoadAdvice(True)
            End If
             'PASS医嘱作废后自动调用审查功能
            If mblnPass And mint场合 = 0 Then
                Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 3)
            End If
        End If
        
    Else
        '用于医技站
        If FuncAdviceRevokeTech Then
            Call LoadAdvice(True)
        End If
    End If
       
End Sub

Private Function FuncAdviceRevokeTech() As Boolean
'删除：当前医嘱作废(一组医嘱作废)
    Dim strSQL As String, lng医嘱ID As Long
    
    Dim str医嘱ID As String, intRule As Integer
    Dim lng签名id As Long, lng证书ID As Long
    Dim strSource As String, strSign As String
    Dim strTimeStamp As String, blnTran As Boolean, strTimeStampCode As String
    Dim datCur As Date, i As Integer
    Dim arrSQL As Variant
    Dim strMsg As String, rsTmp As ADODB.Recordset
    Dim strAdvice输血 As String
    Dim strErr As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以作废。", vbInformation, gstrSysName
            Exit Function
        End If
                
        If InStr(",1,2,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
            MsgBox "当前选择的住院医嘱尚未校对，请直接删除。", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
            MsgBox "当前选择的住院医嘱已经作废或停止。", vbInformation, gstrSysName
            Exit Function
        End If
        If .TextMatrix(.Row, COL_上次执行) <> "" Then
            MsgBox "当前选择的住院医嘱已经发送，不能再作废。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '92129:医嘱已被输血科接收则不能进行作废
        If .TextMatrix(.Row, COL_诊疗类别) = "K" And gbln血库系统 And InStr(1, ",2,5,6,", "," & Val(.TextMatrix(.Row, COL_审核状态)) & ",") <> 0 Then
            On Error GoTo errH
            strSQL = "Select Nvl(执行分类,0) as 执行分类 from 病人医嘱记录 A, 诊疗项目目录 B  where A.相关ID  = [1] and A.诊疗项目ID = B.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "查询诊疗项目的执行分类", lng医嘱ID)
            strSQL = ""
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp!执行分类) = 0 Then
                    MsgBox "本次作废的输血医嘱" & IIF(Val(.TextMatrix(.Row, COL_审核状态)) = 2, "已经完成配血", "处于正在配血阶段") & "，不能直接作废医嘱，若要作废请与输血科联系。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            On Error GoTo 0
            If gbln血库系统 Then strAdvice输血 = lng医嘱ID
        End If
        
        '电子签名检查和提示
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "作废已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能作废。", vbInformation, gstrSysName
                Else
                    MsgBox "作废已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能作废。", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            If gobjESign.CertificateStoped(UserInfo.姓名) = False Then strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
        End If
        
        If RowIn一并给药(.Row, 0, 0) Then
            If MsgBox("该组一并给药的医嘱将会一起作废，确实要作废吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("确实要作废医嘱""" & .TextMatrix(.Row, col_医嘱内容) & """吗？" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        arrSQL = Array()
        
        '医嘱打印检查
        strSQL = Get病人打印记录DelSQL(3, mlng病人ID, mlng主页ID, Val(.TextMatrix(.Row, COL_婴儿ID)), IIF(.TextMatrix(.Row, COL_期效) = "长嘱", 0, 1), lng医嘱ID, , Val(.TextMatrix(.Row, COL_婴儿ID)) <> 0, strMsg)
        If strMsg <> "" Then
            MsgBox "您作废的医嘱中包含已经打印的医嘱，请重打。", vbInformation, gstrSysName
        End If
        If strSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
        datCur = zlDatabase.Currentdate
        strSQL = "ZL_病人医嘱记录_作废(" & lng医嘱ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "',Null,To_Date('" & Format(datCur, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人危急值医嘱_Update(3,null," & lng医嘱ID & ")"    '删除危急值对应关系
        
        '作废时的电子签名
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '获取签名医嘱源文
            str医嘱ID = lng医嘱ID '组ID,返回为明细ID
            intRule = ReadAdviceSignSource(4, mlng病人ID, mlng主页ID, str医嘱ID, 0, mblnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "不能读取需要作废的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lng签名id = zlDatabase.GetNextID("医嘱签名记录")
                strSign = "zl_医嘱签名记录_Insert(" & lng签名id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & str医嘱ID & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSign
                
            Else
                Exit Function
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    If strAdvice输血 <> "" Then
        If InitObjBlood(True) Then
            If gobjPublicBlood.AdviceOperation(p住院医生站, lng医嘱ID, 4, mblnMoved, strErr) = False Then
                gcnOracle.RollbackTrans: blnTran = False
                Screen.MousePointer = 0
                MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    If Not (mclsMipModule Is Nothing) Then
        If mclsMipModule.IsConnect Then
            With vsAdvice
                Call ZLHIS_CIS_003(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mlng病区ID, "", mlng科室ID, "", , mstr床号, _
                    Val(.TextMatrix(.Row, COL_ID)), .TextMatrix(.Row, COL_期效), .TextMatrix(.Row, COL_诊疗类别), .TextMatrix(.Row, COL_操作类型), "", 0, UserInfo.姓名, datCur)
            End With
        End If
    End If
    '调用作废后外挂接口
    On Error Resume Next
    If CreatePlugInOK(p住院医嘱下达, mint场合) Then
        Call gobjPlugIn.AdviceRevoked(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, lng医嘱ID, mint场合)
        Call zlPlugInErrH(err, "AdviceRevoked")
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo 0
    FuncAdviceRevokeTech = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceChargeOff(Index As Integer)
'功能：费用冲销
'参数：Index=冲销子功能索引(0,1,2)
    Dim lng发送号 As String, lng医嘱ID As Long
    Dim strNO As String, bln划价 As Boolean
    Dim strCommon As String, intAtom As Integer
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)) <> 0 Then
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID))
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If lng医嘱ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
        
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    lng发送号 = Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("发送号")))
    If lng发送号 = 0 Then Exit Sub
    
    strNO = vsAppend.TextMatrix(vsAppend.Row, COLSend("单据号"))
    If strNO = "" Then Exit Sub
    
    If Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("记录性质"))) <> 2 Then Exit Sub
    
    '当前单据是否划价单,只是确定缺省的页面是否划价
    bln划价 = vsAppend.TextMatrix(vsAppend.Row, COLSend("计费状态")) = "记帐划价"
        
    '调用费用部件功能
    On Error Resume Next
    If gobjInExse Is Nothing Then
        Set gobjInExse = CreateObject("zl9InExse.clsInExse")
        If gobjInExse Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    
    Set mfrmBilling = Nothing
    
    If Index = 0 Then
        Set mfrmBilling = gobjInExse.CallByNurse( _
            mfrmParent, gcnOracle, gstrDBUser, glngSys, mlng病区ID, mlng科室ID, 0, lng医嘱ID, strNO, bln划价)
    ElseIf Index = 1 Then
        Set mfrmBilling = gobjInExse.CallByNurse( _
            mfrmParent, gcnOracle, gstrDBUser, glngSys, mlng病区ID, mlng科室ID, lng发送号, lng医嘱ID, "", bln划价)
    ElseIf Index = 2 Then
        Set mfrmBilling = gobjInExse.CallByNurse( _
            mfrmParent, gcnOracle, gstrDBUser, glngSys, mlng病区ID, mlng科室ID, lng发送号, 0, "", bln划价)
    End If
    Call GlobalDeleteAtom(intAtom)
    
    If mfrmBilling Is Nothing Then
        '刷新医嘱发送明细
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
    
    RaiseEvent StatusTextUpdate("")
End Sub

Private Function GetUploadAdvice(ByVal lng发送号 As Long, ByVal lng医嘱ID As Long, Optional ByVal blnBat As Boolean) As Recordset
'功能：获取回退医嘱的记账单据记录集
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    '取要回退的记帐NO
    '启用了参数：support允许冲销已结帐的记帐单据 的 允许冲销已结账的单据回退，但是这里只取单据号的，所以不包含记录性质=12的即可
    If blnBat Then
        strSQL = "Select Distinct A.NO From 病人医嘱发送 A,病人医嘱记录 B, 住院费用记录 C" & _
            " Where A.医嘱ID=B.ID And c.No = a.No And c.医嘱序号 = a.医嘱id And c.记录性质 = 2 And c.记录状态 = 1 And A.记录性质=2 And A.发送号=[1] "
    Else
        strSQL = "Select Distinct A.NO From 病人医嘱发送 A,病人医嘱记录 B, 住院费用记录 C" & _
            " Where A.医嘱ID=B.ID And c.No = a.No And c.医嘱序号 = a.医嘱id And c.记录性质 = 2 And c.记录状态 = 1 And A.记录性质=2 And A.发送号=[1] And (B.ID=[2] Or B.相关ID=[2])"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng发送号, lng医嘱ID)
    
    Set GetUploadAdvice = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceRoll()
'功能：医嘱回退
'参数：Index=回退内容在菜单上的索引
    Dim strSQL As String, lngFlag As Long
    Dim lng医嘱ID As Long, blnBat As Boolean
    Dim lng签名id As Long, strSign As String
    Dim vRoll As TYPE_AdviceRoll, str性质 As String, blnDo As Boolean, blnTran As Boolean
    Dim lngStarPage As Long, lng婴儿序号 As Long, strDelPrintTag As String
    Dim strSignIDs As String, arrSignSQL As Variant
    Dim i As Long, arrSQL As Variant
    Dim strAdvices As String, strErr As String
    Dim blnIsMany   As Boolean
    Dim lngBabyEdit As Long
    Dim strAdviceIDs As String
    Dim strAllmsg As String, strMsg As String
    Dim rsUpload As Recordset
    Dim rsTmp As ADODB.Recordset
    Dim bln部分回退 As Boolean
    Dim lng医嘱IDToRis As Long
    Dim strLISIDs As String
    Dim varSend As Variant
    Dim lngTmp As Long
    Dim strAdvices输血 As String
    Dim var输血 As Variant
    Dim colSQL As Collection
    
    '(组ID)取一组医嘱中相关ID为空的医嘱ID(给药途径,中药用法,主要手术,检查项目,及独立医嘱)
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID)) <> 0 Then
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_相关ID))
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If lng医嘱ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub

    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    lngBabyEdit = CheckBabyEdit(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_婴儿ID)))
    If lngBabyEdit = 1 Then
        MsgBox "当前病人不在本科室，不允许回退病人医嘱。", vbInformation, gstrSysName
        Exit Sub
    ElseIf lngBabyEdit = 2 Then
        MsgBox "当前病人的婴儿不在本科室，不允许回退婴儿医嘱。", vbInformation, gstrSysName
        Exit Sub
    End If
    '转科病人
    If CheckOtherDeptPatiOpt = False Then Exit Sub

    '回退信息
    If UBound(marrRollList) < 1 Then Exit Sub
    vRoll = marrRollList(1)

    '权限检查
    If mint场合 = 1 Then
        '护士回退
        If InStr(GetInsidePrivs(p住院医嘱发送), "回退他人操作") = 0 And vRoll.操作人员 <> UserInfo.姓名 Then
            MsgBox "你没有权限回退其他人对医嘱的操作：" & vbCrLf & vbCrLf & vRoll.操作内容 & vbTab, vbInformation, gstrSysName
            Exit Sub
        End If
        '护士不能回退医生操作
        str性质 = Get人员性质(vRoll.操作人员)
        If InStr("," & str性质 & ",", ",医生,") > 0 And InStr("," & str性质 & ",", ",护士,") = 0 Then
            MsgBox "你不能回退医生对医嘱的操作：" & vbCrLf & vbCrLf & vRoll.操作内容 & vbTab, vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        '医生回退：只能回退自已的操作,对电子签名同时也判断了是否回退本人的签名
        If vRoll.操作人员 <> UserInfo.姓名 Then
            MsgBox "你不能回退其他人对医嘱的操作：" & vbCrLf & vbCrLf & vRoll.操作内容 & vbTab, vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '回退发送时，已有费用转出不允许
    If vRoll.操作类型 = 0 Then
        If zlDatabase.DateMoved(vRoll.操作时间) Then
            If MovedBySend(lng医嘱ID, vRoll.发送号, 2) Then
                MsgBox "该医嘱的费用已经全部或部份转出到后备数据库，不允许操作。" & vbCrLf & _
                       "您可以与系统管理员联系，将相应数据抽选返回。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '不能回退发送操作
        If Not RollFirstEnabled Then Exit Sub
    End If

    '皮试结果不允许直接回退
    If vRoll.操作类型 = 10 Then
        MsgBox "皮试结果操作不允许直接回退。", vbInformation, gstrSysName
        Exit Sub
    End If

    '电子签名检查：单独回退时
    '------------------------------------------------------------------
    If vRoll.操作类型 = 0 Then bln部分回退 = InStr(GetInsidePrivs(p住院医嘱发送), ";部分回退医嘱;") > 0
    If mint场合 = 1 Then
        '护士回退
        If (vRoll.操作类型 = 4 Or vRoll.操作类型 = 8) And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 Then
            lng签名id = GetAdviceSign(lng医嘱ID, vRoll.操作类型, vRoll.操作人员, vRoll.操作时间)
            If lng签名id <> 0 Then
                MsgBox "该医嘱" & Decode(vRoll.操作类型, 4, "作废", 8, "停止") & "时已由医生签名，你不能执行回退。", vbInformation, gstrSysName
                Exit Sub
            End If

        ElseIf vRoll.操作类型 = 9 Then
            lng签名id = GetAdviceSign(lng医嘱ID, vRoll.操作类型, vRoll.操作人员, vRoll.操作时间)
        End If

        If MsgBox("确实要回退以下操作吗？" & vbCrLf & vbCrLf & _
                  vRoll.操作内容 & vbTab, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

        If vRoll.操作类型 = 5 Then
            '重整操作必须一起回退
            blnBat = True
        Else
            If InStr(";" & GetInsidePrivs(p住院医嘱发送) & ";", ";医嘱批量回退;") > 0 Then
                If RollBatchNurse(lng医嘱ID, vRoll.操作类型, vRoll.发送号, vRoll.操作时间, vRoll.操作类型 = 4 Or vRoll.操作类型 = 8 Or vRoll.操作类型 = 9, lng签名id, blnIsMany) Then
                    If MsgBox("还有其它医嘱和当前医嘱一起被同时" & _
                              Decode(vRoll.操作类型, 0, "发送", 4, "作废", 5, "重整", 6, "暂停", 7, "启用", 8, "停止", 9, "确认停止", 10, "填写皮试结果") & "，要同时回退这些医嘱吗？" & IIF(blnIsMany, vbCrLf & "选否将只回退同时签名的当前病人的其他医嘱。", ""), _
                              vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        blnBat = True
                    Else
                        If vRoll.操作类型 = 0 And Not bln部分回退 Then
                            MsgBox "您没有“部分回退医嘱”的权限，一起发送的医嘱只能一起回退。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    If blnIsMany Then
                        If MsgBox("存在其它医嘱和当前医嘱一起签名，必须一起回退，是否继续？", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Else
        '医生回退：根据医生是否作电子签名作检查和提示
        blnBat = RollBatchDoctor(lng医嘱ID, vRoll.操作类型, vRoll.发送号, vRoll.操作时间, lng签名id)    '其它一起操作的医嘱是否有签名
        If vRoll.操作类型 = 5 Then
            '重整操作必须一起回退
            blnBat = True
        End If

        strSQL = Decode(vRoll.操作类型, 0, "发送", 4, "作废", 5, "重整", 6, "暂停", 7, "启用", 8, "停止", 9, "确认停止", 10, "填写皮试结果", 13, "停嘱申请")
        If MsgBox("确实要回退以下操作吗？" & vbCrLf & vbCrLf & vRoll.操作内容 & vbTab & _
                  IIF(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 And lng签名id <> 0, _
                      vbCrLf & vbCrLf & "提示：该医嘱" & strSQL & "时已签名，将同时回退与它一起" & strSQL & "并签名的其它医嘱。", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

        '批量回退提示
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 And lng签名id <> 0 Then
            '当前及其它医嘱一起操作并签名,固定一起回退(blnBat=True)
        Else
            If blnBat And vRoll.操作类型 <> 5 Then    '重整操作必须一起回退
                If MsgBox("还有其它医嘱和当前医嘱一起被同时" & strSQL & "，要同时回退这些医嘱吗？", _
                          vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnBat = False
                    If vRoll.操作类型 = 0 And Not bln部分回退 Then
                        MsgBox "您没有“部分回退医嘱”的权限，一起发送的医嘱只能一起回退。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    '对医嘱费用的结帐情况进行检查
    If vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Then
        If Not CheckAdviceBalanceRoll(vRoll.发送号, lng医嘱ID, blnBat) Then Exit Sub
    End If

    '护士回退：对药品医嘱回退的数量进行留存检查
    If mint场合 = 1 And vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Then
        If Not (Not blnBat And InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别)) = 0) Then
            If Not blnBat And InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别)) > 0 Then
                strSQL = CheckAdviceDrugSurplus(vRoll.发送号, lng医嘱ID)
            Else
                strSQL = CheckAdviceDrugSurplus(vRoll.发送号)
            End If
            If strSQL <> "" Then
                If MsgBox(strSQL, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If

    If vRoll.操作类型 = 8 Then '临嘱不会直接回退自动停止
        If Not blnBat Then
            If RowIs配方行(vsAdvice.Row) Then
                lngFlag = 1    '中药配方始终保留执行终止时间
            End If
        End If
    End If

    '如涉及回退已签名的操作，先取消签名
    '-------------------------------------------------------
    If blnBat Then
        If lng签名id = 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 Then
            lng签名id = GetAdviceSign(lng医嘱ID, vRoll.操作类型, vRoll.操作人员, vRoll.操作时间)
        End If
        If vRoll.操作类型 = 9 Then
            strSignIDs = GetAdviceSigns(lng医嘱ID, vRoll.操作类型, vRoll.操作人员, vRoll.操作时间)
        End If
    Else
        lng签名id = 0
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_签名否)) = 1 Then
            lng签名id = GetAdviceSign(lng医嘱ID, vRoll.操作类型, vRoll.操作人员, vRoll.操作时间)
        End If
    End If
    '产生SQL
    arrSignSQL = Array()
    arrSQL = Array()
    If vRoll.操作类型 = 9 And blnBat Then
        If strSignIDs <> "" Then
            For i = 0 To UBound(Split(strSignIDs, ","))
                ReDim Preserve arrSignSQL(UBound(arrSignSQL) + 1)
                arrSignSQL(UBound(arrSignSQL)) = "zl_医嘱签名记录_Delete(" & Split(strSignIDs, ",")(i) & ")"
            Next
        End If
    Else
        If lng签名id <> 0 Then
            strSign = "zl_医嘱签名记录_Delete(" & lng签名id & ")"
        End If
    End If
    '检查能否回退签名
    If strSign <> "" Or UBound(arrSignSQL) > -1 Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "系统没有设置电子签名认证中心，回退操作不能继续。", vbInformation, gstrSysName
            Else
                MsgBox "电子签名部件未能正确安装，回退操作不能继续。", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '如果是回退发送且已计费,暂未处理销帐上传问题(1.可能是部分销帐,2.也可不管,预结时自动上传)
    If blnBat Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_病人医嘱记录_批量回退(" & lng医嘱ID & "," & vRoll.操作类型 & "," & _
                                 "To_Date('" & Format(vRoll.操作时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                 vRoll.发送号 & "," & lngFlag & ")"

        '如果是回退 确认停止操作 需要对消息进行处理
        If vRoll.操作类型 = 9 Then
            strAdviceIDs = GetRollAdviceIDs(lng医嘱ID, 2, vRoll.操作类型, vRoll.操作时间)
        End If
        
    Else
        If vRoll.操作类型 = 9 Then  '回退确认停止操作
            lngStarPage = CheckAdvicePrinted(lng医嘱ID, lng婴儿序号)
            If lngStarPage > 0 Then
                'zl_病人医嘱记录_回退，其中会检查不能回退重整之前的操作
                If MsgBox("该医嘱的停嘱时间已经打印，必须清除打印记录之后才能回退。" & vbNewLine & "是否清除第" & lngStarPage & _
                          "页及之后的医嘱打印记录?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    strDelPrintTag = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & lng婴儿序号 & ",0," & lngStarPage & ")"
                End If
            End If
        End If
        '如果是签名的，则回退的时候，同一签名ID一起回退
        If vRoll.操作类型 = 9 And lng签名id <> 0 Then
            strAdvices = GetAdvicesSameSign(lng签名id)
            If strAdvices = "" Then Exit Sub
            For i = 0 To UBound(Split(strAdvices, ","))
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_病人医嘱记录_回退(" & Split(strAdvices, ",")(i) & "," & lngFlag & ",Null,Null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            Next
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_病人医嘱记录_回退(" & lng医嘱ID & "," & lngFlag & ",Null,Null,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        End If
        
    End If
    
    '路径病人，在回退停止或作废医嘱时
    If mlng路径状态 = 1 And Not gobjPath Is Nothing And (vRoll.操作类型 = 4 Or vRoll.操作类型 = 8) Then
        If blnBat Then
            strAdviceIDs = GetRollAdviceIDs(lng医嘱ID, 2, vRoll.操作类型, vRoll.操作时间)
        Else
            strAdviceIDs = GetRollAdviceIDs(lng医嘱ID, 1)
        End If
    End If
    
    varSend = Array()
    If vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Then
        If blnBat Then
            Call GetAdvicesSameSend(vRoll.发送号, strLISIDs, strAdvices, "C")
        Else
            strAdvices = lng医嘱ID
            
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型)) = 6 Then
                strLISIDs = lng医嘱ID
            End If
        End If
        
        varSend = Split(strAdvices, ",")
        
        '获取回退的记帐单据
        Set rsUpload = GetUploadAdvice(vRoll.发送号, lng医嘱ID, blnBat)
        
        '---RIS项目判断
        If blnBat Then
            blnDo = HaveItemToRis(vRoll.发送号, lng医嘱IDToRis)
            If blnDo Then
                MsgBox "当前启用了影像信息系统接口，批量回退包含两个及以上的项目发送到影像信息系统中，不能进行批量回退，请按单个医嘱进行回退！", vbInformation, gstrSysName
                Exit Sub
            End If
            blnDo = False
        Else
            If InStr(",D,F,", vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别)) > 0 Or _
                vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" And InStr(",0,5,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型))) > 0 And vsAdvice.TextMatrix(vsAdvice.Row, COL_期效) = "临嘱" Then
                
                lng医嘱IDToRis = lng医嘱ID
                
            End If
        End If
        
        If HaveRIS(True) And lng医嘱IDToRis <> 0 Then
            On Error Resume Next
            If gobjRis.HISRollAdvice(lng医嘱IDToRis) <> 1 Then 'RIS医嘱回退操作
                MsgBox "当前启用了影像信息系统接口， 但由于影像信息系统接口(HISRollAdvice)未调用成功，请与系统管理员联系。", vbInformation, gstrSysName
            End If
            err.Clear: On Error GoTo 0
        End If
        '---
        
        '检查医嘱如果关联的药品费用已发药则自动产生药费的销帐申请
        If blnBat Then
            '批量操作
            strSQL = "select a.id from 病人医嘱记录 a,病人医嘱发送 b where a.id=b.医嘱id and b.发送号=[1] and a.诊疗类别='D' and a.相关id is null"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vRoll.发送号)
            For i = 1 To rsTmp.RecordCount
                lngTmp = MakeBillCharge(Val(rsTmp!ID & ""))
                If lngTmp = 1 Then
                    Exit Sub
                End If
                rsTmp.MoveNext
            Next
        Else
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "D" Then
                lngTmp = MakeBillCharge(lng医嘱ID)
                If lngTmp = 1 Then
                    Exit Sub
                End If
            End If
        End If
        
        '医嘱回退发送前调用外挂接口
        Call CreatePlugInOK(p住院医嘱下达)
        
        If Not gobjPlugIn Is Nothing Then '医嘱回退发送前外挂接口
            If UBound(varSend) > -1 Then
                On Error Resume Next
                For i = 0 To UBound(varSend)
                    If Val(varSend(i)) <> 0 Then
                        strMsg = ""
                        blnDo = gobjPlugIn.AdviceRollSendBefore(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(varSend(i)), mint场合, strMsg)
                        Call zlPlugInErrH(err, "AdviceRollSendBefore")
                        If 0 = err.Number Then '接口没有出错的情况下再判断接口的返回值
                            If Not blnDo Then
                                MsgBox strMsg, vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                blnDo = False
                If err.Number <> 0 Then err.Clear
                On Error GoTo 0
            End If
        End If
    End If
    
    If gbln血库系统 Then
        If vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Or vRoll.操作类型 = 4 Then '回退发送或者回退作废操作
            strAdvices = ""
            
            If blnBat Then
                If vRoll.操作类型 = 4 Then
                    strAdvices输血 = GetRollAdviceIDs(lng医嘱ID, 2, vRoll.操作类型, vRoll.操作时间, True)
                    strAdvices = strAdvices输血
                Else
                    Call GetAdvicesSameSend(vRoll.发送号, strAdvices输血, strAdvices, "K")
                End If
            Else
                If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "K" Then
                    strAdvices输血 = lng医嘱ID
                    strAdvices = lng医嘱ID
                End If
            End If
            If strAdvices输血 <> "" Then
                var输血 = Split(strAdvices输血, ",")
            End If

            If strAdvices <> "" Then
                If UBound(varSend) > -1 Then
                    varSend = Split(Join(varSend, ",") & "," & strAdvices, ",")
                Else
                    varSend = Split(strAdvices, ",")
                End If
            End If
        End If
    End If
    
    If Not (mclsMipModule Is Nothing) Then
        If mclsMipModule.IsConnect Then
            With vsAdvice
                i = .Row
                If .TextMatrix(i, COL_诊疗类别) = "Z" And .TextMatrix(i, COL_操作类型) = "5" And .TextMatrix(i, COL_状态) = "停止" Then
                    strSQL = "Select b.Id, b.病区id, b.科室id From 病人医嘱记录 A, 病人变动记录 B" & _
                        " Where a.病人id = b.病人id And a.主页id = b.主页id And b.开始原因 = 10 And a.开始执行时间 = b.开始时间 And a.Id = [1]"
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                 End If
            End With
        End If
    End If
    
    '临床路径关联检查
    If Not gobjPath Is Nothing And mlng路径状态 = 1 And (vRoll.操作类型 = 8 Or vRoll.操作类型 = 4) And strAdviceIDs <> "" Then
        Call gobjPath.zlAddOutPathItem(strAdviceIDs, mlng病人ID, mlng主页ID, vRoll.操作类型, colSQL)
        If GetInsidePrivs(p临床路径应用) <> "" Then
            Call gobjPath.zlRefresh(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, mintPState, mblnMoved, True)
        End If
    End If
    
    '执行SQL
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    If vRoll.操作类型 = 9 Then
        If UBound(arrSignSQL) > -1 Then
            For i = 0 To UBound(arrSignSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSignSQL(i)), Me.Name)
            Next
        End If
    Else
        If strSign <> "" Then
            Call zlDatabase.ExecuteProcedure(strSign, Me.Name)
        End If
    End If

    If strDelPrintTag <> "" Then
        Call zlDatabase.ExecuteProcedure(strDelPrintTag, Me.Name)
    End If

    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Name)
    Next
    
    '临床路径
    If Not colSQL Is Nothing Then
        For i = 1 To colSQL.Count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Name)
        Next
    End If
    
    If strAdvices输血 <> "" Then
        If InitObjBlood(True) Then
            For i = 0 To UBound(var输血)
                If gobjPublicBlood.AdviceOperation(p住院医生站, Val((var输血(i))), IIF(vRoll.操作类型 = 0, 6, 7), mblnMoved, strErr) = False Then
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    MsgBox "血库系统接口调用失败：" & strErr, vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End If
    End If
    
    '医保数据上传
    strAllmsg = ""
    If mint险类 <> 0 And vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Then
        If gclsInsure.GetCapability(support医嘱上传, mlng病人ID, mint险类) And Not gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, mint险类) Then
            Do While Not rsUpload.EOF
                strMsg = "" '因为现在一张NO内肯定为一个病人的,所以最后病人参数可以不传
                'strAdvance中传入“总单据数|当前单据数”以便医保接口处理
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 2, strMsg, , mint险类, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    '未提交前上传失败则回滚并中止发送
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName '每张提示
                    Else
                        MsgBox "费用上传失败，回退操作将被中止。", vbExclamation, gstrSysName
                    End If
                    Exit Sub
                Else
                    If strMsg <> "" Then strAllmsg = strAllmsg & rsUpload!NO & ":" & strMsg & vbCrLf
                End If
                rsUpload.MoveNext
            Loop
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    If strAllmsg <> "" Then
        Screen.MousePointer = 0
        MsgBox strAllmsg, vbInformation, gstrSysName
    End If
    If Not (mclsMipModule Is Nothing) Then
        If mclsMipModule.IsConnect Then
            With vsAdvice
            i = .Row
                Call ZLHIS_CIS_024(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mlng科室ID, "", _
                    .TextMatrix(i, COL_ID), .TextMatrix(i, COL_诊疗类别), .TextMatrix(i, COL_操作类型))
                '回退产生变动记录的医嘱时发送消，回退出院医嘱的发送。
                If .TextMatrix(i, COL_诊疗类别) = "Z" And .TextMatrix(i, COL_操作类型) = "5" And .TextMatrix(i, COL_状态) = "停止" Then
                    Call ZLHIS_PATIENT_006(mclsMipModule, mlng病人ID, mlng主页ID, mstr姓名, mstr性别, mstr住院号, rsTmp!ID, "预出院", NVL(rsTmp!病区ID, 0), NVL(rsTmp!科室ID, 0), NVL(rsTmp!病区ID, 0), NVL(rsTmp!科室ID, 0), "")
                End If
            End With
        End If
    End If
    '回退确认停止操作的消息处理
    If vRoll.操作类型 = 9 Then
        If blnBat Then
            If strAdviceIDs <> "" Then
                strSQL = "select a.病人ID,a.主页ID,nvl(a.紧急标志,0) as 紧急,max(id) as 医嘱ID from 病人医嘱记录 a " & _
                    " where a.id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
                    " group by a.病人ID,a.主页ID,a.紧急标志"
                    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAdviceIDs)
                strMsg = ""
                rsTmp.Filter = "紧急=1"
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        If InStr("," & strMsg & ",", "," & rsTmp!病人ID & "," & rsTmp!主页ID & ",") = 0 Then
                            strMsg = strMsg & "," & rsTmp!病人ID & "," & rsTmp!主页ID
                            Call SetCISMsg(rsTmp!病人ID, rsTmp!主页ID, rsTmp!医嘱ID, 1)
                        End If
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = "紧急<>1"
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        If InStr("," & strMsg & ",", "," & rsTmp!病人ID & "," & rsTmp!主页ID & ",") = 0 Then
                            strMsg = strMsg & "," & rsTmp!病人ID & "," & rsTmp!主页ID
                            Call SetCISMsg(rsTmp!病人ID, rsTmp!主页ID, rsTmp!医嘱ID, 0)
                        End If
                        rsTmp.MoveNext
                    Next
                End If
                strMsg = ""
            End If
        Else
            With vsAdvice
                strSQL = "select 1 From 业务消息清单 A Where a.病人id=[1] And a.就诊id=[2] And a.类型编码 ='ZLHIS_CIS_002' And a.优先程度=[3] And a.是否已阅=0 And Rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, IIF(Val(.TextMatrix(.Row, COL_标志)) = 1, 2, 1))
                If rsTmp.EOF Then
                    strSQL = "Zl_业务消息清单_Insert(" & mlng病人ID & "," & mlng主页ID & "," & mlng科室ID & "," & mlng病区ID & "," & IIF(mlng病人性质 = 1, 1, 2) & ",'有新停止医嘱。','0010','ZLHIS_CIS_002'," & _
                        Val(.TextMatrix(.Row, COL_ID)) & "," & IIF(Val(.TextMatrix(.Row, COL_标志)) = 1, 2, 1) & ",0,null," & mlng病区ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
                End If
            End With
        End If
    End If

    '调用LIS作废申请单
    If strLISIDs <> "" Then
        Call InitObjLis(p住院医生站)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(strLISIDs, strErr) = False Then
                MsgBox "删除检验申请失败：" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If

    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    '调用数据交换平台，向LIS,PACS取消申请单
    If Not gobjExchange Is Nothing And vRoll.操作类型 = 0 Then
        With vsAdvice
            If .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                blnDo = True
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "E" Then
                blnDo = RowIs检验行(.Row)
            End If
            If blnDo Then
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_诊疗类别) = "D", 2, 1), "病人ID::" & mlng病人ID & "||主页ID::" & mlng主页ID & "||医嘱ID::" & lng医嘱ID & "||操作类型::0||批量回退::" & IIF(blnBat, "1", "0"))
            End If
        End With
    End If
    
    If Not gobjPlugIn Is Nothing Then '医嘱回退发送后外挂接口
        If UBound(varSend) > -1 Then
            On Error Resume Next
            For i = 0 To UBound(varSend)
                If Val(varSend(i)) <> 0 Then
                    strMsg = ""
                    blnDo = gobjPlugIn.AdviceRollSend(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(varSend(i)), mint场合, strMsg)
                    Call zlPlugInErrH(err, "AdviceRollSend")
                    If 0 = err.Number Then '接口没有出错的情况下再判断接口的返回值
                        If Not blnDo Then
                            MsgBox strMsg, vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            Next
            blnDo = False
            If err.Number <> 0 Then err.Clear
            On Error GoTo 0
            On Error GoTo errH
        End If
    End If
    
    '刷新数据
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "Z" _
       And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型))) > 0 _
       And vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Then
        '回退出院医嘱刷新主界面
        RaiseEvent RequestRefresh(False)
    Else
        RaiseEvent StatusTextUpdate("")
        Call LoadAdvice
    End If
    
    '医保数据上传
    If mint险类 <> 0 And vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Then
        If gclsInsure.GetCapability(support医嘱上传, mlng病人ID, mint险类) And gclsInsure.GetCapability(support记帐完成后上传, mlng病人ID, mint险类) Then
            Do While Not rsUpload.EOF
                strMsg = ""
                Screen.MousePointer = 0
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 2, strMsg, , mint险类, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    '提交后上传失败,仅提示
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName
                    Else
                        MsgBox "记帐单""" & rsUpload!NO & """上传失败，HIS端数据已提交，按确定继续回退。", vbExclamation, gstrSysName
                    End If
                Else
                    If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                End If
                Screen.MousePointer = 11
                rsUpload.MoveNext
            Loop
        End If
    End If
    Screen.MousePointer = 0
    
    'PASS医嘱回退后自动调用审查功能
    If mblnPass And mint场合 = 0 Then
        Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 2)
    End If
    
    Exit Sub
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
        If blnTran Then
        blnTran = False
        'HIS事务回滚再调用RIS发送 lngRIS医嘱ID
        If HaveRIS(True) And lng医嘱IDToRis <> 0 Then
            strSQL = "Select a.病人id, a.主页id, a.挂号单, a.开嘱科室id, a.执行科室id, a.诊疗项目ID,a.诊疗类别 As 类别, b.发送号, a.Id As 医嘱id, Decode(a.挂号单, Null, 2, 1) As 病人来源" & vbNewLine & _
                "From 病人医嘱记录 A, 病人医嘱发送 B Where a.Id = b.医嘱id And a.Id =[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱IDToRis)
            If Not rsTmp.EOF Then
                Call gobjRis.HISSendAdvice(rsTmp, 2, Val(rsTmp!病人ID & ""), Val(rsTmp!主页ID & ""), "", Val(rsTmp!发送号 & ""))
            End If
        End If
    End If
End Sub

Private Function CheckAdvicePrinted(ByVal lng医嘱ID As Long, ByRef lng婴儿序号 As Long) As Long
'功能：检查当前医嘱的停止时间是否已打印
'返回：起始页号，lng婴儿序号=用于传递给清除打印记录的过程
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(Min(页号), 0) 页号, Nvl(Min(b.婴儿), 0) 婴儿序号" & vbNewLine & _
            "From 病人医嘱打印 A, 病人医嘱记录 B" & vbNewLine & _
            "Where a.打印标记 = 1 And a.医嘱id = b.Id And (b.Id = [1] Or b.相关id = [1])"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    CheckAdvicePrinted = Val(rsTmp!页号)
    lng婴儿序号 = Val(rsTmp!婴儿序号)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RollBatchDoctor(ByVal lng医嘱ID As Long, ByVal int类型 As Integer, ByVal lng发送号 As Long, ByVal dat时间 As Date, lng签名id As Long) As Boolean
'功能：检查指定医嘱当前操作是否与其它医嘱一起批量执行的,以判断是否可以批量回退
'参数：lng医嘱ID=相关ID为空的医嘱的ID(一组医嘱的ID)
'      int类型=医嘱操作类型
'      dat时间=医嘱操作的时间
'返回：是否有可以一起回退的其它医嘱
'      lng签名ID=这些要回退的医嘱是否已签名(作废,停止),如有则返回签名ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    lng签名id = 0
    If int类型 = 0 Then
        strSQL = "Select 医嘱ID From 病人医嘱发送 A Where 发送号=[2]" & _
            " And Not Exists(Select ID From 病人医嘱记录 B Where B.ID=A.医嘱ID And (ID=[1] Or 相关ID=[1]))"
    Else
        strSQL = "Select 操作类型,操作时间,操作人员 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[3] And 操作时间=[4]"
        strSQL = "Select 医嘱ID,Nvl(签名ID,0) as 签名ID From 病人医嘱状态 A Where (操作类型,操作时间,操作人员)=(" & strSQL & ")" & _
            " And Not Exists(Select ID From 病人医嘱记录 B Where B.ID=A.医嘱ID And (ID=[1] Or 相关ID=[1] Or (A.操作类型=8 And 医嘱期效=1)))"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng医嘱ID, lng发送号, int类型, dat时间)
    If Not rsTmp.EOF Then
        If int类型 = 0 Then
'            '不能通过批量回退已出院或预出院病人的医嘱发送
'            strSQL = "Select C.病人ID,C.主页ID From 病人医嘱发送 A,病人医嘱记录 B,病案主页 C" & _
'                " Where A.医嘱ID=B.ID And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
'                " And (C.出院日期 is Not NULL Or C.状态=3) And A.发送号=[1] And Rownum=1"
'            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng发送号)
'            If Not rsTmp.EOF Then Exit Function
        ElseIf int类型 <> 0 Then
            rsTmp.Filter = "签名ID<>0"
            If Not rsTmp.EOF Then lng签名id = rsTmp!签名ID
        End If
        RollBatchDoctor = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RollBatchNurse(ByVal lng医嘱ID As Long, ByVal int类型 As Integer, _
    Optional ByVal lng发送号 As Long, Optional ByVal dat时间 As Date, Optional ByVal blnCheckSign As Boolean, _
    Optional ByVal lng签名id As Long, Optional ByRef blnIsMany As Boolean) As Boolean
'功能：护士回退，检查指定医嘱当前操作是否与其它医嘱一起批量执行的,以判断是否可以批量回退
'参数：lng医嘱ID=相关ID为空的医嘱的ID(一组医嘱的ID)
'      int类型=0-发送,n-医嘱操作类型
'      lng发送号=回退发送时的发送号
'      dat时间=医嘱操作的时间
'      blnCheckSign=是否检查电子签名，只有全部未签名的才允许一起批量回退(确认停止除外)
'      lng签名ID=当前医嘱的签名ID,当回退确认停止签名且启用了签名功能时使用
'      blnIsMany=当回退确认停止签名且启用了签名功能时用于返回是否有相同签名ID的多条医嘱
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    blnIsMany = False
    
    If int类型 = 0 Then
        strSQL = "Select 医嘱ID From 病人医嘱发送 A Where 发送号=[2]" & _
            " And Not Exists(Select ID From 病人医嘱记录 B Where B.ID=A.医嘱ID And (ID=[1] Or 相关ID=[1]))"
    Else
        '排开临嘱发送(表现为停止)
        strSQL = "Select 操作类型,操作时间,操作人员 From 病人医嘱状态 Where 医嘱ID=[1] And 操作类型=[3] And 操作时间=[4]"
        strSQL = "Select a.医嘱ID,c.相关ID,Nvl(a.签名ID,0) as 签名ID From 病人医嘱状态 A,病人医嘱记录 C Where a.医嘱ID=c.ID And (a.操作类型,a.操作时间,a.操作人员)=(" & strSQL & ")" & _
            " And Not Exists(Select ID From 病人医嘱记录 B Where B.ID=A.医嘱ID And (b.ID=[1] Or b.相关ID=[1] Or (A.操作类型=8 And b.医嘱期效=1)))"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng医嘱ID, lng发送号, int类型, dat时间)
    If Not rsTmp.EOF Then
        If int类型 = 0 Then
            '不能通过批量回退已出院或预出院病人的医嘱发送
            strSQL = "Select C.病人ID,C.主页ID From 病人医嘱发送 A,病人医嘱记录 B,病案主页 C" & _
                " Where A.医嘱ID=B.ID And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
                " And (C.出院日期 is Not NULL Or C.状态=3) And A.发送号=[1] And Rownum=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng发送号)
            If Not rsTmp.EOF Then Exit Function
        ElseIf int类型 <> 0 And blnCheckSign Then
            If int类型 = 9 Then
                '如果同一个签名ID中有多条医嘱，则返回true，再提示用户操作的是多条医嘱
                If lng签名id <> 0 Then
                    rsTmp.Filter = "签名ID=" & lng签名id
                    blnIsMany = rsTmp.RecordCount > 0
                    
                    '是否存在多个签名ID，是则批量回退，不是则只回退当前病人（再根据blnIsMany决定是否回退多条医嘱）
                    rsTmp.Filter = "签名ID<>" & lng签名id
                    If rsTmp.EOF Then Exit Function
                End If
            Else
                '如有医生签名则不允许护士一起批量回退(作废或停止)
                rsTmp.Filter = "签名ID<>0"
                If Not rsTmp.EOF Then Exit Function
            End If
        End If
        RollBatchNurse = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceSend(ByVal blnOnePati As Boolean, ByVal Control As XtremeCommandBars.ICommandBarControl)
'功能：医嘱发送
    Dim blnRefresh As Boolean, blnOK As Boolean, lngTmp As Long
    
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    If mint场合 = 1 Then
        '护士站发送
        On Error Resume Next
        If mintPState = ps预出 Or mintPState = ps出院 Then
            Call MsgBox("该病人已" & IIF(mintPState = ps预出, "预", "") & "出院，不允许进行医嘱发送！", vbInformation, gstrSysName)
            Exit Sub
        End If
        If Control.ID = conMenu_Edit_SendInfusion Then
            If frmAdviceSendInfusion.ShowMe(mfrmParent, mlng病区ID, mlng病人ID, mlng主页ID, mMainPrivs, blnRefresh, mblnDirect And Not mblnBatch Or mblnInsideTools Or blnOnePati, mlng医护科室ID, mlng婴儿病区ID) Then
                blnOK = True
            End If
        ElseIf Control.ID = conMenu_Edit_Send Then
            If frmAdviceSendALL.ShowMe(mfrmParent, mlng病区ID, mlng病人ID, mlng主页ID, mMainPrivs, blnRefresh, mblnDirect And Not mblnBatch Or mblnInsideTools Or blnOnePati, mlng医护科室ID, mlng婴儿病区ID, mclsMipModule) Then
                blnOK = True
            End If
        End If
    Else
        If Control.ID = conMenu_Edit_SendCharge Then
            lngTmp = 1  '门诊收费
        Else
            If mlng病人性质 = 1 Then
                If InStr(GetInsidePrivs(p住院医嘱下达), ";发送门诊记帐;") > 0 Then
                    lngTmp = 2  '门诊记帐
                Else
                    lngTmp = 1  '门诊收费
                End If
            Else
                lngTmp = 0  '住院记帐
                If mintPState = ps预出 Or mintPState = ps出院 Then
                    Call MsgBox("该病人已" & IIF(mintPState = ps预出, "预", "") & "出院，不允许进行医嘱发送！", vbInformation, gstrSysName)
                    Exit Sub
                End If
            End If
        End If
        '医生(技)发送
        If frmInAdviceSend.ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mstr前提IDs, mlng病区ID, mlng科室ID, mlng界面科室ID, blnRefresh, lngTmp, mlng病人性质, mlng医护科室ID, mclsMipModule) Then
            blnOK = True
        End If
    End If
    
    If (blnOK Or blnRefresh) And mblnDirect = False Then
        If blnRefresh Then
            RaiseEvent RequestRefresh(False)
        Else
            RaiseEvent StatusTextUpdate("")
            Call LoadAdvice
        End If
    End If
End Sub

Private Sub FuncDrugSendQuery()
'功能：药疗收发查询
    Call frmDrugSendQuery.ShowQuery(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng病区ID, mlng病人ID, mblnDirect And Not mblnBatch Or mblnInsideTools)
End Sub

Private Sub FuncAdviceStop()
'功能：停止医嘱
    Dim blnRefresh As Boolean, lng医嘱ID As Long

    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub

    If mlng前提ID = 0 Or mblnDirect Then
        '用于医生站,护士站

        If mblnDirect Then
            lng医嘱ID = 0
        Else
            lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        End If

        If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 1, mlng病人ID, mlng主页ID, mlng病区ID, _
                                   lng医嘱ID, mint场合 = 1, blnRefresh, , , , True, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
            If mblnDirect = False Then
                If blnRefresh Then
                    '重新读取病人以刷新护理等级、病情
                    RaiseEvent RequestRefresh(False)
                Else
                    Call LoadAdvice(True)
                End If
            End If
             'PASS医嘱停止医嘱自动调用审查功能
            If mblnPass And mint场合 = 0 Then
                Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 1)
            End If
        End If
    Else
        '用于医技站
        If FuncAdviceStopTech Then
            Call LoadAdvice(True)
        End If
    End If
End Sub

Private Sub FuncAdviceStopAudit()
'功能：停嘱审核
    Dim blnRefresh As Boolean, lng医嘱ID As Long
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lng医嘱ID = 0
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If

    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 7, mlng病人ID, mlng主页ID, mlng病区ID, _
        lng医嘱ID, mint场合 = 1, blnRefresh, , , , True, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then
            If blnRefresh Then
                '重新读取病人以刷新护理等级、病情
                RaiseEvent RequestRefresh(False)
            Else
                Call LoadAdvice(True)
            End If
        End If
    End If
End Sub

Private Function FuncAdviceStopTech() As Boolean
'删除：当前医嘱停止(仅用于住院长嘱)
    Dim strSQL As String, lng医嘱ID As Long
    Dim strStopTime As String
    
    Dim str医嘱ID As String, intRule As Integer
    Dim lng签名id As Long, lng证书ID As Long
    Dim strSource As String, strSign As String, strTimeStamp As String, strTimeStampCode As String
    Dim colStopTime As New Collection, blnTran As Boolean
    
    With vsAdvice
        '检查是否可以作废
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lng医嘱ID = 0 Then
            MsgBox "该病人没有医嘱可以停止。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '检查
        If .TextMatrix(.Row, COL_期效) <> "长嘱" Then
            MsgBox "当前选择的医嘱不是住院长期医嘱。", vbInformation, gstrSysName
            Exit Function
        End If
        If .TextMatrix(.Row, COL_总量) <> "" Then
            MsgBox "中药配方在发送后会自动停止。", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",1,2,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
            MsgBox "当前选择的住院医嘱尚未校对，请直接删除。", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_医嘱状态))) > 0 Then
            MsgBox "当前选择的住院医嘱已经作废或停止。", vbInformation, gstrSysName
            Exit Function
        End If
        
        '电子签名检查和提示
        If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "停止已签名医嘱时需要再次签名，但系统没有设置签名认证中心，不能停止。", vbInformation, gstrSysName
                Else
                    MsgBox "停止已签名医嘱时需要再次签名，但电子签名部件未能正确安装，不能停止。", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            strSign = vbCrLf & vbCrLf & "提示：该医嘱已经签名，作废时你需要再次签名。"
        End If
        
        '停嘱时缺省的医嘱终止时间
        strStopTime = frmAdviceStopTime.ShowMe(Me, lng医嘱ID, mlng科室ID)
        If strStopTime = "" Then Exit Function
        
        strSQL = "ZL_病人医嘱记录_停止(" & lng医嘱ID & ",To_Date('" & strStopTime & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.姓名 & "')"
        
        '停止时的电子签名
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '获取签名医嘱源文
            str医嘱ID = lng医嘱ID '组ID,返回为明细ID
            colStopTime.Add Format(strStopTime, "yyyy-MM-dd HH:mm:00"), "_" & lng医嘱ID
            intRule = ReadAdviceSignSource(8, mlng病人ID, mlng主页ID, str医嘱ID, 0, mblnMoved, strSource, , colStopTime)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "不能读取需要停止的已签名医嘱源文内容。", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lng签名id = zlDatabase.GetNextID("医嘱签名记录")
                strSign = "zl_医嘱签名记录_Insert(" & lng签名id & ",8," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & str医嘱ID & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
            Else
                Exit Function
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    If strSign <> "" Then
        zlDatabase.ExecuteProcedure strSign, Me.Name
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    If Not (mclsMipModule Is Nothing) Then
        If mclsMipModule.IsConnect Then
            Call ZLHIS_CIS_002(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mlng病区ID, , mlng科室ID, "", , mstr床号, _
                lng医嘱ID, 0, 0, vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别), vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型), UserInfo.姓名, strTimeStamp, vsAdvice.TextMatrix(vsAdvice.Row, COL_标志))
        End If
    End If
    FuncAdviceStopTech = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceTest()
'功能：填写皮试结果
    Dim strSQL As String, str结果 As String
    Dim int结果 As Integer, strLabel As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTran As Boolean
    Dim dateInput As Date
    Dim strSelect As String, i As Long
    Dim strSelectInput As String
    Dim strTextInput As String
    
    Dim cnNew As ADODB.Connection
    Dim strOwner As String
    
    If mlng病人ID = 0 Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
        If Not CheckAdviceIsAduit Then Exit Sub
    
    If CheckOtherDeptPatiOpt = False Then Exit Sub
    
    If Not (vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型) = "1") Then
        MsgBox "当前医嘱内容不是过敏试验项目。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If InStr(",1,2,", vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) > 0 Then
        MsgBox "该过敏试验医嘱尚未通过校对，请先校对。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 3 Then
        MsgBox "该过敏试验医嘱尚未发送，不能填写结果。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) = 4 Then
        MsgBox "该过敏试验医嘱已经作废，不能填写结果。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) <> "" Then
        If MsgBox("该过敏试验医嘱已经填写了结果，要重新填写吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    On Error GoTo errH
    
    '先作身份验证
    If mbln皮试验证 Then
        Set cnNew = New ADODB.Connection
        If zlDatabase.UserIdentify(Me, "在填写皮试结果前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "皮试医嘱结果", cnNew) = "" Then Exit Sub
    End If
    
    strSQL = "Select Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    '阳性
    For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(0), ","))
        strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(0), ",")(i) & "|0"
    Next
    '阴性
    For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(1), ","))
        strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(1), ",")(i) & "|0|2"
    Next
    strSelect = Mid(strSelect, 2)
    
    str结果 = zlCommFun.ShowMsgBox("皮试结果", vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容) & "：^^请根据过敏试验结果选择相应的按钮操作。", _
            "确定(&O),?取消(&C)", Me, vbQuestion, "皮试时间", dateInput, "yyyy-MM-dd HH:mm", "皮试结果(&P):" & strSelect, strSelectInput, _
            "过敏反应(&F)", 50, strTextInput, , True)
    If str结果 = "" Then Exit Sub
    If strSelectInput = "" Then Exit Sub
    
    
    
    If Format(IIF(mvarCond.显示模式 = 0, vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_开始时间), vsAdvice.TextMatrix(vsAdvice.Row, COL_开始时间)), "yyyy-MM-dd HH:mm") > dateInput Then
        MsgBox "皮试时间不能在医嘱生效时间以前，请重新录入。", vbInformation, gstrSysName
        Exit Sub
    End If
    If mbln护士签名 Then
        If Not (Check电子签名) Then Exit Sub
    End If
    Call GetTestLabel(rsTmp!标本部位, strSelectInput, strLabel, int结果)
    strSQL = "ZL_病人医嘱记录_皮试(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & strLabel & "'," & int结果 & _
    ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
    
    
    If mbln皮试验证 And Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)
        
        On Error GoTo errNew
        
        Call SQLTest(App.ProductName, Me.Name, strSQL)
        cnNew.Execute strOwner & "." & strSQL, , adCmdStoredProc
        Call SQLTest
        
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        zlDatabase.ExecuteProcedure strSQL, Me.Name
    End If
    
    vsAdvice.TextMatrix(vsAdvice.Row, COL_皮试) = strLabel
    If mvarCond.显示模式 = 0 Then
        '如果是简洁模式，则设置药品后的皮试结果。
        If InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(+)") > 0 Or InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(-)") > 0 Then
            vsAdvice.TextMatrix(vsAdvice.Row, col_内容) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(+)", strLabel)
            vsAdvice.TextMatrix(vsAdvice.Row, col_内容) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_内容), "(-)", strLabel)
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, col_内容) = vsAdvice.TextMatrix(vsAdvice.Row, col_内容) & strLabel
        End If
    End If
    If int结果 = 1 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_皮试) = vbRed
    ElseIf int结果 = 0 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_皮试) = vbBlue
    End If
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
errNew:
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Sub FuncAdviceVerify()
'功能：医嘱校对
    Dim blnRefresh As Boolean, lng医嘱ID As Long, blnOnePati As Boolean
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lng医嘱ID = 0
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If mblnDirect And mblnBatch Then
        blnOnePati = False
    ElseIf mblnDirect And Not mblnBatch Or mblnInsideTools Then
        blnOnePati = True
    Else
        If mint场合 = 1 Then
            blnOnePati = Val(zlDatabase.GetPara("批量医嘱校对", glngSys, p住院医嘱发送)) = 0
        Else
            blnOnePati = True
        End If
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 3, mlng病人ID, mlng主页ID, mlng病区ID, _
        lng医嘱ID, mint场合 = 1, blnRefresh, , , , blnOnePati, , , , , mlng医护科室ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
 
        If mblnInsideTools Then
            If blnRefresh Then Call LoadAdvice
        ElseIf mblnDirect = False Then
            If blnRefresh Then
                '重新读取病人以刷新护理等级、病情
                RaiseEvent RequestRefresh(False)
            Else
                Call LoadAdvice
            End If
        End If
    End If
    
    
End Sub

Private Sub FuncAdviceSign()
'功能：对医嘱进行电子签名
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lng签名id As Long, lng证书ID As Long
    Dim intRule As Integer, strTimeStamp As String, strTimeStampCode As String
    Dim ColIDs As Collection, ColSource As Collection
    
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.姓名) Then
        MsgBox "您的签名证书已被停用，请联系信息科。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '获取签名医嘱源文
    intRule = ReadAdviceSignSource(1, mlng病人ID, mlng主页ID, strIDs, 0, mblnMoved, strSource, mstr前提IDs, , ColIDs, ColSource)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "该病人目前没有可以签名的医嘱。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    For i = 1 To ColIDs.Count
        strSign = gobjESign.Signature(ColSource(i), gstrDBUser, lng证书ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lng签名id = zlDatabase.GetNextID("医嘱签名记录")
            strSQL = "zl_医嘱签名记录_Insert(" & lng签名id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng证书ID & ",'" & ColIDs(i) & "'," & strTimeStamp & ",'" & UserInfo.姓名 & "','" & strTimeStampCode & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
        End If
    Next
    If strSign <> "" Then
        Call LoadAdvice '刷新界面
        MsgBox "已完成电子签名。", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'功能：校验医嘱的电子签名(可对已转移的数据)
    Dim strSource As String
    
    If mlng病人ID = 0 Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.姓名) Then
        MsgBox "您的签名证书已被停用，请联系信息科。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "签名" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '获取签名医嘱源文
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '验证签名
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub


Private Function Check电子签名() As Boolean
    '判断是否启用数字签名
    Check电子签名 = True
    If gintCA > 0 And CheckSign(2, mlng医护科室ID, , , , False, gobjESign) Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If gobjESign Is Nothing Then
            MsgBox "电子签名部件未能正确安装，签名操作不能继续。", vbInformation, gstrSysName
            Check电子签名 = False
            Exit Function
        Else
            If Not gobjESign.CheckCertificate(UserInfo.用户名) Then
                Check电子签名 = False
                Exit Function
            End If
        End If
    End If
End Function

Private Sub FuncAdviceSignErase()
'功能：取消医嘱的电子签名
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If mlng病人ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.姓名) Then
        MsgBox "您的签名证书已被停用，请联系信息科。", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "签名" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "当前选择的医嘱没有签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '作废和停止医嘱的签名不能取消
        If InStr(",4,8,", .Cell(flexcpData, .Row, 0)) > 0 Then
            MsgBox "不能直接取消作废或停止医嘱的签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        '新开签名必须是在新开或校对疑问状态
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If InStr(",1,2,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态))) = 0 Then
                MsgBox "由于医嘱已经经过校对，该签名不能取消。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '只能取消自已签的名
        If .TextMatrix(.Row, 2) <> UserInfo.姓名 Then
            MsgBox "该签名人不是你本人，不能取消签名。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("确实要取消这次签名吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_医嘱签名记录_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice '刷新界面
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncToolScheme()
'功能：调用成套方案维护
    On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "诊疗基础部件没有正确安装，功能无法执行。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallClinicScheme(mfrmParent, gcnOracle, glngSys, gstrDBUser, IIF(mint场合 = 2, 3, IIF(mlng病人性质 = 1, 1, 2)))
End Sub

Private Sub FuncEPRReport(ByVal lngMenu As Long)
'功能：查阅、打印、预览报告
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strBill As String, strTmp As String
    Dim strNO As String, int性质 As Long, i As Long
    Dim lng医嘱ID As Long, lngReportID As Long, blnPrint As Boolean, bln打印 As Boolean
    Dim bln检验行 As Boolean, bln配方行 As Boolean, arrRPTPar(19) As String, strFlagString As String
    Dim strSQLEPR As String, rsTmpEPR As ADODB.Recordset
    Dim str检查报告ID As String
    Dim lngViewMode As Long ' 1-病历格式，6-报表格式
    Dim objRichEPR As New zlRichEPR.cRichEPR
    Dim blnLis接口 As Boolean
    
        If mblnMoved Then
        MsgBox "当前病人报告数据已转出，请统一到电子病案查阅模块中进行查看。", vbInformation, gstrSysName
        Exit Sub
    End If
    '调用数据交换平台，向LIS,PACS查阅报告
    If lngMenu = conMenu_Edit_Compend * 10# + 1 Or lngMenu = conMenu_Edit_Compend * 10# + 6 Or lngMenu = conMenu_Edit_Compend Then
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Then
            lngViewMode = 1
        ElseIf lngMenu = conMenu_Edit_Compend * 10# + 6 Then
            lngViewMode = 6
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 1 Then
                lngViewMode = 1
            Else
                lngViewMode = 6
            End If
        End If
        
        If gobjExchange Is Nothing Then
            On Error Resume Next
            Set gobjExchange = CreateObject("zlExchange.clsExchange")
            If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
            err.Clear: On Error GoTo 0
        End If
        If Not gobjExchange Is Nothing Then
            With vsAdvice
                '检验行存的是采集方法（诊疗类别为E），所以只判断检查行
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_诊疗类别) = "D", 4, 3), "医嘱ID::" & .TextMatrix(.Row, COL_ID) & "||操作员姓名::" & UserInfo.姓名 & "||操作员缺省部门::" & UserInfo.部门名)
            End With
            Exit Sub
        End If
    End If
    
    lngReportID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告ID))
    lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    str检查报告ID = vsAdvice.TextMatrix(vsAdvice.Row, COL_检查报告ID)
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS报告ID)) <> 0 Then
        Call FuncLisRptFileView(mfrmParent, lng医嘱ID)   '三方的LIS文件报告
        If lngReportID = 0 And str检查报告ID = "" Then Exit Sub
    End If
    
    '先判断是否可以继续操作
    Select Case CheckEPRReport(lng医嘱ID, lngReportID, , , mblnMoved)
    Case 0
        MsgBox "该医嘱的报告没有书写！", vbInformation, gstrSysName
        Exit Sub
    Case 2
        strTmp = ""
        '紧急医嘱或者标记绿色通的项目可以查看未完成的报告
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_标志)) = 1 Then
            strTmp = "允许查看未完成报告"
        Else
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "D" Then
                strSQL = "select 1 from 影像检查记录 a where a.绿色通道=1 and a.医嘱id=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                If Not rsTmp.EOF Then
                    strTmp = "允许查看未完成报告"
                End If
            End If
        End If
        If InStr(GetInsidePrivs(p住院医嘱下达), "查阅未完成报告") > 0 Or strTmp <> "" Then
            MsgBox "注意：该医嘱的报告还没有正式签名！", vbInformation, gstrSysName
        Else
            MsgBox "该医嘱的报告还没有完成(没有正式签名或完成执行)，你没有权限操作！", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS报告ID)) <> 0 Then
        If HaveRIS Then 'RIS报告兼容
            i = gobjRis.ShowViewReport(mfrmParent.hwnd, lng医嘱ID, InStr(GetInsidePrivs(p住院医嘱下达), ";报告打印;") > 0)
            If i = 0 Then Exit Sub
        End If
    End If
    
    '执行操作
    '新版PACS报告，直接强制使用新版PACS报告编辑器
    If str检查报告ID <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(lng医嘱ID, , mblnAutoRead, mfrmParent)
    Else
        bln打印 = InStr(GetInsidePrivs(p住院医嘱下达), ";报告打印;") > 0 And (mintPState = ps在院 Or mintPState = ps待诊)
        
        '检验项目应该调用LIS接口
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E" Then
            Call InitObjLis(p住院医生站)
            If Not gobjLIS Is Nothing Then
                blnLis接口 = True
            End If
        End If
        
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Or (lngMenu = conMenu_Edit_Compend And lngViewMode = 1) Then
            '查阅报告
            If blnLis接口 Then
                strTmp = ""
                Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lng医嘱ID, 0, strTmp)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                If mfrmParent.Name <> "frmPatiFeeQuery" Then
                    RaiseEvent ViewEPRReport(lngReportID, bln打印)
                Else
                    objRichEPR.InitRichEPR gcnOracle, Me, glngSys, False
                    Call objRichEPR.ViewDocument(Me, lngReportID, bln打印)
                End If
            End If
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_报告项)) = 1 And lngMenu <> conMenu_Edit_Compend * 10# + 6 And Not (lngMenu = conMenu_Edit_Compend And lngViewMode = 6) Then
                '按编辑格式打印、预览报告
                If blnLis接口 Then
                    strTmp = ""
                    Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lng医嘱ID, 0, strTmp)
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    RaiseEvent PrintEPRReport(lngReportID, lngMenu = conMenu_Edit_Compend * 10# + 3)
                End If
            Else
                bln检验行 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "E"
                If Not bln检验行 Then bln配方行 = RowIs配方行(vsAdvice.Row)
                
                If bln检验行 Then
                    If blnLis接口 Then
                        strTmp = ""
                        Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lng医嘱ID, 1, strTmp)
                        If strTmp <> "" Then
                            MsgBox strTmp, vbInformation, gstrSysName
                            Exit Sub
                        End If
                    Else
                        '调用LisWork打印检验报告
                        blnPrint = IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, True, False)
                        If Not Open_LIS_Report(Me, lng医嘱ID, mlng病人ID, mblnMoved, blnPrint, Not bln打印) Then
                            MsgBox "该医嘱的报告为新版LIS产生，请使用(浏览检验结果)功能！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    '读取最近一次发送的NO,性质
                    If bln检验行 Or bln配方行 Then
                        '检验医嘱应以检验项目的NO为准
                        strSQL = "Select ID From 病人医嘱记录 Where 相关ID=[1] And Rownum=1"
                        strSQL = "Select 医嘱ID,NO,记录性质 From 病人医嘱发送 Where 医嘱ID=(" & strSQL & ") Order by 发送号 Desc"
                    Else
                        strSQL = "Select 医嘱ID,NO,记录性质 From 病人医嘱发送 Where 医嘱ID=[1] Order by 发送号 Desc"
                    End If
                    On Error GoTo errH
                                        If mblnMoved Then
                        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                    If Not rsTmp.EOF Then
                        strNO = NVL(rsTmp!NO): int性质 = NVL(rsTmp!记录性质, 0)
                    End If
                    
                    '按报表格式打印、预览报告
                    strSQL = "Select 编号 From 病历文件列表 Where ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_文件ID)))
                    If Not rsTmp.EOF Then
                        strBill = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-2"
                    End If
                    
                    If lngMenu = conMenu_Edit_Compend * 10# + 2 Then
                        If Not ReportPrintSet(gcnOracle, glngSys, strBill, Me) Then Exit Sub
                    End If
                    
                    If Not bln检验行 And Not bln配方行 Then
                        strFlagString = GetRPTPicture(mblnMoved, lngReportID, strBill, arrRPTPar)
                    End If
                    
                    If lngMenu <> conMenu_Edit_Compend * 10# + 2 And Not bln打印 Then
                        strTmp = "DisabledPrint=1"
                    Else
                        strTmp = "DisabledPrint=0"
                    End If
                    
                    '医嘱ID为采集方式的ID，即检验的相关ID
                    Call ReportOpen(gcnOracle, glngSys, strBill, Me, "NO=" & strNO, "性质=" & int性质, _
                        "医嘱ID=" & lng医嘱ID, _
                        strFlagString, _
                        arrRPTPar(0), arrRPTPar(1), arrRPTPar(2), arrRPTPar(3), arrRPTPar(4), arrRPTPar(5), _
                        arrRPTPar(6), arrRPTPar(7), arrRPTPar(8), arrRPTPar(9), arrRPTPar(10), arrRPTPar(11), _
                        arrRPTPar(12), arrRPTPar(13), arrRPTPar(14), arrRPTPar(15), arrRPTPar(16), arrRPTPar(17), _
                        arrRPTPar(18), arrRPTPar(19), strTmp, _
                        IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, 2, 1))
                End If
            End If
        End If
        '自动标记为已查阅：护士查阅不算
        If mblnAutoRead And mint场合 <> 1 Then Call FuncExecReportRead(True, True)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecReportRead(ByVal blnRead As Boolean, Optional ByVal blnAuto As Boolean)
'功能：设置当前报告为已查阅，或者取消当前报告的查阅状态
'参数：blnRead=已阅或者取消阅读状态
'      blnAuto=设置为已阅时，是否自动设置在调用
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strAdvice As String
    Dim strTmp As String
    Dim strErr As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) = 0 Then Exit Sub
        '新版PACS编辑器报告，直接调用接口标记已阅
        If .TextMatrix(.Row, COL_检查报告ID) = "" Then
            If Val(.TextMatrix(.Row, COL_报告ID)) = 0 Then Exit Sub
            If CheckOtherDeptPatiOpt = False Then Exit Sub
            
            If blnRead Then
                If Not blnAuto Then
                    If Val(.Cell(flexcpData, .Row, COL_查阅状态)) = 1 Then Exit Sub '自动标记时不计次数
                    If MsgBox("请确认该项目报告您已经仔细阅读了吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                strSQL = "Zl_报告查阅记录_Insert(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_报告ID)) & ")"
            Else
                If MsgBox("你确实要取消该报告的查阅状态吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                strSQL = "Zl_报告查阅记录_Cancel(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_报告ID)) & ",'" & UserInfo.姓名 & "')"
            End If
            Call InitObjLis(p住院医生站)
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, "FuncExecReportRead")
            If Not gobjLIS Is Nothing Then
                '检验调用标记接口
                strTmp = "Select ID From 病人医嘱记录 Where ID=[1] Or 相关ID=[1] order by 序号"
                Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
                Do While Not rsTmp.EOF
                    strAdvice = strAdvice & "," & rsTmp!ID
                    rsTmp.MoveNext
                Loop
                If .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "6" Then
                    gobjLIS.WriteAdvicesLookState Mid(strAdvice, 2), IIF(blnRead, 1, 0)
                End If
            End If
            On Error GoTo 0
        Else
            Call CreateObjectPacs(mobjPublicPACS)
            Call mobjPublicPACS.zlDocViewStateUpdate(blnRead, Val(.TextMatrix(.Row, COL_ID)))
        End If
        
        '设置界面状态
        If blnRead Then
            .Cell(flexcpData, .Row, COL_查阅状态) = 1 '我已查阅
        Else
            On Error GoTo errH
            strSQL = "Select Count(1) as 次数 From 报告查阅记录 Where 医嘱ID=[1] And 取消时间 Is Null"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
            If NVL(rsTmp!次数, 0) = 0 Then
                .Cell(flexcpData, .Row, COL_查阅状态) = 0 '我未查阅
            End If
        End If
        Call SetAdviceReportIcon(.Row)
        .TextMatrix(.Row, COL_查阅状态) = "查阅"
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckAdviceIsAduit() As Boolean
'判断医嘱是否核对
    Dim strSQL As String, rsTmp As Recordset
    Dim strTmp As String
    Dim lngTmp As String
    
    If Val(gstr医嘱核对) = 0 Then CheckAdviceIsAduit = True: Exit Function
    With vsAppend
        If .TextMatrix(.Row, COLSend("诊疗类别")) = "E" And .TextMatrix(.Row, COLSend("操作类型")) = "1" And Mid(gstr医嘱核对, 2, 1) = "1" Or _
           .TextMatrix(.Row, COLSend("操作类型")) = "8" And .TextMatrix(.Row, COLSend("诊疗类别")) = "E" And Mid(gstr医嘱核对, 1, 1) = "1" Or _
           .TextMatrix(.Row, COLSend("诊疗类别")) = "K" And Mid(gstr医嘱核对, 1, 1) = "1" Then
            strTmp = IIF(.TextMatrix(.Row, COLSend("操作类型")) = "1", "皮试", "输血")
            strSQL = "Select 核对人 From 病人医嘱执行 Where 医嘱id = [1] And 发送号 = [2]"
            On err GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COLSend("医嘱ID"))), Val(.TextMatrix(.Row, COLSend("发送号"))))
            If rsTmp.RecordCount = 1 Then
                If rsTmp!核对人 & "" <> "" Then
                    CheckAdviceIsAduit = True
                Else
                    MsgBox "当前医嘱是" & strTmp & "医嘱，必须核对了才能完成。", vbInformation, gstrSysName
                End If
            ElseIf rsTmp.RecordCount > 1 Then
                lngTmp = rsTmp.RecordCount
                rsTmp.Filter = "核对人<>''"
                If lngTmp <> rsTmp.RecordCount Then
                    MsgBox "当前医嘱是" & strTmp & "医嘱，存在未核对的执行登记，必须全部核对了才能完成。", vbInformation, gstrSysName
                Else
                    CheckAdviceIsAduit = True
                End If
            Else
                MsgBox "当前医嘱是" & strTmp & "医嘱，必须记录执行情况后核对了才能完成。", vbInformation, gstrSysName
            End If
        Else
            CheckAdviceIsAduit = True
        End If
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncExecFinish()
'功能：确认执行完成
    Dim rsTmp As New ADODB.Recordset
    Dim lng医嘱ID As Long, lng发送号 As Long, lng相关ID As Long
    Dim strSQL As String, strTest As String, blnTran As Boolean
    Dim str结果 As String, int结果 As Integer, strLabel As String
    Dim cnNew As ADODB.Connection, i As Long
    Dim strUserName As String, strOwner As String
    Dim dateInput As Date, blnIsAbnormal As Boolean
    Dim strSelect As String
    Dim strSelectInput As String
    Dim strTextInput As String
    Dim lng执行科室ID As Long

    Dim curMoney As Currency, str类别 As String, str类别名 As String

    With vsAppend
        lng医嘱ID = Val(.TextMatrix(.Row, COLSend("医嘱ID")))
        lng发送号 = Val(.TextMatrix(.Row, COLSend("发送号")))
        lng相关ID = Val(.TextMatrix(.Row, COLSend("相关ID")))
        lng执行科室ID = Val(.Cell(flexcpData, .Row, COLSend("执行科室")))
        If Val(.Cell(flexcpData, .Row, COLSend("执行状态"))) = 1 Then
            MsgBox "该执行项目当前已经执行完成。", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
        
        If Not CheckAdviceIsAduit Then Exit Sub
        
        '检查病人是否正在审核
        If Not CheckPatiIsAduit Then Exit Sub

        '是否允许完成未收费病人的项目:不管记帐划价,因为要执行后审核,临嘱才可能发送到门诊收费
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_期效) = "临嘱" And Val(.TextMatrix(.Row, COLSend("记录性质"))) = 1 And .Cell(flexcpData, .Row, COLSend("计费状态")) > 0 Then
            If Not ItemHaveCash(2, False, Val(.TextMatrix(.Row, COLSend("医嘱ID"))), Val(.TextMatrix(.Row, COLSend("相关ID"))), _
                Val(.TextMatrix(.Row, COLSend("发送号"))), .TextMatrix(.Row, COLSend("诊疗类别")), .TextMatrix(.Row, COLSend("单据号")), _
                    1, 0, 0, mblnMoved, CDate(.TextMatrix(.Row, COLSend("发送时间"))), "", "", blnIsAbnormal) Then
                If blnIsAbnormal Then
                    MsgBox "该病人还存在异常费用，请检查。", vbInformation, gstrSysName
                Else
                    MsgBox "该病人还存在未收费的费用，请检查。", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
        End If
        
        If Val(.TextMatrix(.Row, COLSend("记录性质"))) = 2 Then
            curMoney = GetAdviceMoney(IIF(lng相关ID = 0, lng医嘱ID, lng相关ID), lng医嘱ID, lng发送号, str类别, str类别名, False, _
                IIF(Val(.TextMatrix(.Row, COLSend("门诊记帐"))) = 0, 2, 1))
            If curMoney > 0 Then
                '住院出院病人费用控制
                If Not PatiCanBilling(mlng病人ID, mlng主页ID, GetInsidePrivs(p住院医嘱发送), p住院医嘱发送) Then Exit Sub
                '记帐报警
                If InitObjPublicExpense Then
                    If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, p住院医嘱发送, "", .TextMatrix(.Row, COLSend("单据号")), GetInsidePrivs(p住院医嘱发送), mlng病区ID) = False Then Exit Sub
                End If
                
                
                '门诊一卡通消费身份验证,只检查门诊记帐费用
                If gdbl预存款消费验卡 <> 0 And Val(.TextMatrix(.Row, COLSend("门诊记帐"))) = 1 Then
                    If Not zlDatabase.PatiIdentify(Me, glngSys, mlng病人ID, curMoney, , , , IIF(-1 * gdbl预存款消费验卡 >= Val(curMoney), False, True), , , (gdbl预存款消费验卡 <> 0), (2 = gdbl预存款消费验卡)) Then Exit Sub
                End If
            End If
        End If
    End With
    
    On Error GoTo errH

    '判断是否皮试,再填写结果
    strSQL = "Select A.诊疗类别,A.皮试结果,B.操作类型,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID" & IIF(mbln叮嘱发送执行, "(+)", "") & " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        '已经填写了皮试结果则不再填写
        If rsTmp!诊疗类别 = "E" And NVL(rsTmp!操作类型) = "1" And IsNull(rsTmp!皮试结果) Then
            '先作身份验证
            If mbln皮试验证 Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "在填写皮试结果前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "皮试医嘱结果", cnNew)
                If strUserName = "" Then Exit Sub
            End If
            '阳性
            For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(0), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(0), ",")(i) & "|0"
            Next
            '阴性
            For i = 0 To UBound(Split(Split(rsTmp!标本部位 & "", ";")(1), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!标本部位 & "", ";")(1), ",")(i) & "|0|2"
            Next
            strSelect = Mid(strSelect, 2)
            
            '填写皮试结果
            str结果 = zlCommFun.ShowMsgBox("皮试结果", vsAdvice.TextMatrix(vsAdvice.Row, col_医嘱内容) & "：^^请根据过敏试验结果选择相应的按钮操作。", _
            "确定(&O),?取消(&C)", Me, vbQuestion, "皮试时间", dateInput, "yyyy-MM-dd HH:mm", "皮试结果(&P):" & strSelect, strSelectInput, _
            "过敏反应(&F)", 50, strTextInput, , True)
            
            If str结果 = "" Then Exit Sub
            If strSelectInput = "" Then Exit Sub
            If mbln护士签名 Then
                If Not (Check电子签名) Then Exit Sub
            End If
            Call GetTestLabel(rsTmp!标本部位, strSelectInput, strLabel, int结果)
            strTest = "ZL_病人医嘱记录_皮试(" & lng医嘱ID & ",'" & strLabel & "'," & int结果 & _
            ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
        End If
    Else
        MsgBox "对应的医嘱记录不存在，无法完成操作。", vbInformation, gstrSysName
        Exit Sub
    End If
    If strTest = "" Then
        If MsgBox("确认该执行项目执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    strSQL = "ZL_病人医嘱执行_Finish(" & lng医嘱ID & "," & lng发送号 & ",Null,0,'" & UserInfo.编号 & "','" & UserInfo.姓名 & "'," & lng执行科室ID & ")"

    If strTest <> "" And Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)

        On Error GoTo errNew
        cnNew.BeginTrans

        Call SQLTest(App.ProductName, Me.Caption, strTest)
        cnNew.Execute strOwner & "." & strTest, , adCmdStoredProc
        Call SQLTest

        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        cnNew.Execute strOwner & "." & strSQL, , adCmdStoredProc
        Call SQLTest

        cnNew.CommitTrans
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        gcnOracle.BeginTrans: blnTran = True
        If strTest <> "" Then
            Call zlDatabase.ExecuteProcedure(strTest, Me.Caption)
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        gcnOracle.CommitTrans: blnTran = False
    End If

    'Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Call LoadAdvice
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
errNew:
    cnNew.RollbackTrans
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Sub FuncExecCancel()
'功能：取消执行完成
    Dim lng组ID As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim str诊疗类别 As String, strSQL As String
    Dim byt来源 As Byte, lng执行科室ID As Long
    Dim strOwner As String, strUserName As String
    Dim cnNew As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim bln清除皮试结果 As Boolean
    
    With vsAppend

        '必须是已执行才可以取消
        If Val(.Cell(flexcpData, .Row, COLSend("执行状态"))) <> 1 Then
            MsgBox "该执行项目当前不处于已执行状态，不能取消执行。", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
        
        '检查病人是否正在审核
        If Not CheckPatiIsAduit Then Exit Sub
        
        lng执行科室ID = Val(.Cell(flexcpData, .Row, COLSend("执行科室")))
        lng医嘱ID = Val(.TextMatrix(.Row, COLSend("医嘱ID")))
        lng发送号 = Val(.TextMatrix(.Row, COLSend("发送号")))
        str诊疗类别 = .TextMatrix(.Row, COLSend("诊疗类别"))
        lng组ID = IIF(Val(.TextMatrix(.Row, COLSend("相关ID"))) = 0, lng医嘱ID, Val(.TextMatrix(.Row, COLSend("相关ID"))))
    
        If Val(.TextMatrix(.Row, COLSend("记录性质"))) <> 1 Then
            If Val(.TextMatrix(.Row, COLSend("门诊记帐"))) = 0 Then
                byt来源 = 2
            Else
                byt来源 = 1
            End If
            '费用结帐判断
            If Not ItemCanCancel(lng医嘱ID, lng发送号, lng组ID, str诊疗类别, False, mblnMoved, byt来源) Then Exit Sub
        End If
    End With
    
    If MsgBox("确实要将该执行项目取消执行完成吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '判断是否皮试,再填写结果
    strSQL = "Select A.诊疗类别,A.皮试结果,B.操作类型,Nvl(B.标本部位,'阳性(+);阴性(-)') as 标本部位 From 病人医嘱记录 A,诊疗项目目录 B Where A.诊疗项目ID=B.ID" & IIF(mbln叮嘱发送执行, "(+)", "") & " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        '已经填写了皮试结果则不再填写
        If rsTmp!诊疗类别 = "E" And NVL(rsTmp!操作类型) = "1" And Not IsNull(rsTmp!皮试结果) Then
            '先作身份验证
            If mbln皮试验证 Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "在取消完成皮试医嘱前，请您先输入用户名和密码进行身份验证。", glngSys, p住院医嘱发送, "皮试医嘱结果", cnNew)
                If strUserName = "" Then Exit Sub
                bln清除皮试结果 = True
            Else
                If MsgBox("是否清除皮试结果？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    bln清除皮试结果 = False
                Else
                    bln清除皮试结果 = True
                End If
            End If
            strSQL = "ZL_病人医嘱执行_Cancel(" & lng医嘱ID & "," & lng发送号 & "," & IIF(bln清除皮试结果, 1, 0) & ",0," & lng执行科室ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        Else
            strSQL = "ZL_病人医嘱执行_Cancel(" & lng医嘱ID & "," & lng发送号 & "," & "Null,0," & lng执行科室ID & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
        End If
    End If
    If Not cnNew Is Nothing Then
        strOwner = SystemOwner(glngSys)

        On Error GoTo errNew
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        cnNew.Execute strOwner & "." & strSQL, , adCmdStoredProc
        Call SQLTest
        cnNew.Close: Set cnNew = Nothing
        On Error GoTo errH
    Else
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End If
    'Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '要更新执行状态
    Call LoadAdvice
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
errNew:
    MsgBox err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
    cnNew.Close: Set cnNew = Nothing
End Sub

Private Function FuncThingNew(Optional ByVal blnRefresh As Boolean = True) As Boolean
    Dim lng科室id As Long, lng执行科室ID As Long
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim blnOK As Boolean
    
    With vsAppend
        If Val(.Cell(flexcpData, .Row, COLSend("执行状态"))) = 1 Then
            MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
            Exit Function
        End If
        If CheckDataMoved Then Exit Function
        
        lng科室id = mlng病区ID
        lng医嘱ID = Val(.TextMatrix(.Row, COLSend("医嘱ID")))
        lng发送号 = Val(.TextMatrix(.Row, COLSend("发送号")))
        
        RaiseEvent ExecLogNew(lng医嘱ID, lng发送号, lng科室id, blnOK)
        If blnOK Then
            If blnRefresh Then
                Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '可能要更新执行状态
            End If
            FuncThingNew = True
        End If
    End With
End Function

Private Sub FuncThingModi()
    Dim lng科室id As Long, lng医嘱ID As Long, lng发送号 As Long
    Dim str执行时间 As String, blnOK As Boolean
        
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub '只能操作最近一次执行

    If Val(gstr医嘱核对) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) <> "" Then
        MsgBox "该医嘱还已经核对，请取消核对后再试。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsAppend
        If Val(.Cell(flexcpData, .Row, COLSend("执行状态"))) = 1 Then
            MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
        
        lng科室id = mlng病区ID
        lng医嘱ID = Val(.TextMatrix(.Row, COLSend("医嘱ID")))
        lng发送号 = Val(.TextMatrix(.Row, COLSend("发送号")))
        str执行时间 = vsExec.Cell(flexcpData, vsExec.Row, 1)
    
        RaiseEvent ExecLogModi(lng医嘱ID, lng发送号, lng科室id, str执行时间, blnOK)
        If blnOK Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '可能要更新执行状态
    End With
End Sub

Private Sub FuncThingDel()
    Dim lng医嘱ID As Long, lng发送号 As Long
    Dim str执行时间 As String, strSQL As String
    
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub '只能操作最近一次执行

    With vsAppend
        If Val(.Cell(flexcpData, .Row, COLSend("执行状态"))) = 1 Then '子项和独项同执行状态
            MsgBox "该执行项目已经执行完成，不能再继续操作。", vbInformation, gstrSysName
            Exit Sub
        End If
        If Val(gstr医嘱核对) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("核对人")) <> "" Then
            MsgBox "该医嘱还已经核对，请取消核对后再试。", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
            
        If MsgBox("确实要删除该条执行情况吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lng医嘱ID = Val(.TextMatrix(.Row, COLSend("医嘱ID")))
        lng发送号 = Val(.TextMatrix(.Row, COLSend("发送号")))
        str执行时间 = vsExec.Cell(flexcpData, vsExec.Row, 1)
    
        strSQL = "ZL_病人医嘱执行_Delete(" & lng医嘱ID & "," & lng发送号 & ",To_Date('" & str执行时间 & "','YYYY-MM-DD HH24:MI:SS'),0,0," & mlng病区ID & ")"
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '要更新执行状态'可能要更新执行状态
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbcAppend_GotFocus()
    If vsAppend.Visible And vsAppend.Enabled Then
        vsAppend.SetFocus
    ElseIf rtfAppend.Visible And rtfAppend.Enabled Then
        rtfAppend.SetFocus
    End If
End Sub

Private Sub tbcAppend_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim blnDo As Boolean
    
    If Item.Tag = "" Then Exit Sub
    
    If Visible Then
        If Decode(vsAppend.Tag, "计价", True, "发送", True, "签名", True, False) Then
            Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
        End If
    End If
    vsAppend.Tag = Item.Tag '用于公共函数区分个性化
    
    Call SetExecShow(False, False)
    
    If Item.Tag = "计价" Then
        Call InitPriceTable
    ElseIf Item.Tag = "发送" Then
        Call InitSendTable
        Call InitExecTable '实际只需执行一次即可
    ElseIf Item.Tag = "签名" Then
        Call InitSignTable
    ElseIf Item.Tag = "附项" Then
        'NoneCode
    ElseIf Item.Tag = "安排" Then
        'NoneCode
    ElseIf Item.Tag = "配药" Then
    
    End If
    
    If Visible Then
        If Decode(Item.Tag, "计价", True, "发送", True, "签名", True, False) Then
            Call RestoreFlexState(vsAppend, App.ProductName & "\" & Me.Name)
        End If
    End If
    
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    
    If Visible Then
        If vsAppend.Visible And vsAppend.Enabled Then
            vsAppend.SetFocus
        ElseIf rtfAppend.Visible And rtfAppend.Enabled Then
            rtfAppend.SetFocus
        End If
    End If
End Sub

Private Sub SetExecShow(ByVal blnBar As Boolean, ByVal blnData As Boolean, Optional blnBloodExec As Boolean = False)
    Dim blnDo As Boolean, blnBlood As Boolean
    Dim lngH As Long
    
    If Not blnBar Then blnData = False
    If blnData Then blnBar = True
    
    picAppend.Tag = "不执行"
    If blnBar Then
        If blnBloodExec = False Then
            If picExec.Tag = "" Then
                lngH = vsAppend.Height - (vsAppend.Top + vsAppend.RowPos(vsAppend.Rows - 1) + vsAppend.RowHeight(vsAppend.Rows - 1) * 2)
                If lngH < picExec.Height Then
                    tbcAppend.Height = tbcAppend.Height + picExec.Height
                End If
                picExec.Visible = True: picExec.Tag = "可见"
                blnDo = True
            End If
            If picBlood.Tag = "可见" Then
                picBlood.Visible = False: picBlood.Tag = "": DkpBlood.Tag = ""
                blnDo = True: blnBlood = True
            End If
        Else
            If picBlood.Tag = "" Then
                picBlood.Visible = True: picBlood.Tag = "可见"
                Call DkpBlood_AttachPane(DkpBlood.Panes(1))
                If Not mobjFrmBlood Is Nothing Then
                    mobjFrmBlood.IsShowExec = mblnShowExec
                End If
                blnDo = True
            End If
            If picExec.Tag = "可见" Then
                picExec.Visible = False: picExec.Tag = ""
                blnDo = True
            End If
        End If
    Else
        If picExec.Tag = "可见" Then
            picExec.Visible = False: picExec.Tag = ""
            blnDo = True
        End If
        If picBlood.Tag = "可见" Then
            picBlood.Visible = False: picBlood.Tag = "": DkpBlood.Tag = ""
            blnDo = True
        End If
    End If
    
    If blnData Then
        If vsExec.Tag = "" Then '可见时Tag=1
            lngH = vsAppend.Height - IIF(picExec.Tag = "可见", picExec.Height, 0) - (vsAppend.Top + vsAppend.RowPos(vsAppend.Rows - 1) + vsAppend.RowHeight(vsAppend.Rows - 1) * 2)
            If lngH < vsExec.Height + fraExecUD.Height Then
                tbcAppend.Height = tbcAppend.Height + fraExecUD.Height + vsExec.Height
            End If
            
            fraExecUD.Visible = True: vsExec.Visible = True: vsExec.Tag = "可见"
            blnDo = True
        End If
    Else
        If vsExec.Tag = "可见" Then
            fraExecUD.Visible = False: vsExec.Visible = False: vsExec.Tag = ""
            blnDo = True
        End If
        '输血执行blnData始终为False
        If picBlood.Tag = "可见" Then
            fraExecUD.Visible = True
            blnDo = True
        End If
    End If
    If picBlood.Tag = "" Then TimShow.Enabled = False
    picAppend.Tag = ""
    
    If blnDo Then
        Call picAppend_Resize
        Call cbsSub_Resize
    End If
    TimShow.Enabled = (picBlood.Tag = "可见")
End Sub

Private Sub TimShow_Timer()
    Dim blnShowExec As Boolean
    Dim lngH As Long
    On Error GoTo ErrHand
    If picBlood.Visible = False Then Exit Sub
    If Not mobjFrmBlood Is Nothing Then
        blnShowExec = mobjFrmBlood.IsShowExec
        If DkpBlood.Tag <> IIF(blnShowExec, "可见", "不可见") Then
            DkpBlood.Tag = IIF(blnShowExec, "可见", "不可见")
            picBlood.Height = mobjFrmBlood.Height
            Call SetExecShow(True, False, True)
        End If
    End If
    Exit Sub
ErrHand:
    TimShow.Enabled = False
End Sub

Private Sub vsAdvice_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If fraMore.Visible = True Then
        fraMore.Tag = ""
        fraMore.Visible = False
        PicAdviceDetail.Visible = False
    End If
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnExist As Boolean, blnSel As Boolean, bln输血 As Boolean
    Dim varDraw As RedrawSettings, intIdx As Integer
    Dim i As Integer
    
    If NewRow = OldRow And vsAdvice.Visible = False Then Exit Sub
     'PASS
    If mblnPass And Me.Visible Then
        If NewRow <> OldRow Then
            Call gobjPass.zlPassSetDrug(mobjPassMap)
        End If
    End If
    With vsAdvice
        If fraMore.Visible = True Then fraMore.BackColor = .BackColorSel
        If .Col >= .FixedCols Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, COL_开始时间)
        End If
        If .Redraw <> flexRDNone Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                If mint场合 = 1 And OldRow <> -1 And OldCol <> -1 And Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 And .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "1" Then
                    For i = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(i).Visible And tbcAppend(i).Tag = "发送" Then
                            tbcAppend.Item(i).Selected = True
                            Exit For
                        End If
                    Next
                End If
            
                '显示报告是否我已阅读
                If Val(.TextMatrix(NewRow, COL_报告ID)) <> 0 Or .TextMatrix(NewRow, COL_检查报告ID) <> "" Then
                    On Error GoTo errH
                    strSQL = "Select 1 From 报告查阅记录 Where 医嘱ID=[1]  And 查阅人=[2] And 取消时间 Is NULL"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", _
                        Val(.TextMatrix(NewRow, COL_ID)), UserInfo.姓名)
                    If Not rsTmp.EOF Then
                        If .TextMatrix(NewRow, COL_检查报告ID) = "" Then
                            .Cell(flexcpData, NewRow, COL_查阅状态) = 1
                        Else
                            '部分查阅的
                            strSQL = "Select 1 From 病人医嘱报告 A Where not exists(select 1 from 报告查阅记录 B where B.医嘱ID=A.医嘱ID And A.检查报告ID=B.检查报告ID And B.查阅人=[2] And B.取消时间 Is NULL) and A.医嘱ID=[1] "
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", Val(.TextMatrix(NewRow, COL_ID)), UserInfo.姓名)
                            .Cell(flexcpData, NewRow, COL_查阅状态) = IIF(Not rsTmp.EOF, 2, 1)
                        End If
                    Else
                        .Cell(flexcpData, NewRow, COL_查阅状态) = 0
                    End If
                    On Error GoTo 0
                End If
                
                mbln确认会诊 = False
                If NewRow <> 0 And Val(.TextMatrix(NewRow, COL_申请序号)) <> 0 And Val(.TextMatrix(NewRow, COL_操作类型)) = 7 And .TextMatrix(NewRow, COL_诊疗类别) = "Z" And .TextMatrix(NewRow, COL_状态) = "停止" Then
                    mbln确认会诊 = Get确认会诊(Val(.TextMatrix(NewRow, COL_ID)))
                End If
                
                '显示医嘱附加表格的内容
                If mblnAppend Then
                    '判断单据附项是否有内容
                    blnSel = False: blnExist = False
                    Call ShowBillAppend(NewRow, blnExist)
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "附项" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '根据条件屏蔽重复调用
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '判断附加信息的显示
                    blnSel = False: blnExist = False
                    Call ShowAdvicePlan(NewRow, blnExist)
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "安排" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '根据条件屏蔽重复调用
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '判断输液配药的显示
                    blnSel = False: blnExist = False
                    If mint场合 = 1 Then
                        Call ShowCompoundInfo(NewRow, blnExist)
                    End If
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "配药" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '根据条件屏蔽重复调用
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '判断医嘱是否审核(医嘱作废不显示)
                    blnSel = False: blnExist = False: bln输血 = False
                    If gbln血库系统 And .TextMatrix(NewRow, COL_诊疗类别) = "K" Then
                        bln输血 = True
                        With vsAdvice
                            '用血医嘱审核状态=1表明是输血科发血产生的待核对医嘱，对于输血医嘱，审核状态=4，紧急医嘱和未用输血分级管理时，显示为等待配血
                            If Val(.TextMatrix(NewRow, COL_审核状态)) = 1 And Val(.TextMatrix(NewRow, COL_检查方法)) = 1 Then
                                blnExist = True
                            Else
                                blnExist = InStr(",,2,3,4,5,6,", "," & .TextMatrix(NewRow, COL_审核状态) & ",") > 0 And Not (.TextMatrix(NewRow, COL_医嘱状态) = "4")
                            End If
                        End With
                    End If
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "血液" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '根据条件屏蔽重复调用
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    blnSel = False: blnExist = False
                    If bln输血 = False Then
                        With vsAdvice
                            blnExist = InStr(",2,3,4,5,", "," & .TextMatrix(NewRow, COL_审核状态) & ",") > 0
                            '是输血医嘱时，用血库系统后才有为4的审核状态。紧急医嘱，未用输血分级管理时。 审核状态为4时没有相应的操作记录<病人医嘱状态>
                            If Val(.TextMatrix(NewRow, COL_审核状态)) = 4 And .TextMatrix(NewRow, COL_诊疗类别) = "K" Then
                                If Val(.TextMatrix(NewRow, COL_标志)) = 1 Or Not gbln输血分级管理 Then blnExist = False
                            End If
                        End With
                    End If
                    
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "其他" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '根据条件屏蔽重复调用
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '判预约信息的显示
                    blnSel = False: blnExist = False
                    Call ShowAdviceRISSch(NewRow, blnExist)
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "预约" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '根据条件屏蔽重复调用
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    If tbcAppend.Selected.Tag = "计价" Then
                        Call ShowPrice(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "发送" Then
                        If NewRow <> 0 And Val(.TextMatrix(NewRow, COL_申请序号)) <> 0 And Val(.TextMatrix(NewRow, COL_操作类型)) = 7 And .TextMatrix(NewRow, COL_诊疗类别) = "Z" Then
                            vsAppend.ColHidden(COLSend("接受时间")) = False
                            vsAppend.ColHidden(COLSend("接受人")) = False
                            vsAppend.ColHidden(COLSend("到场时间")) = False
                        Else
                            vsAppend.ColHidden(COLSend("接受时间")) = True
                            vsAppend.ColHidden(COLSend("接受人")) = True
                            vsAppend.ColHidden(COLSend("到场时间")) = True
                        End If
                        Call ShowSendList(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "签名" Then
                        Call ShowSignList(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "附项" Then
                        '前面已固定读取
                    ElseIf tbcAppend.Selected.Tag = "预约" Then
                        '前面已固定读取
                    ElseIf tbcAppend.Selected.Tag = "安排" Then
                        '前面已固定读取
                    ElseIf tbcAppend.Selected.Tag = "配药" Then
                        intIdx = IIF(vsAdvice.TextMatrix(NewRow, COL_期效) = "长嘱", 0, 1)
                        Call mfrmCompoundMedicine.RefreshData(Val(vsAdvice.TextMatrix(NewRow, COL_相关ID)), mlng病区ID, mlng病人ID, mlng主页ID, mlng病人性质, intIdx, mclsMipModule, mfrmParent)
                    ElseIf tbcAppend.Selected.Tag = "其他" Then
                        Call ShowOtherAppend(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "血液" Then
                        If Not mobjFrmBloodList Is Nothing Then
                            Call mobjFrmBloodList.zlRefresh(Val(vsAdvice.TextMatrix(NewRow, COL_ID)), mlngFontSize, mblnMoved)
                        End If
                    End If
                End If
                
                '显示医嘱可回退内容
                Call LoadRollList(NewRow)
                
                If (Not mblnShowExec) And mint场合 = 1 And OldRow <> -1 And OldCol <> -1 And Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 And .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "1" And tbcAppend.Selected.Tag = "发送" Then
                    mblnShowExec = Not mblnShowExec
                    Call SetExecShow(True, mblnShowExec)
                    Call vsAppend_AfterRowColChange(-1, -1, vsAppend.Row, vsAppend.Col)
                End If
                
            ElseIf mblnAppend Then
                Call ClearAppendData
            End If
            Call LoadBillList '显示可打印的诊疗单据
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col_医嘱内容 Or Col = col_内容 Then
        vsAdvice.AutoSize Col, COL_用法
    ElseIf Col = COL_皮试 Then
        If vsAdvice.ColWidth(Col) > 1200 Then vsAdvice.ColWidth(Col) = 1200
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = -1 Then
        If Col <= vsAdvice.FixedCols - 1 Then
            Cancel = True
        ElseIf Col = COL_警示 Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_Click()
'功能：查阅报告
    Dim lngMouseRow As Long, lngMouseCol As Long
    
    'PASS
    If mblnPass And Me.Visible Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap, 1)
    End If
    
    If mvarCond.过滤模式 <> 3 Then Exit Sub
    With vsAdvice
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        
        If lngMouseRow > -1 And lngMouseCol > -1 Then
            If .Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
                .Redraw = False
                Call FuncEPRReport(conMenu_Edit_Compend)
                .Cell(flexcpForeColor, lngMouseRow, COL_查阅状态) = &H80& '暗红
                .Redraw = True
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_DblClick()
    Dim lng医嘱ID As Long
    Dim lngNo As Long
    Dim bln用血 As Boolean
    'PASS
    If mblnPass Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
    End If
    '双击的医嘱如果是申请单方式下达的弹出查看界面 输血，手术，会诊，检查，检验
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_申请序号))
        
        If lng医嘱ID <> 0 And lngNo <> 0 Then
            If .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                '输血
                If Val(Mid(gstrInUseApp, 3, 1)) = 1 Then
                    bln用血 = Val(.TextMatrix(.Row, COL_检查方法)) = 1
                    If gbln血库系统 = True Then
                        Call frmApplyBloodNew.ShowMe(Me, mlng病人ID, mlng主页ID, 0, 2, lng医嘱ID, mlng科室ID, mlng病区ID, Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2), mintPState, , mrsDefine, mclsMipModule, , , , , mbyt婴儿, , mlng前提ID, IIF(bln用血 = True, 1, 0))
                    Else
                        Call frmApplyBlood.ShowMe(Me, mlng病人ID, mlng主页ID, 0, 2, lng医嘱ID, mlng科室ID, mlng病区ID, Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2), mintPState, , mrsDefine, mclsMipModule, , , , , mbyt婴儿, , mlng前提ID)
                    End If
                End If
                
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                '手术
                If Val(Mid(gstrInUseApp, 4, 1)) = 1 Then Call frmApplyOperation.ShowMe(Me, 0, 2, mlng病人ID, mlng主页ID, 0, lng医嘱ID, , , , , , , , , , , , mbyt婴儿)
               
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "Z" And .TextMatrix(.Row, COL_操作类型) = "7" Then
                '会诊
                If Val(Mid(gstrInUseApp, 5, 1)) = 1 Then Call frmApplyConsultation.ShowMe(Me, lng医嘱ID, lngNo, 2, , mlng病人ID, mlng主页ID, , , , , , , , , , mbyt婴儿)
                 
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                '检查
                If Val(Mid(gstrInUseApp, 1, 1)) = 1 Then
                    Call ShowApply检查(Me, lngNo)
                End If
                
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "E" And .TextMatrix(.Row, COL_操作类型) = "6" Then
                '检验
            End If
        End If
    End With
End Sub

Private Function GetPatiInfo(ByVal lng病人ID As Long, ByVal lng主页ID As Long) As ADODB.Recordset
'功能：根据病人ID、主页ID获取病人基本信息
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select A.住院号, A.当前床号, A.出生日期, Nvl(B.姓名, A.姓名) 姓名, Nvl(B.性别, A.性别) 性别, Nvl(B.年龄, A.年龄) 年龄, A.门诊号, A.健康号,b.费别" & vbNewLine & _
            "From 病人信息 A, 病案主页 B" & vbNewLine & _
            "Where A.病人id = B.病人id And A.病人id = [1] And B.主页id = [2]"

    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng病人ID, lng主页ID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long

    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '擦除固定列中的表格线
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '仅左边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅上边表格线
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅下边表格线
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '仅右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '擦除一并给药相关行列的边线及内容
            lngLeft = COL_期效: lngRight = COL_开始时间
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_天数: lngRight = COL_用法
            End If
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_皮试: lngRight = COL_皮试
            End If
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            
            If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '擦除左边表格线
            vRect.Right = Right - 1 '保留右边表格线
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '首行保留文字内容
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '底行保留下边线
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                '为了支持预览输出
                If .TextMatrix(Row, Col) <> "" Then .TextMatrix(Row, Col) = ""
            End If
            
            If Between(Row, .Row, .RowSel) Then
                SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
        Done = True
    End With
End Sub

Private Sub vsAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, strPrompt As String
    
    With vsAdvice
        lngRow = .MouseRow
        If Button = 0 And lngRow > 0 Then  '简洁模式才显该列
            If .MouseCol = col_内容 Then
                If Val(fraMore.Tag) <> lngRow Then
                    If InStr(.TextMatrix(lngRow, col_内容), "重整医嘱") = 0 Then
                     
                        fraMore.Visible = False
                        fraMore.Tag = lngRow
                        If lngRow = .Row Then
                            fraMore.BackColor = .BackColorSel
                        Else
                            fraMore.BackColor = .BackColor
                        End If
                        fraMore.Height = .RowHeight(lngRow) - 10
                        If fraMore.Height > 250 Then fraMore.Height = 250
                        
                        fraMore.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - fraMore.Height
                        If fraMore.Top + fraMore.Height > .Top + .Height - IIF(Grid.HScrollVisible(vsAdvice), 230, 0) Then Exit Sub
                        
                        fraMore.Left = .Left + .ColPos(col_内容) + IIF(.ColWidth(col_内容) > .ColWidthMax, .ColWidthMax, .ColWidth(col_内容)) - fraMore.Width
                        
                        fraMore.Visible = True
                        
                    Else
                        fraMore.Tag = ""
                        fraMore.Visible = False
                        PicAdviceDetail.Visible = False
                    End If
                ElseIf PicAdviceDetail.Visible = True Then
                    fraMore.Tag = ""
                    fraMore.Visible = False
                    PicAdviceDetail.Visible = False
                End If
            Else
                If fraMore.Visible Then
                    fraMore.Tag = ""
                    fraMore.Visible = False
                    PicAdviceDetail.Visible = False
                End If
                
                strPrompt = ""
                If .MouseCol = COL_F标志 Then
                    If Val(.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
                        strPrompt = "自由录入的医嘱"
                    ElseIf Val(.TextMatrix(lngRow, COL_标志)) = 1 Then
                        strPrompt = "紧急医嘱"
                    ElseIf Val(.TextMatrix(lngRow, COL_标志)) = 2 Then
                        strPrompt = "补录医嘱"
                    ElseIf .TextMatrix(lngRow, COL_频率) = "必要时" Or .TextMatrix(lngRow, COL_频率) = "需要时" Then
                        strPrompt = "备用医嘱"
                    End If
                     
                     '如果有抗菌用药审核信息，优先显示
                    If Val(.TextMatrix(lngRow, COL_医嘱状态)) = 1 Then
                        Select Case Val(.TextMatrix(lngRow, COL_审核状态))
                        Case 1
                            If .TextMatrix(lngRow, COL_诊疗类别) = "K" And Val(.TextMatrix(lngRow, COL_检查方法)) = 1 Then '用血医嘱审核
                                strPrompt = "用血医嘱待核对"
                            Else
                                strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "F", "手术", "K", "输血", "抗菌用药") & "待审核"
                            End If
                        Case 2
                            If Not (.TextMatrix(lngRow, COL_诊疗类别) = "K" And Val(.TextMatrix(lngRow, COL_检查方法)) = 1) Then
                                strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "F", "手术", "K", "输血", "抗菌用药") & "审核通过"
                            End If
                        Case 3
                            strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "F", "手术", "K", "输血", "抗菌用药") & "审核未通过:" & GetKSSAuditQuestion(Val(.TextMatrix(lngRow, COL_ID)))
                        Case 7
                            strPrompt = Decode(.TextMatrix(lngRow, COL_诊疗类别), "F", "手术", "K", "输血", "抗菌用药") & "待签发"
                        Case 4
                            If gbln血库系统 = False Then strPrompt = "输血待血库审核"
                        Case 5
                            If gbln血库系统 = False Then strPrompt = "输血血库正在配血"
                        End Select
                    End If
                ElseIf .MouseCol = COL_查阅状态 Then
                    If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then strPrompt = "报告未出"
                    If Val(.TextMatrix(lngRow, COL_报告ID)) <> 0 Or .TextMatrix(lngRow, COL_检查报告ID) <> "" Or _
                        Val(.TextMatrix(lngRow, COL_RIS报告ID)) <> 0 Or Val(.TextMatrix(lngRow, COL_LIS报告ID)) <> 0 Then
                        
                        If Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 0 Then
                            strPrompt = "报告未阅，点击查看"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 1 Then
                            strPrompt = "报告已阅，点击查看"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 2 Then
                            strPrompt = "报告部分已阅，点击查看"
                        End If
                    End If
                ElseIf .MouseCol = COL_F报告 Then
                    strPrompt = GetAdviceReportTip(lngRow)
                End If
            End If
            If .MouseRow > -1 And .MouseCol > -1 And mvarCond.过滤模式 = 3 And .MouseCol = COL_查阅状态 Then
                If .Cell(flexcpFontUnderline, .MouseRow, .MouseCol) = True Then
                    .MousePointer = 99
                Else
                    .MousePointer = 0
                End If
            Else
                .MousePointer = 0
            End If
            If strPrompt <> "" Then
                Call zlCommFun.ShowTipInfo(.hwnd, strPrompt)
                mlngPromptRow = lngRow
            ElseIf mvarCond.过滤模式 = 3 And strPrompt = "" Then
                Call zlCommFun.ShowTipInfo(.hwnd, "")
                mlngPromptRow = 0
            ElseIf mlngPromptRow <> 0 And lngRow <> mlngPromptRow Then
            '隐藏之前的提示内容
                Call zlCommFun.ShowTipInfo(.hwnd, "")
                mlngPromptRow = 0
            End If
        End If
    End With
End Sub

Private Sub vsfAdivceDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraMore.Tag = ""
    fraMore.Visible = False
    PicAdviceDetail.Visible = False
End Sub

Private Sub imgMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If PicAdviceDetail.Visible = False And vsAdvice.MouseRow > 0 Then
        Call LoadAdviceDetail(vsAdvice.MouseRow)
    End If
End Sub

Private Sub LoadAdviceDetail(lngRow As Long)
'功能：显示某行医嘱的详细内容
    Dim i As Long, j As Long
        
    vsfAdivceDetail.Redraw = flexRDNone
    vsfAdivceDetail.Clear
    vsfAdivceDetail.Rows = vsfAdivceDetail.FixedRows
    vsfAdivceDetail.Cols = 2
    j = 0
    With vsAdvice
        For i = 0 To .Cols - 1
             If .Cell(flexcpData, 0, i) = "Detail" Then
                j = j + 1
                vsfAdivceDetail.Rows = vsfAdivceDetail.FixedRows + j
                If .TextMatrix(0, i) = "确认停嘱时间" Then
                    vsfAdivceDetail.TextMatrix(j - 1, 0) = "确认时间" & "："
                Else
                    vsfAdivceDetail.TextMatrix(j - 1, 0) = .TextMatrix(0, i) & "："
                End If
                vsfAdivceDetail.TextMatrix(j - 1, 1) = .TextMatrix(lngRow, i)
                
                vsfAdivceDetail.Col = 0: vsfAdivceDetail.Row = j - 1
                vsfAdivceDetail.CellForeColor = &H8000000C
             End If
        Next
    End With
    
    With vsfAdivceDetail
        If .Rows > 0 Then
            .AutoSize 0, 1
            .Height = IIF(.RowHeight(0) < .RowHeightMin, .RowHeightMin, .RowHeight(0)) * .Rows + 100
            .Width = .ColWidth(0) + .ColWidth(1)
            .Row = -1
            
            PicAdviceDetail.Height = .Height
            PicAdviceDetail.Width = .Width
            PicAdviceDetail.Left = fraMore.Left + fraMore.Width
            If PicAdviceDetail.Height + fraMore.Top + fraMore.Height > Me.Top + Me.Height Then
                PicAdviceDetail.Top = fraMore.Top + fraMore.Height - PicAdviceDetail.Height - 10
            Else
                PicAdviceDetail.Top = fraMore.Top - 10 '避免顶端和表格线重合
            End If
            
            Call SetPicAdviceDetailEffect
            If PicAdviceDetail.Visible = False Then PicAdviceDetail.Visible = True
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetPicAdviceDetailEffect()
    Dim lngR As Long
    
    '边框：API=RoundRect
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, 0)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (0, Screen.TwipsPerPixelY)-(0, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (PicAdviceDetail.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
           
    '形状
    lngR = CreateRoundRectRgn(0, 0, PicAdviceDetail.ScaleX(PicAdviceDetail.Width, PicAdviceDetail.ScaleMode, vbPixels) + 1, PicAdviceDetail.ScaleY(PicAdviceDetail.Height, PicAdviceDetail.ScaleMode, vbPixels) + 1, 2, 2)
    Call SetWindowRgn(PicAdviceDetail.hwnd, lngR, False)

End Sub

Private Sub vsfAdivceDetail_LostFocus()
    PicAdviceDetail.Visible = False
End Sub

Private Sub vsAdvice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button = 2 Then
        If mcbsMain Is Nothing Then Exit Sub
        Set objPopup = mcbsMain.ActiveMenuBar.FindControl(, conMenu_EditPopup)
        If Not objPopup Is Nothing Then
            mintBillPrint = 0
            objPopup.CommandBar.ShowPopup
        End If
    ElseIf Button = 1 Then
        If mblnPass And Me.Visible Then
            Call gobjPass.zlPassCloseHint
        End If
    End If
End Sub

Private Sub OutputList(bytStyle As Byte)
'功能：输入出列表
'参数：bytStyle=1-打印,2-预览,3-输出到Excel
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mlng病人ID = 0 Then Exit Sub
    strSQL = "Select NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别,NVL(B.年龄,A.年龄) 年龄 ,B.住院号,B.出院病床 as 床号,B.入院日期,B.出院日期" & _
        " From 病人信息 A,病案主页 B Where A.病人ID=B.病人ID And A.病人ID=[1] And B.主页ID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID)
    If rsTmp.EOF Then Exit Sub
    
    '表头
    objOut.Title.Text = "病人医嘱清单"
    objOut.Title.Font.Name = "楷体_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '表上
    Set objRow = New zlTabAppRow
    objRow.Add "病人：" & NVL(rsTmp!姓名) & " 性别：" & NVL(rsTmp!性别) & " 年龄：" & NVL(rsTmp!年龄)
    objRow.Add "住院号：" & NVL(rsTmp!住院号) & " 床号：" & NVL(rsTmp!床号)
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "入院日期：" & Format(NVL(rsTmp!入院日期), "yyyy-MM-dd HH:mm")
    objRow.Add "出院日期：" & Format(NVL(rsTmp!出院日期), "yyyy-MM-dd HH:mm")
    objOut.UnderAppRows.Add objRow
    
    '表下
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & UserInfo.姓名
    objRow.Add "打印日期：" & Format(zlDatabase.Currentdate(), "yyyy年MM月dd日")
    objOut.BelowAppRows.Add objRow
    
    '表体
    Set objOut.Body = vsAdvice
    
    '输出
    vsAdvice.Redraw = False
    lngRow = vsAdvice.Row: lngCol = vsAdvice.Col
    
    strWidth = ""
    For i = 0 To vsAdvice.FixedCols - 1
        strWidth = strWidth & "," & vsAdvice.ColWidth(i)
        vsAdvice.ColWidth(i) = 0
    Next
        
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    strWidth = Mid(strWidth, 2)
    For i = 0 To vsAdvice.FixedCols - 1
        vsAdvice.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    
    vsAdvice.Row = lngRow: vsAdvice.Col = lngCol
    vsAdvice.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim strTab As String, i As Integer
    Dim intType As Integer
        
    mblnFirst = False
    Set mrsPlugInBar = Nothing
    mlngPromptRow = 0
    mlngFontSize = 9
    If Not grsSkinTest Is Nothing Then
        grsSkinTest.Close
        Set grsSkinTest = Nothing
    End If
    
    '医嘱清单
    '-----------------------------------------------------
    Call InitAdviceTable
    Call InitColumnSelect '初始化列选择器
    
    'CommandBars
    '-----------------------------------------------------
    Call GetFilterSetting '本地过滤参数
    Call InitExecBar
    Call InitFilterBar
    
    'TabControl
    '-----------------------------------------------------
    With tbcMain
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        .InsertItem(0, "长嘱临嘱", picMain.hwnd, 0).Tag = "长嘱和临嘱"
        .InsertItem(1, " 长  嘱 ", picMain.hwnd, 0).Tag = "长嘱"
        .InsertItem(2, " 临  嘱 ", picMain.hwnd, 0).Tag = "临嘱"
        .InsertItem(3, " 报  告 ", picMain.hwnd, 0).Tag = "报告"
    End With
    tbcMain.Item(tbcMain.ItemCount - 1).Selected = True
    tbcMain.Item(mvarCond.过滤模式).Selected = True
    If mvarCond.过滤模式 = 3 Then mbln报告 = True
    
    With tbcAppend
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
        End With
        .InsertItem(0, "医嘱计价内容", picAppend.hwnd, 0).Tag = "计价"
        .InsertItem(1, "医嘱发送记录", picAppend.hwnd, 0).Tag = "发送"
        If Not gobjESign Is Nothing Then  '电子签名记录
            .InsertItem(2, "医嘱签名记录", picAppend.hwnd, 0).Tag = "签名"
        End If
        .InsertItem(3, "申请附项", rtfAppend.hwnd, 0).Tag = "附项"
        .InsertItem(4, "安排情况", rtfInfo.hwnd, 0).Tag = "安排"
        
        If gstr输液配置中心 <> "" Then
            Set mfrmCompoundMedicine = New frmCompoundMedicine
            .InsertItem(4, "输液配药记录", mfrmCompoundMedicine.hwnd, 0).Tag = "配药"
        End If
        .InsertItem(5, "预约信息", rtfSche.hwnd, 0).Tag = "预约" 'RIS预约信息
        .InsertItem(6, "其他信息", rtfOther.hwnd, 0).Tag = "其他"  '抗菌药物审核信息
        If gbln血库系统 = True Then
            If InitObjBlood = True Then
                Set mobjFrmBloodList = gobjPublicBlood.zlGetBloodListInfo
                .InsertItem(7, "血液信息", mobjFrmBloodList.hwnd, 0).Tag = "血液"  '血液配血信息
            End If
        End If
        '因为绑定相同,最后要切换回第1个;无数据不影响速度
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    mblnAppend = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "AppendData", 1)) <> 0
    tbcAppend.Visible = mblnAppend: fraAdviceUD.Visible = mblnAppend
    If mblnAppend Then
        strTab = zlDatabase.GetPara("医嘱子列表", glngSys, p住院医嘱下达, "")
        If strTab <> "" Then
            For i = 0 To tbcAppend.ItemCount - 1
                If tbcAppend(i).Visible And tbcAppend(i).Tag = strTab Then
                    tbcAppend.Item(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
    
    '待发血液布局
    If gbln血库系统 = True Then
        With DkpBlood
            .Options.UseSplitterTracker = False '实时拖动
            .Options.ThemedFloatingFrames = True
            .Options.AlphaDockingContext = True
            .Options.HideClient = True
            
            Set objPane = .CreatePane(1, 100, 100, DockLeftOf, Nothing)
            objPane.Title = "输血执行登记"
            objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
        End With
    End If
    
    '恢复个性化设置
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    vsAdvice.ColWidth(COL_F标志) = 11 * Screen.TwipsPerPixelX
    vsAdvice.ColWidth(COL_F报告) = 11 * Screen.TwipsPerPixelX
    
    '变量初始化
    '-----------------------------------------------------
    mstr部门IDs = ""
    mMainPrivs = gMainPrivs '主界面模块权限
    Set mfrmEdit = Nothing
    ReDim marrRollList(0)
    Set mobjReport = New clsReport
    Set mrsDefine = InitAdviceDefine
    
    Call GetLocalSetting
    mblnAutoRead = Val(zlDatabase.GetPara("自动标记报告查阅状态", glngSys, p住院医嘱下达, "1", , , intType)) = 1
    mbln叮嘱发送执行 = Val(zlDatabase.GetPara("叮嘱需要发送执行", glngSys)) = 1
    '医嘱打印模式
    mlngPrintType = Val(zlDatabase.GetPara("医嘱单打印模式", glngSys, p住院医嘱下达))
    '转科出院打印
    mlngPrintPos = Val(zlDatabase.GetPara("转科和出院打印", glngSys, p住院医嘱发送, 1))
    
    mstr检查入院诊断 = zlDatabase.GetPara("要求输入入院诊断", glngSys, p住院医嘱下达)
    
    mblnAutoReadEnabled = Not ((intType = 3 Or intType = 15))
    mblnHaveAuditPriv = HaveAuditPriv
        
    If gblnKSSStrict Then Call CheckKSSPrivilege(1)
    If mint场合 = 0 Then Call InitObjLis(p住院医生站)
    On Error Resume Next
    Set gobjExchange = CreateObject("zlExchange.clsExchange")
    If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
    err.Clear: On Error GoTo 0
End Sub

Private Sub GetLocalSetting()
'功能：读取本地参数
    '执行天数
    mbln天数 = Val(zlDatabase.GetPara("医嘱执行天数", glngSys, p住院医嘱下达)) <> 0
    '皮试验证
    mbln皮试验证 = Val(zlDatabase.GetPara("皮试验证身份", glngSys, p住院医嘱发送)) <> 0
    '申请单打印模式
    mint申请单打印模式 = Val(zlDatabase.GetPara("输血申请单打印模式", glngSys, p住院医嘱发送, "1"))

    mbln医嘱定位最后 = Val(zlDatabase.GetPara("医嘱光标默认定位到最后一行", glngSys, p住院医嘱下达)) = 1
    
    mbln危急值 = InStr(GetInsidePrivs(p住院医生站), ";危急值处理;") > 0
    
    mbln护士签名 = Val(zlDatabase.GetPara("校对医嘱电子签名", glngSys, p住院医嘱发送)) <> 0 And gintCA <> 0 And Mid(gstrESign, 2, 1) = "1"
    
End Sub

Private Sub InitFilterBar()
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsSub.VisualTheme = xtpThemeOffice2003
    With Me.cbsSub.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    Set cbsSub.Icons = zlCommFun.GetPubIcons
    cbsSub.EnableCustomization False
    cbsSub.ActiveMenuBar.Visible = False
    
    Call InitFilterAddBar
End Sub

Private Sub InitFilterAddBar()
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom

    Set objBar = cbsSub.Add("内部工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 24, 24
    objBar.Visible = False  '只有内部调用时才显示(zlDefCommandBars)
    

    Set objBar = cbsSub.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 16, 16
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, ID_时间标签, "时间")   '医嘱时间
        Set objCustom = .Add(xtpControlCustom, ID_时间, "时间")
            objCustom.Handle = cboTime.hwnd
        Set objControl = .Add(xtpControlButton, ID_在用医嘱, "在用医嘱")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_未到终止时间, "未到终止时间")
            objControl.ToolTipText = "显示未到执行终止时间的长期医嘱"
        Set objControl = .Add(xtpControlButton, ID_所有医嘱, "所有医嘱")
        
        '----------------报告页面
        Set objControl = .Add(xtpControlButton, ID_全部, "全部")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_检查, "检查")
            objControl.IconId = 1 '初始时不置图标
        Set objControl = .Add(xtpControlButton, ID_检验, "检验")
        Set objControl = .Add(xtpControlButton, ID_其他, "其他")
            objControl.IconId = 1
        '----------------
        
        Set objPopup = .Add(xtpControlPopup, ID_婴儿, "病人医嘱")
            objPopup.ID = ID_婴儿: objPopup.BeginGroup = True
            objPopup.IconId = 2608
        Set objControl = .Add(xtpControlButton, ID_重整, "重整后")
            objControl.ToolTipText = "显示最近一次重整后的医嘱"
        Set objControl = .Add(xtpControlButton, ID_未记帐, "未记帐")
            objControl.BeginGroup = True
            objControl.ToolTipText = "仅显示包含尚未记帐的划价费用的医嘱"
        Set objControl = .Add(xtpControlButton, ID_科内, "本科下达")
            objControl.ToolTipText = "只显示医技本科下达的医嘱"
            
        Set objControl = .Add(xtpControlButton, ID_是报告医嘱, "需要报告")
            objControl.ToolTipText = "显示需要填写报告的医嘱,和不需要报告两个选项至少选择一个。"
            objControl.IconId = 11
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, ID_非报告医嘱, "不需要报告")
            objControl.ToolTipText = "显示不需要填写报告的医嘱,和需要报告两个选项至少选择一个。"
            objControl.IconId = 11
            
        Set objControl = .Add(xtpControlButton, ID_未出报告, "未出报告")
            objControl.ToolTipText = "显示未出报告"
            objControl.BeginGroup = True
            mvarCond.未出报告 = True
            
        Set objControl = .Add(xtpControlButton, ID_已出报告, "已出报告")
            objControl.ToolTipText = "显示已出报告"
            mvarCond.已出报告 = True
        Set objControl = .Add(xtpControlButton, ID_医嘱颜色示例, "医嘱颜色示例")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButton, ID_简洁, "简洁")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButton, ID_详细, "详细")
            objControl.Flags = xtpFlagRightAlign
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    objBar.Visible = Not mblnHideFilter
    fraHide.Visible = mblnHideFilter
    fraHide.BackColor = cbsSub.GetSpecialColor(STDCOLOR_BTNFACE)
    
    '缺省医嘱时间
    cboTime.Clear
    cboTime.AddItem "所有"
    cboTime.AddItem "今天"
    cboTime.AddItem "昨天"
    cboTime.AddItem "最近三天"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近两周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "[指定..]"
    Call zlControl.CboSetIndex(cboTime.hwnd, 0)
    mintPreTime = 0
    cboTime.Visible = Not mblnHideFilter
End Sub

Private Sub InitExecBar()
    Dim objBar As CommandBar
    Dim objControl As CommandBarControl

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsExec.VisualTheme = xtpThemeOfficeXP
    With Me.cbsExec.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .UseFadedIcons = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    Set cbsExec.Icons = zlCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("工具栏", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_显示执行, "显示执行内容")
        Set objControl = .Add(xtpControlButton, ID_完成执行, "执行完成")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_Complete
        Set objControl = .Add(xtpControlButton, ID_取消完成, "取消完成")
            objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, ID_执行记录, "记录执行情况")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_ThingAdd
        Set objControl = .Add(xtpControlButton, ID_执行调整, "调整执行情况")
            objControl.IconId = conMenu_Manage_ThingModi
        Set objControl = .Add(xtpControlButton, ID_执行删除, "删除执行情况")
            objControl.IconId = conMenu_Manage_ThingDel
        Set objControl = .Add(xtpControlButton, ID_核对, "核对")
            objControl.IconId = conMenu_Manage_ThingAudit
        Set objControl = .Add(xtpControlButton, ID_取消核对, "取消核对")
            objControl.IconId = conMenu_Manage_ThingDelAudit
    End With
    For Each objControl In objBar.Controls
        If objControl.Type <> xtpControlLabel Then
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbsExec.KeyBindings
        '.Add FCONTROL, vbKeyH, 0
    End With
End Sub

Private Sub mfrmParent_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：捕获主窗体的按键,用于处理医嘱过滤热键
'说明：
'1.当医嘱子窗体未激活时,子窗体CommandBar的热键无效
'2.主窗体CommandBar或KeyDown事件处理了的键不会再激活该事件
    
    If Not Me.Visible Then Exit Sub '在其他子窗体时仍会激活
    If mlng病人ID = 0 Then Exit Sub
    Call ActiveHotKey(KeyCode, Shift)
End Sub

Private Sub ActiveHotKey(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim lngID As Long
    Dim intTab As Integer
    
    If Not Me.Visible Then Exit Sub
    If mlng病人ID = 0 Then Exit Sub
    intTab = -1
    
    If Shift = vbCtrlMask And KeyCode >= vbKey0 And KeyCode <= vbKey5 Then
        lngID = ID_婴儿 * 100# + KeyCode - vbKey0 + 1
    ElseIf Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKey6
                intTab = 0
            Case vbKey7
                intTab = 1
            Case vbKey8
                intTab = 2
            Case vbKey9
                intTab = 3
            Case vbKeyB
                lngID = ID_婴儿 * 100#
            Case vbKeyJ
                lngID = ID_重整
            Case vbKeyK
                lngID = ID_科内
            Case vbKeyU
                If mvarCond.过滤模式 = 3 Then
                    lngID = ID_全部
                Else
                    lngID = ID_在用医嘱
                End If
            Case vbKeyX
                lngID = ID_检查
            Case vbKeyY
                lngID = ID_检验
            Case vbKeyQ
                lngID = ID_其他
        End Select
    ElseIf KeyCode = vbKeyEscape Then '关闭列选择器
        If vsColumn.Visible Then
            vsColumn.Visible = False
            If vsAdvice.Enabled Then vsAdvice.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then '打开列选择器
        Call imgColSel_MouseUp(1, 0, 0, 0)
    ElseIf KeyCode = vbKeyF8 Then
        Call RefreshData
    End If
    If lngID <> 0 Then
        Set objControl = cbsSub.FindControl(, lngID, , True)
        If Not objControl Is Nothing Then objControl.Execute
    End If
    If intTab <> -1 Then tbcMain.Item(intTab).Selected = True
End Sub

Private Sub Form_Resize()
    If WindowState = 1 Then Exit Sub
    With Me.tbcMain
        .Left = 0
        .Top = 0
        .Height = Me.Height
        .Width = Me.Width
    End With
    
End Sub

Private Sub fraAdviceUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsAdvice.Height + Y < 1000 Or tbcAppend.Height - Y < 500 Then Exit Sub
        fraAdviceUD.Top = fraAdviceUD.Top + Y
        vsAdvice.Height = vsAdvice.Height + Y
        tbcAppend.Top = tbcAppend.Top + Y
        tbcAppend.Height = tbcAppend.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not mfrmEdit Is Nothing Then Unload mfrmEdit: Set mfrmEdit = Nothing
    If Not mfrmEac Is Nothing Then Unload mfrmEac: Set mfrmEac = Nothing
    Set mobjReport = Nothing
    Set mcbsMain = Nothing
    Set mrsPlugInBar = Nothing
    If Not mfrmCompoundMedicine Is Nothing Then
        Unload mfrmCompoundMedicine
    End If
    Set mfrmCompoundMedicine = Nothing
    Set gobjExchange = Nothing
    Set gobjLIS = Nothing
    Set mobjPublicPACS = Nothing
    Set gobjRecipeAudit = Nothing
    Set gobjPublicBlood = Nothing
    
    If Not mobjFrmBlood Is Nothing Then
        Unload mobjFrmBlood
        Set mobjFrmBlood = Nothing
    End If
    If Not mobjFrmBloodList Is Nothing Then
        Unload mobjFrmBloodList
        Set mobjFrmBloodList = Nothing
    End If
    Set mrsDefine = Nothing
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    If Not gobjEmrInterface Is Nothing Then
        Set gobjEmrInterface = Nothing
    End If
    'PASS
    If mblnPass Then
        Call gobjPass.zlPassClearLight(mobjPassMap, 1)
    End If
    mblnPass = False
    Set mobjPassMap = Nothing
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name, "AppendData", IIF(mblnAppend, 1, 0)
    If mblnAppend And Not tbcAppend.Selected Is Nothing Then
        Call zlDatabase.SetPara("医嘱子列表", tbcAppend.Selected.Tag, glngSys, p住院医嘱下达)
    End If
    Call SaveFilterSetting
    Call SaveWinState(Me, App.ProductName)
    
    '外挂程序对象终止
    Call CreatePlugInOK(IIF(mint场合 = 1, p住院医嘱发送, p住院医嘱下达), mint场合)
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If mint场合 = 0 Then '医生站调用
            Call gobjPlugIn.Terminate(glngSys, p住院医嘱下达, 0)
        ElseIf mint场合 = 1 Then '护士站调用
            Call gobjPlugIn.Terminate(glngSys, p住院医嘱发送, 1)
        ElseIf mint场合 = 2 Then '医技站调用
            Call gobjPlugIn.Terminate(glngSys, p住院医嘱下达, 2)
        End If
        Call zlPlugInErrH(err, "Terminate")
        err.Clear: On Error GoTo 0
    End If
    Set mclsMipModule = Nothing
    Set mrs危急值 = Nothing
    mbln危急值 = False
    mlng危急值ID = 0
End Sub

Private Sub RefreshData()
'功能：刷新数据
    If mlng病人ID = 0 Then
        '清除医嘱清单
        Call ClearAdviceData
        Call ClearAppendData
        mlngBabyDept = 0
    Else
        '显示医嘱清单
        Call LoadAdvice
    End If
End Sub

Private Sub Refresh报告()
'功能：在报告页面不同报告之间切换时界面的刷新，不重新读数据库设置表格的隐藏和显示
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lng医嘱ID As Long
    Dim strFormat As String
    Dim strSameDay As String
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))      '记录当前行如果是在当前界面刷新医嘱行应该不变
        
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_报告项)) <> 0 Then
                If mvarCond.报告 = 0 Then ' 全部
                    blnTmp = True
                ElseIf mvarCond.报告 = 1 Then ' 检查
                    blnTmp = .TextMatrix(i, COL_诊疗类别) = "D"
                ElseIf mvarCond.报告 = 2 Then '检验
                    blnTmp = (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "C")
                ElseIf mvarCond.报告 = 3 Then ' 其它
                    blnTmp = Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "D" Or .TextMatrix(i, COL_诊疗类别) = "C")
                End If
                
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                
                .RowHidden(i) = Not blnTmp
            Else
                .RowHidden(i) = True: .RowHeight(i) = 0
            End If
            
            '增加过滤未出的报告和已出的报告
            If .RowHidden(i) = False Then
                blnTmp = IIF(.TextMatrix(i, COL_查阅状态) = "未出", mvarCond.未出报告, mvarCond.已出报告)
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                .RowHidden(i) = Not blnTmp
            End If
        Next
    End With
    Call LocatedDefaultAdviceRow(lng医嘱ID)
End Sub

Private Sub LocatedDefaultAdviceRow(Optional ByVal lng医嘱ID As Long)
'功能：医嘱清单的缺省定位，如果有医嘱id跟据医嘱id定位
    '缺省定位，当前选择的医嘱为显示行则定位，否则定位到最后一行。
    Dim i As Long
    
    With vsAdvice
        .Redraw = flexRDNone
        .Row = .Rows - 1
        If lng医嘱ID <> 0 Then
            lng医嘱ID = .FindRow(CStr(lng医嘱ID), , COL_ID)
            If lng医嘱ID <> -1 Then
                If Not .RowHidden(lng医嘱ID) Then .Row = lng医嘱ID
            End If
        End If
        If mint场合 = 1 Then
            If lng医嘱ID = -1 Or lng医嘱ID = 0 Then
                vsAdvice.Row = IIF(mbln医嘱定位最后, vsAdvice.Rows - 1, vsAdvice.FixedRows)
            End If
        End If
        If .RowHidden(.Row) Then    '定位到了隐藏行的处理
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            For i = .Row - 1 To .FixedRows Step -1
                If Not .RowHidden(i) Then .Row = i: Exit For
            Next
        End If
        If .RowHidden(.Row) Then
            .AddItem "": .Row = .Rows - 1
        End If
        .Col = .FixedCols
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
        .Refresh
    End With
End Sub

Private Sub GetFilterSetting()
'功能：读取医嘱过滤设置条件
    Dim strPar As String
    
    mvarCond.婴儿 = 0
    mvarCond.未记帐 = False
    mblnHideFilter = Val(zlDatabase.GetPara("过滤条件自动隐藏", glngSys, p住院医嘱下达, "0")) <> 0
    mvarCond.重整 = Val(zlDatabase.GetPara("重整医嘱过滤", glngSys, p住院医嘱下达, "0")) <> 0
    mvarCond.科内 = Val(zlDatabase.GetPara("科内医嘱过滤", glngSys, p住院医嘱下达, "1")) <> 0
    
    strPar = Val(zlDatabase.GetPara("显示模式", glngSys, p住院医嘱下达, "0"))
    mvarCond.显示模式 = IIF(Val(strPar) = 0, 0, 1)
    
    mlngBaby = Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0")) - 1
    
    strPar = Val(zlDatabase.GetPara("医嘱过滤方式", glngSys, p住院医嘱下达, "0"))
    If InStr(",0,1,2,3,", "," & strPar & ",") > 0 Then
        mvarCond.过滤模式 = Val(strPar)
    Else
        mvarCond.过滤模式 = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("报告查看类型", glngSys, p住院医嘱下达, "0"))
    If InStr(",0,1,2,3,", "," & strPar & ",") > 0 Then
        mvarCond.报告 = Val(strPar)
    Else
        mvarCond.报告 = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("医嘱显示在用", glngSys, p住院医嘱下达, "0"))
    mvarCond.医嘱显示 = IIF(Val(strPar) = 0, 0, 1)
    
    mvarCond.未到终止时间 = Val(zlDatabase.GetPara("医嘱显示在用未到终止时间", glngSys, p住院医嘱下达, "1")) = 1
    
    strPar = Val(zlDatabase.GetPara("医嘱显示报告需要", glngSys, p住院医嘱下达, "0"))
    If strPar = "1" Then
        mvarCond.是报告医嘱 = True: mvarCond.非报告医嘱 = False
    ElseIf strPar = "2" Then
        mvarCond.是报告医嘱 = False: mvarCond.非报告医嘱 = True
    Else
        mvarCond.是报告医嘱 = True: mvarCond.非报告医嘱 = True
    End If
End Sub

Private Sub SaveFilterSetting()
'功能：保存医嘱过滤设置条件
    Dim strPar As String
    
    Call zlDatabase.SetPara("重整医嘱过滤", IIF(mvarCond.重整, 1, 0), glngSys, p住院医嘱下达)
    Call zlDatabase.SetPara("科内医嘱过滤", IIF(mvarCond.科内, 1, 0), glngSys, p住院医嘱下达)
    Call zlDatabase.SetPara("显示模式", mvarCond.显示模式, glngSys, p住院医嘱下达)
    Call zlDatabase.SetPara("医嘱过滤方式", mvarCond.过滤模式, glngSys, p住院医嘱下达)
    Call zlDatabase.SetPara("报告查看类型", mvarCond.报告, glngSys, p住院医嘱下达)
    Call zlDatabase.SetPara("过滤条件自动隐藏", IIF(mblnHideFilter, 1, 0), glngSys, p住院医嘱下达)
    Call zlDatabase.SetPara("医嘱显示在用", mvarCond.医嘱显示, glngSys, p住院医嘱下达)
    Call zlDatabase.SetPara("医嘱显示在用未到终止时间", IIF(mvarCond.未到终止时间, 1, 0), glngSys, p住院医嘱下达)
    
    If mvarCond.是报告医嘱 And Not mvarCond.非报告医嘱 Then
        strPar = "1"
    ElseIf Not mvarCond.是报告医嘱 And mvarCond.非报告医嘱 Then
        strPar = "2"
    Else
        strPar = "0"
    End If
    Call zlDatabase.SetPara("医嘱显示报告需要", strPar, glngSys, p住院医嘱下达)
End Sub

Private Sub ClearAppendData()
'功能：清除附加表格和申请附项的数据
    Dim blnSel As Boolean, intIdx As Integer
    Dim varDraw As RedrawSettings
    
    If vsAppend.FixedRows = 2 Then vsAppend.RemoveItem 0
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
    vsAppend.Row = vsAppend.FixedRows
    
    If rtfAppend.Visible Then rtfAppend.Text = ""
    If rtfInfo.Visible Then rtfInfo.Text = ""
    
    For intIdx = 0 To tbcAppend.ItemCount - 1
        If InStr("附项,安排,配药,预约,其他,血液", tbcAppend(intIdx).Tag) > 0 Then
            If tbcAppend(intIdx).Selected Then blnSel = True
            tbcAppend(intIdx).Visible = False
        End If
    Next
   
    If blnSel Then
        varDraw = vsAdvice.Redraw '根据条件屏蔽重复调用
        vsAdvice.Redraw = flexRDNone
        tbcAppend.Item(0).Selected = True
        vsAdvice.Redraw = varDraw
    End If
    
    ReDim marrRollList(0)
End Sub

Private Sub InitPriceTable()
'功能：初始化计价清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "计价医嘱,2000,1;类别,650,1;收费项目,2500,1;单位,500,4;计价数量,850,1;单价,900,7;执行科室,1000,1;费用类型,800,1;从项,450,4;收费方式,1500,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLPrice.Count <> UBound(arrHead) + 1 Then COLPrice.Add i, Split(arrHead(i), ",")(0)
            .MergeCol(i) = False
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeRestrictAll
        .MergeCompare = flexMCIncludeNulls
    End With
End Sub

Private Sub InitSendTable()
'功能：初始化发送清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "发送号;发送时间,1530,1;单据号,850,1;发送医嘱,1800,1;收费项目,1800,1;发送数次,850,1;计费状态,850,1;" & _
        "执行状态,850,1;状态说明,1800,1;执行科室,1000,1;执行人,800,1;执行时间,1530,1;最后执行时间,1530,1;执行说明,1800,1;首次时间,1530,1;末次时间,1530,1;发送人,800,1;医嘱ID;相关ID;记录性质;门诊记帐;记录状态;诊疗类别;操作类型;跟踪在用;完成时间;输血类型;接受时间,1530,1;接受人,800,1;到场时间,1530,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1  '隐式调用了vsAppend_AfterRowColChange
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            
            If COLSend.Count <> UBound(arrHead) + 1 Then COLSend.Add i, Split(arrHead(i), ",")(0)
            .MergeCol(i) = False
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .Redraw = flexRDDirect
        
        .MergeCells = flexMergeRestrictAll  '自动设置MergeCellsFixed为相同格式
        .MergeCompare = flexMCIncludeNulls
    End With
End Sub

Private Sub InitExecTable()
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "要求时间,1530,1;执行时间,1530,1;本次数次,850,1;执行摘要,2500,1;执行人,750,1;登记时间,1530,1;登记人,750,1;执行结果,1000,1;核对人,750,1;核对时间,1530,1;说明,500,1;来源,600,1"
    arrHead = Split(strHead, ";")
    With vsExec
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLExec.Count <> UBound(arrHead) + 1 Then COLExec.Add i, Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeNever
    End With
End Sub

Private Sub InitSignTable()
'功能：初始化签名清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long, blnDo As Boolean
    
    strHead = "签名类型,1150,1;签名时间,1900,1;签名人,800,1;时间戳,1900,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
            If COLSign.Count <> UBound(arrHead) + 1 Then COLSign.Add i, Split(arrHead(i), ",")(0)
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .MergeCells = flexMergeNever
    End With
End Sub

Private Sub ClearAdviceData()
'功能：清除医嘱清单数据
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitColumnSelect()
'功能：根据医嘱清单原始列显示状态初始化列选择器
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vsAdvice
        For i = .FixedCols To .Cols - 1
            If Not (.ColHidden(i) Or .ColWidth(i) = 0) Then
                If .TextMatrix(0, i) <> "" And Not (i = COL_查阅状态 Or i = COL_标本状态) Then  '审查结果,皮试
                    vsColumn.Rows = vsColumn.Rows + 1
                    lngRow = vsColumn.Rows - 1
                    vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                    vsColumn.RowData(lngRow) = i
                    
                    '固定显示列
                    If InStr(",开始时间,医嘱内容,开嘱医生,", "," & .TextMatrix(0, i) & ",") > 0 Then
                        vsColumn.TextMatrix(lngRow, 0) = 1
                        vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                    End If
                    
                    '默认隐藏开嘱时间
                     If InStr(",开嘱时间,", "," & .TextMatrix(0, i) & ",") > 0 Then
                        vsAdvice.ColWidth(i) = 0
                        vsAdvice.ColHidden(i) = True
                        vsColumn.TextMatrix(lngRow, 0) = 0
                    End If
                End If
            End If
        Next
    End With
    If vsColumn.Rows > 1 Then vsColumn.Row = 1
End Sub

Private Sub InitAdviceTable()
'功能：初始化医嘱清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "ID;相关ID;序号;婴儿ID;医嘱状态;诊疗类别;操作类型;毒理分类;标志;" & _
              ",240,4;期效,500,4;生效时间,1530,1;,200,7;医嘱内容,3000,1;内容,4000,1;,375,1;总量,850,1;单量,850,1;天数,450,1;频率,1000,1;用法,1000,1;医生嘱托,1000,1;执行时间,1000,1;" & _
              "终止时间,1530,1;执行科室,1000,1;执行性质,850,1;上次执行,1560,1;状态,500,4;开嘱医生,850,1;开嘱时间,1530,1;校对护士,850,1;校对时间,1530,1;停嘱医生,850,1;" & _
              "停嘱时间,1530,1;停嘱护士,850,1;确认停嘱时间,1530,1;基本药物,850,1;查阅状态,700,4;标本状态,850,1;" & _
              "诊疗项目ID;试管编码;执行标记;屏蔽打印;前提ID;签名否;文件ID;报告项;报告ID;收费细目ID;单量单位;开嘱科室ID;审核状态;申请序号;" & _
              "审核标记;高危药品;标本部位;用药目的;检查报告ID;处方审查状态;处方审查结果;RIS预约ID;RIS报告ID;LIS报告ID;RIS预约状态;诊疗项目名称;检查方法;危急值ID;易跌倒"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 2
        .Cols = .FixedCols + UBound(arrHead) + 1

        .Rows = .FixedRows + 1

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                If mlngFontSize <> 9 Then
                    lngWidth = lngWidth * mlngFontSize / 9
                    If lngWidth > .ColWidthMax And .ColWidthMax <> 0 Then lngWidth = .ColWidthMax
                End If
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        '未启用合理用药时，该列不可见，启用美康，太元通时，即当gbytPass=1 or 3 时 可见
        .ColHidden(COL_警示) = True
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(COL_F标志) = 11 * Screen.TwipsPerPixelX
        .ColWidth(COL_F报告) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Sub SetAdviceColVisible()
'功能：设置医嘱表格列的可见性和表头列名
    Dim i As Long
    
    '根据显示模式调整显示列
    With vsAdvice
        If (mvarCond.过滤模式 = 1 Or mvarCond.过滤模式 = 2) And mvarCond.显示模式 = 0 Then
            .ColHidden(COL_期效) = True
        Else
            .ColHidden(COL_期效) = False
        End If
        
        .ColHidden(col_医嘱内容) = mvarCond.显示模式 = 0
        .ColHidden(col_内容) = mvarCond.显示模式 = 1
        .ColHidden(COL_皮试) = False
        .ColHidden(COL_总量) = mvarCond.显示模式 = 0
        .ColHidden(COL_单量) = mvarCond.显示模式 = 0
        .ColHidden(COL_天数) = mvarCond.显示模式 = 0
        .ColHidden(COL_频率) = mvarCond.显示模式 = 0
        .ColHidden(COL_执行时间) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_执行时间) = "Detail"
        .ColHidden(COL_执行性质) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_执行性质) = "Detail"
        .ColHidden(COL_上次执行) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_上次执行) = "Detail"
        .ColHidden(COL_状态) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_状态) = "Detail"
        .ColHidden(COL_开嘱时间) = True
        .ColHidden(COL_校对护士) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_校对护士) = "Detail"
        .ColHidden(COL_校对时间) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_校对时间) = "Detail"
        .ColHidden(COL_停嘱医生) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_停嘱医生) = "Detail"
        .ColHidden(COL_停嘱时间) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_停嘱时间) = "Detail"
        .ColHidden(COL_停嘱护士) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_停嘱护士) = "Detail"
        .ColHidden(COL_确认停嘱时间) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_确认停嘱时间) = "Detail"
        .ColHidden(COL_基本药物) = mvarCond.显示模式 = 0: .Cell(flexcpData, 0, COL_基本药物) = "Detail"
        .ColHidden(COL_高危药品) = True
        .ColHidden(COL_标本部位) = True
        .ColHidden(COL_审核标记) = True
        .ColHidden(COL_用药目的) = True
        .ColHidden(COL_检查报告ID) = True
        .ColHidden(COL_处方审查状态) = True
        .ColHidden(COL_处方审查结果) = True
        .ColHidden(COL_并) = True
        .ColHidden(COL_标本状态) = True
        
        If mvarCond.过滤模式 = 3 Then '如是报告卡片先藏再显示
            For i = COL_开始时间 + 1 To COL_标本部位
                .ColHidden(i) = True
            Next
            .ColHidden(COL_期效) = True
            .ColHidden(COL_开始时间) = False
            .ColHidden(col_内容) = False
            .ColHidden(COL_执行科室) = False
            .ColHidden(COL_开嘱医生) = False
            .TextMatrix(0, COL_开嘱医生) = "申请医生"
            .ColHidden(COL_查阅状态) = mfrmParent Is Nothing    '电子病案查阅未传入主窗体,禁止显示查阅状态
            .ColWidth(COL_查阅状态) = 700
            .TextMatrix(0, COL_查阅状态) = "报告"
            .ColHidden(COL_标本状态) = False
            .ColWidth(COL_标本状态) = 850
        Else
            .ColHidden(COL_并) = False
            .TextMatrix(0, COL_开嘱医生) = "开嘱医生"
            If mvarCond.过滤模式 = 0 And mvarCond.显示模式 = 0 Then .ColHidden(COL_期效) = False
            .ColHidden(COL_用法) = False
            .ColHidden(COL_医生嘱托) = False
            .ColHidden(COL_终止时间) = (mvarCond.显示模式 = 0 And mvarCond.过滤模式 = 2)
            .ColHidden(COL_查阅状态) = True
            .TextMatrix(0, COL_查阅状态) = "查阅状态"
        End If
        '只有长嘱时隐藏天数列
        If mvarCond.显示模式 = 1 Then .ColHidden(COL_天数) = mvarCond.过滤模式 = 1 Or Not mbln天数
    End With
End Sub

Private Function LoadAdvice(Optional ByVal blnRefreshNotify As Boolean) As Boolean
'功能：根据当前界面设置读取并显示医嘱清单
'参数：blnRefreshNotify刷新医嘱提醒(F5手动刷新,新开医嘱，停止医嘱，作废医嘱时)
    Dim rsTmp As ADODB.Recordset
    Dim rs血型 As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim i As Long, j As Long
    Dim strFormat As String, strTmp As String
    Dim bln给药途径 As Boolean, bln中药用法 As Boolean
    Dim bln采集方法 As Boolean, bln输血途径 As Boolean, blnFirst As Boolean
    Dim str状态SQL As String, lng医嘱ID As Long
    Dim str未记帐 As String
    Dim blnDo As Boolean, strCurr As String, strTime As String
    Dim str医嘱期效 As String, str医嘱状态 As String
    Dim strSameDay As String, strGroupBy As String
    Dim strPreDay1 As String, strPreDay2 As String
    If mlng病人ID = 0 Then Exit Function

    Screen.MousePointer = 11

    On Error GoTo errH
    
    lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))    '记录当前行如果是在当前界面刷新医嘱行应该不变
    If mvarCond.医嘱ID <> 0 And lng医嘱ID = 0 Then
        lng医嘱ID = mvarCond.医嘱ID
    End If

    '医嘱过滤条件
    If mstr婴儿 <> "" Then
        If mblnFirstBaby = False Then
            mvarCond.婴儿 = mlngBaby
            mbyt婴儿 = IIF(mvarCond.婴儿 = -1, 0, mvarCond.婴儿)
            Call zlDatabase.SetPara("病人婴儿过滤", mvarCond.婴儿, glngSys, p住院医嘱下达)
            mblnFirstBaby = True
        End If

        '母婴分离的处理
        If mlngBabyDept <> mlng婴儿科室ID Then
            If mlng婴儿科室ID <> 0 Then
                If (mvarCond.婴儿 = -1 Or mvarCond.婴儿 = 0) And (mlng婴儿科室ID = mlng医护科室ID Or mlng婴儿病区ID = mlng医护科室ID) Then
                    '婴儿科室病区默认选中婴儿
                    mvarCond.婴儿 = 1: mbyt婴儿 = mvarCond.婴儿
                ElseIf (mvarCond.婴儿 = -1 Or mvarCond.婴儿 = 1) And (mlng科室ID = mlng医护科室ID Or mlng病区ID = mlng医护科室ID) Then
                    '病人科室病区默认选择病人
                    mvarCond.婴儿 = 0: mbyt婴儿 = mvarCond.婴儿
                End If
            End If
            mlngBabyDept = mlng婴儿科室ID
        End If
    Else
        mlngBabyDept = 0
        mblnFirstBaby = False
    End If
    strWhere = ""
    If mvarCond.婴儿 <> -1 Then
        strWhere = strWhere & " And Nvl(A.婴儿,0)=[4]"
    End If
        
    If mvarCond.过滤模式 = 1 Then
        strWhere = strWhere & " And A.医嘱期效=0"
    ElseIf mvarCond.过滤模式 = 2 Then
        strWhere = strWhere & " And A.医嘱期效=1"
    End If
    
    If mvarCond.医嘱显示 = 0 And mvarCond.过滤模式 <> 3 Then  '在用医嘱，如果是在报告页面不用区分医嘱的范围，直接取出这个病人的所有报告。
        strWhere = strWhere & " And Nvl(A.医嘱状态,0)<>4  And (A.医嘱期效=0 and " & _
            IIF(mvarCond.未到终止时间, " (a.执行终止时间>[3] or a.执行终止时间 is null) ", " a.执行终止时间 is null ") & _
            " or A.医嘱期效=1 and A.开始执行时间 >=[6])"
    End If
    
    If Not (mvarCond.过滤模式 <> 3 And mvarCond.医嘱显示 = 0) Then
        If mvarCond.开始时间 <> CDate(0) And mvarCond.结束时间 <> CDate(0) Then
            strWhere = strWhere & " And A.开嘱时间+0 Between [7] And [8]"
        End If
    End If
    
    '只显示包含未记帐费用的医嘱
    If mvarCond.未记帐 Then
        str未记帐 = _
        " And Exists" & vbNewLine & _
                 " (Select 1" & vbNewLine & _
                 "       From (Select Nvl(C.相关id, C.ID) As 医嘱id" & vbNewLine & _
                 "              From 病人医嘱发送 A, 住院费用记录 B, 病人医嘱记录 C" & vbNewLine & _
                 "              Where A.医嘱id = C.ID And A.NO = B.NO And A.记录性质 = B.记录性质 And A.记录性质 = 2 And B.记录状态 = 0 And" & vbNewLine & _
                 "                    C.病人id = [1] And C.主页id = [2]" & IIF(mvarCond.婴儿 <> -1, " And Nvl(C.婴儿, 0) = [4]", "") & ")" & vbNewLine & _
                 "       Where A.ID = 医嘱id Or A.相关id = 医嘱id)"
    End If
    
    '医技站  本科下达
    If mlng前提ID <> 0 And mvarCond.科内 Then
        strWhere = strWhere & " And Nvl(A.前提ID,0)<>0 and (A.前提ID in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([5]) As zlTools.t_Numlist)) X) or a.开嘱科室ID=[9])"
    End If
    
    '医嘱记录：不含附加手术,手术麻醉,检查部位,中药煎法'总量及用法计算
    str状态SQL = "Decode(A.医嘱状态,1,'新开',2,'疑问',3,'校对',4,'作废',5,'重整',6,'暂停',7,'启用',8,'停止',9,'确认停止')"
    strSQL = _
        " Select /*+ RULE */ A.ID,A.相关ID,A.序号,Nvl(A.婴儿,0) as 婴儿ID,A.医嘱状态,Nvl(A.诊疗类别,'*') as 诊疗类别,B.操作类型,C.毒理分类,A.紧急标志 as 标志,A.审查结果 as 警示," & _
        " Decode(Nvl(A.医嘱期效,0),0,'长嘱','临嘱') as 期效,To_Char(A.开始执行时间,'YYYY-MM-DD HH24:MI') as 开始时间,Null as 并,A.医嘱内容,Null as 内容,A.皮试结果 as 皮试," & _
        " Decode(A.总给予量,NULL,NULL,Decode(A.诊疗类别,'E',Decode(B.操作类型,'4',A.总给予量||'付',A.总给予量||B.计算单位),'4',A.总给予量||G.计算单位,'5',Round(A.总给予量/D.住院包装,5)||D.住院单位,'6',Round(A.总给予量/D.住院包装,5)||D.住院单位,A.总给予量||B.计算单位)) as 总量," & _
        " Decode(A.首次用量,Null,'',A.首次用量||Decode(A.诊疗类别,'4',G.计算单位,B.计算单位)||':')||Decode(A.单次用量,NULL,NULL,decode(sign(1-A.单次用量),1,'0'||A.单次用量,A.单次用量)||Decode(A.诊疗类别,'4',G.计算单位,B.计算单位)) as 单量," & _
        " A.天数,A.执行频次 as 频率,Decode(A.诊疗类别,'E',Decode(Instr('2468',Nvl(B.操作类型,'0')),0,NULL,B.名称),NULL) as 用法,A.医生嘱托,A.执行时间方案 as 执行时间," & _
        " To_Char(A.执行终止时间,'YYYY-MM-DD HH24:MI') as 终止时间,Nvl(E.名称,Decode(Nvl(A.执行性质,0),0,'<叮嘱>',5,'-')) as 执行科室," & _
        " Decode(Instr('567E',Nvl(A.诊疗类别,'*')),0,NULL,A.执行性质) as 执行性质,To_Char(A.上次执行时间,'YYYY-MM-DD HH24:MI') as 上次执行," & str状态SQL & " as 状态," & _
        " A.开嘱医生,To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 开嘱时间,A.校对护士,To_Char(A.校对时间,'YYYY-MM-DD HH24:MI') as 校对时间,A.停嘱医生," & _
        " To_Char(A.停嘱时间,'YYYY-MM-DD HH24:MI') as 停嘱时间,A.确认停嘱护士 as 停嘱护士,To_Char(A.确认停嘱时间,'YYYY-MM-DD HH24:MI') as 确认停嘱时间,D.基本药物,D.是否易至跌倒,Decode(Max(NVL(y.查阅状态,0)),Min(NVL(y.查阅状态,0)),Max(NVL(y.查阅状态,0)),2) As 查阅状态,null as 标本状态,A.诊疗项目ID," & _
        " B.试管编码,A.执行标记,A.屏蔽打印,A.前提ID,Decode(A.新开签名ID,NULL,0,1) as 签名否,M.病历文件ID as 文件ID,Nvl(N.通用,0) as 报告项,Max(y.病历id) As 报告id," & _
        " A.收费细目ID,B.计算单位 as 单量单位,A.开嘱科室ID,A.审核状态,A.申请序号," & _
        " A.审核标记,d.高危药品,A.标本部位,A.用药目的 ,Max(y.检查报告id)||'' As 检查报告id,J.状态 as 处方审查状态,J.审查结果 as 处方审查结果,f.预约ID as RIS预约ID,Max(y.RISID) As RIS报告ID,Max(y.报告ID) as LIS报告ID,f.是否调整 as RIS预约状态,b.名称 as 诊疗项目名称,Max(a.检查方法) as 检查方法,max(h.危急值id) as 危急值ID"
    strSQL = strSQL & _
        " From 病人医嘱记录 A,部门表 E,药品特性 C,药品规格 D,诊疗项目目录 B,收费项目目录 G,病人医嘱报告 Y,病历单据应用 M,病历文件列表 N,处方审查明细 I,处方审查记录 J,RIS检查预约 F,病人危急值医嘱 H" & _
        " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=E.ID(+) And A.诊疗项目ID=C.药名ID(+) And a.ID = i.医嘱ID(+) And I.审方ID = J.ID(+) and (I.最后提交 =1 Or I.审方ID is NULL) and a.id=f.医嘱ID(+) and a.id=h.医嘱ID(+)" & _
        " And A.收费细目ID=D.药品ID(+) And A.收费细目ID=G.ID(+) And A.ID=Y.医嘱ID(+) And (Not(A.诊疗类别 IN ('F','G','D','E') And A.相关ID is Not NULL) Or A.诊疗类别='E' And B.操作类型='8')" & _
        " And A.诊疗项目ID=M.诊疗项目ID(+) And M.应用场合(+)=2 And M.病历文件ID=N.ID(+) And N.种类(+)=7 And A.病人ID=[1] And A.主页ID=[2] And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1" & _
        IIF(mint场合 = 2, "", " And A.病人来源<>3") & strWhere & str未记帐
    strGroupBy = _
        " Group By a.Id,a.相关id,a.序号,a.婴儿,a.医嘱状态,a.诊疗类别,b.操作类型,c.毒理分类,a.紧急标志,a.审查结果,a.医嘱期效,a.开始执行时间,a.医嘱内容,a.皮试结果," & _
        " a.总给予量,a.首次用量,g.计算单位,d.住院包装,d.住院单位,a.单次用量,a.天数,a.执行频次,a.医生嘱托,b.名称,a.执行性质,a.执行时间方案,a.执行终止时间,e.名称,a.上次执行时间," & _
        " a.开嘱时间,a.开嘱医生,a.校对护士,a.校对时间,a.停嘱医生,a.停嘱时间,a.确认停嘱护士,a.确认停嘱时间,a.诊疗项目id,b.试管编码,a.执行标记,a.屏蔽打印,a.前提id,a.新开签名id," & _
        " m.病历文件id,n.通用,a.收费细目id,b.计算单位,a.开嘱科室id,a.审核状态,a.申请序号,a.审核标记,d.基本药物,d.高危药品,a.标本部位,a.用药目的,J.状态,J.审查结果,f.预约ID,f.是否调整,b.名称,D.是否易至跌倒"
    '重整显示格式处理
    If mdat重整 <> CDate("1900-01-01") Then
        If mvarCond.重整 Then
            '只显示最后一次重整之后的医嘱
            strSQL = strSQL & " And (Nvl(A.重整标志,1)=1 Or A.医嘱状态 IN(1,2)) " & strGroupBy & " Order by 婴儿ID,序号"
        Else
            '显示重整前后分隔
            strSQL = _
                " Select * From (" & strSQL & " And Nvl(A.重整标志,1)=0 And A.医嘱状态 Not IN(1,2) " & strGroupBy & " Order by 婴儿ID,序号)" & _
                " Union ALL" & _
                " Select -Null as ID,-Null as 相关ID,-Null as 序号,-Null as 婴儿ID,-Null as 医嘱状态,Null as 诊疗类别,Null as 操作类型,Null as 毒理分类,-Null as 标志,-Null as 警示," & _
                " Null as 期效,Null as 开始时间,Null as 并,Null as 医嘱内容,Null as 内容,Null as 皮试,Null as 总量,Null as 单量,-Null as 天数,Null as 频率,Null as 用法,Null as 医生嘱托,Null as 执行时间," & _
                " Null as 终止时间,Null as 执行科室,Null as 执行性质,Null as 上次执行,Null as 状态,Null as 开嘱医生,Null as 开嘱时间,Null as 校对护士,Null as 校对时间,Null as 停嘱医生," & _
                " Null as 停嘱时间,Null as 停嘱护士,Null as 确认停嘱时间,Null as 基本药物,-Null as 是否易至跌倒,-Null as 查阅状态,null as 标本状态,-Null as 诊疗项目ID,Null as 试管编码,-Null as 执行标记,-Null as 屏蔽打印,-Null as 前提ID,-Null as 签名否,-Null as 文件ID," & _
                " -Null as 报告项,-Null as 报告ID,-Null as 收费细目ID, Null as 单量单位, -Null as 开嘱科室ID,-Null as 审核状态,-Null as 申请序号,-Null as 审核标记," & _
                " -Null as 高危药品,-Null as 标本部位,-NULL AS 用药目的,-NULL AS 检查报告id,-Null as 处方审查状态,-Null as 处方审查结果,-null as RIS预约ID,-null as RIS报告ID,-null as LIS报告ID,-null as RIS预约状态,null as 诊疗项目名称,-null as 检查方法,-null as 危急值ID From Dual" & _
                " Union ALL" & _
                " Select * From (" & strSQL & " And (Nvl(A.重整标志,1)=1 Or A.医嘱状态 IN(1,2)) " & strGroupBy & " Order by 婴儿ID,序号)"
        End If
    Else
        strSQL = strSQL & strGroupBy & " Order by 婴儿ID,序号"
    End If

    '访问历史空间处理
    If mblnMoved Then
        strSQL = Replace(strSQL, "/*+ RULE */", "/*+driving_site(a) driving_site(y)*/")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
    strCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID, CDate(strCurr), mvarCond.婴儿, IIF(mstr前提IDs = "", "0", mstr前提IDs), CDate(Format(strCurr, "yyyy-MM-dd 00:00:00")), mvarCond.开始时间, mvarCond.结束时间, mlng界面科室ID)
    
    If Not rsTmp.EOF Then
        strSQL = "Select a.医嘱id,decode(a.输血血型,1,'A',2,'B',3,'AB',4,'O','') As 血型 From 输血申请记录 A, 病人医嘱记录 B Where 医嘱id = b.Id And b.病人ID=[1] and b.主页ID=[2] And a.输血血型>0 and b.诊疗类别='K'"
        Set rs血型 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng病人ID, mlng主页ID)
        
        With vsAdvice
            .Redraw = False
            .MergeCells = flexMergeNever
            Call ClearAdviceData
            Call AddDataToVsf(rsTmp)
            '处理每行医嘱
            i = .FixedRows
            Do While i <= .Rows - 1
                .Cell(flexcpData, i, COL_开始时间) = CStr(.TextMatrix(i, COL_开始时间))    '合理用药接口调用时取数
                .Cell(flexcpData, i, COL_查阅状态) = Val(.TextMatrix(i, COL_查阅状态)) '报告查阅状态值
                If mvarCond.显示模式 = 0 Then
                    '简洁模式下处理日期的显示
                    strFormat = Format(.TextMatrix(i, COL_开始时间), "yyyy-MM-dd")
                    If strFormat = Format(strCurr, "yyyy-MM-dd") Then
                        .TextMatrix(i, COL_开始时间) = "今 天 " & Format(.TextMatrix(i, COL_开始时间), "HH:mm")
                    Else
                        If strFormat = strPreDay1 Then
                            .TextMatrix(i, COL_开始时间) = "昨 天 " & Format(.TextMatrix(i, COL_开始时间), "HH:mm")
                        ElseIf strFormat = strPreDay2 Then
                            .TextMatrix(i, COL_开始时间) = "前 天 " & Format(.TextMatrix(i, COL_开始时间), "HH:mm")
                        Else
                            .TextMatrix(i, COL_开始时间) = Format(.TextMatrix(i, COL_开始时间), "MM-dd HH:mm")
                        End If
                    End If
                End If
                
                If .TextMatrix(i, COL_诊疗类别) = "K" And gbln血库系统 Then
                    strSQL = "select zl_Get_输血执行血型([1]) as 血型 from dual"
                    Set rs血型 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(i, COL_ID)))
                    If Not rs血型.EOF Then
                        If rs血型!血型 & "" <> "" Then .TextMatrix(i, COL_皮试) = "(" & rs血型!血型 & ")"
                    End If
                End If
                
                '成药及中药的一些处理
                bln给药途径 = False: bln中药用法 = False: bln采集方法 = False: bln输血途径 = False
                If .TextMatrix(i, COL_诊疗类别) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln给药途径 = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '显示成药的给药途径+滴速
                                    .TextMatrix(j, COL_用法) = .TextMatrix(i, COL_用法) & .TextMatrix(i, COL_医生嘱托)

                                    If mvarCond.显示模式 = 0 Then    '合并用法列:用法 频率 天数
                                        strFormat = .TextMatrix(j, COL_用法)
                                        strTmp = .TextMatrix(j, COL_频率)
                                        If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                                        strTmp = .TextMatrix(j, COL_天数)
                                        If strTmp <> "" Then
                                            strFormat = strFormat & IIF(strFormat <> "", ",", "") & "共" & strTmp & "天"
                                        End If
                                        .TextMatrix(j, COL_用法) = strFormat
                                    End If

                                    '显示成药的执行性质
                                    If Val(.TextMatrix(j, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                        If Val(.TextMatrix(j, COL_执行标记)) = 2 Then
                                            .TextMatrix(j, COL_执行性质) = "不取药"
                                        Else
                                            .TextMatrix(j, COL_执行性质) = "自备药"
                                        End If
                                    ElseIf Val(.TextMatrix(j, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                        .TextMatrix(j, COL_执行性质) = "离院带药"
                                    Else
                                        .TextMatrix(j, COL_执行性质) = IIF(Val(.TextMatrix(j, COL_执行标记)) = 1, "自取药", "")
                                    End If
                                    
                                    '危急值ID是只关联在主医嘱主的，复制到药品行上
                                    .TextMatrix(j, COL_危急值ID) = .TextMatrix(i, COL_危急值ID)

                                    If mvarCond.显示模式 = 0 Then
                                        If .TextMatrix(j, COL_皮试) <> "" Then
                                            If Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "1") Then
                                                .TextMatrix(j, col_内容) = .TextMatrix(j, col_内容) & "," & .TextMatrix(j, COL_皮试)
                                            End If
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then
                            bln中药用法 = .TextMatrix(i - 1, COL_诊疗类别) = "7"    '中药用法行
                            bln采集方法 = .TextMatrix(i - 1, COL_诊疗类别) = "C"    '采集方法行

                            '采集方式的管码与一并的第一个检验相同
                            If bln采集方法 Then
                                j = .FindRow(.TextMatrix(i, COL_ID), .FixedRows, COL_相关ID)
                                If j <> -1 Then
                                    .TextMatrix(i, COL_试管编码) = .TextMatrix(j, COL_试管编码)
                                End If
                                .TextMatrix(i, COL_开始时间) = .TextMatrix(j, COL_开始时间)
                                .Cell(flexcpData, i, COL_开始时间) = CStr(.TextMatrix(j, COL_开始时间))
                                .Cell(flexcpData, i, COL_皮试) = .TextMatrix(i, COL_皮试)
                                .TextMatrix(i, COL_皮试) = "" '耐受试验的时间ID界面上不显示
                            End If

                            '显示中药配方或检验组合的执行科室
                            .TextMatrix(i, COL_执行科室) = .TextMatrix(i - 1, COL_执行科室)

                            If bln中药用法 Then
                                '显示中药配方执行性质
                                If Val(.TextMatrix(i - 1, COL_执行性质)) = 5 And Val(.TextMatrix(i, COL_执行性质)) <> 5 Then
                                    If Val(.TextMatrix(i - 1, COL_执行标记)) = 2 Then
                                        .TextMatrix(i, COL_执行性质) = "不取药"
                                    Else
                                        .TextMatrix(i, COL_执行性质) = "自备药"
                                    End If
                                ElseIf Val(.TextMatrix(i - 1, COL_执行性质)) <> 5 And Val(.TextMatrix(i, COL_执行性质)) = 5 Then
                                    .TextMatrix(i, COL_执行性质) = "离院带药"
                                Else
                                    .TextMatrix(i, COL_执行性质) = IIF(Val(.TextMatrix(i - 1, COL_执行标记)) = 1, "自取药", "")
                                End If
                            Else
                                .TextMatrix(i, COL_执行性质) = ""
                            End If

                            '删除单味中药行,以及检验组合中的检验项目
                            strTmp = ""
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_相关ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    .TextMatrix(i, COL_报告项) = .TextMatrix(j, COL_报告项)    '检验、配方以首行医嘱为准
                                    .TextMatrix(i, COL_文件ID) = .TextMatrix(j, COL_文件ID)
                                    If bln中药用法 Then  '单味中药行ID记录下来，合理用药删除使用
                                        strTmp = strTmp & IIF(strTmp = "", .TextMatrix(j, COL_ID), "," & .TextMatrix(j, COL_ID))
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    If bln中药用法 Then
                                        .Cell(flexcpData, i, COL_相关ID) = strTmp
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf .TextMatrix(i - 1, COL_诊疗类别) = "K" And Val(.TextMatrix(i - 1, COL_ID)) = Val(.TextMatrix(i, COL_相关ID)) Then
                        bln输血途径 = True
                        '显示输血途径
                        .TextMatrix(i - 1, COL_用法) = .TextMatrix(i, COL_用法) & .TextMatrix(i, COL_医生嘱托)
                    Else
                        .TextMatrix(i, COL_执行性质) = ""
                    End If
                End If
                '会诊医嘱与请会诊病历可能有关联
                If .TextMatrix(i, COL_诊疗类别) = "Z" And .TextMatrix(i, COL_操作类型) = "7" And .TextMatrix(i, COL_报告ID) <> "" Then
                     .TextMatrix(i, COL_报告ID) = ""
                End If
                '处理可见行的的一些标识:排开不可见但暂时未删除的行
                If Not (bln给药途径 Or bln输血途径) And .TextMatrix(i, COL_诊疗类别) <> "7" Then
                    '行高：为了支持zl9PrintMode:Resize之后,取RowHeight可能小于RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    '只显示需的报告的医嘱
                    If mvarCond.过滤模式 = 3 Then
                        If Val(.TextMatrix(i, COL_报告项)) = 0 Then .RowHidden(i) = True: .RowHeight(i) = 0
                        '显示各种报告的医嘱
                        If mvarCond.报告 = 1 Then ' 检查
                            If Not .TextMatrix(i, COL_诊疗类别) = "D" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.报告 = 2 Then '检验
                            If Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "C") Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.报告 = 3 Then ' 其它
                            If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" Or .TextMatrix(i, COL_诊疗类别) = "D" Or .TextMatrix(i, COL_诊疗类别) = "C" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    Else
                        '只显示需的报告的医嘱
                        If mvarCond.是报告医嘱 And Not mvarCond.非报告医嘱 And Val(.TextMatrix(i, COL_报告项)) = 0 Then
                            .RowHidden(i) = True: .RowHeight(i) = 0
                        ElseIf Not mvarCond.是报告医嘱 And mvarCond.非报告医嘱 And Val(.TextMatrix(i, COL_报告项)) <> 0 Then
                            .RowHidden(i) = True: .RowHeight(i) = 0
                        End If
                    End If
                    
                    '重整医嘱分隔
                    If Val(.TextMatrix(i, COL_ID)) = 0 Then
                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = "━━━━ 重整医嘱(" & Format(mdat重整, "yyyy-MM-dd HH:mm") & ") ━━━━"
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed
                        .Cell(flexcpAlignment, i, .FixedCols, i, .Cols - 1) = 4

                        .MergeRow(i) = True
                        .MergeCells = flexMergeFree
                    End If

                    '处理小数点问题,暂未想到办法
                    If Left(.TextMatrix(i, COL_总量), 1) = "." Then
                        .TextMatrix(i, COL_总量) = "0" & .TextMatrix(i, COL_总量)
                    End If
                    If Left(.TextMatrix(i, COL_单量), 1) = "." Then
                        .TextMatrix(i, COL_单量) = "0" & .TextMatrix(i, COL_单量)
                    End If

                    '医嘱颜色
                    blnDo = False
                    If Val(.TextMatrix(i, COL_医嘱状态)) = 2 Then
                        '校对疑问
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H80&    '深红
                        blnDo = True
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 4 Then
                        '已作废
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '灰色
                        .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                        blnDo = True
                    ElseIf InStr(",8,9,", Val(.TextMatrix(i, COL_医嘱状态))) > 0 Then
                        '已停止,已确认停止:长嘱都以终止时间进行判断
                        If strCurr >= .TextMatrix(i, COL_终止时间) Or .TextMatrix(i, COL_期效) = "临嘱" Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '灰色
                            blnDo = True
                        ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 8 And strCurr < .TextMatrix(i, COL_终止时间) Then
                            '长嘱,停止后,停止时间未到这一种情况
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080    '浅蓝
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 6 Then
                        '已暂停
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 6), "yyyy-MM-dd HH:mm")
                        If strCurr >= strTime Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000&    '深绿
                            blnDo = True
                        Else
                            '长嘱,暂停后,暂停时间未到这一种情况
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080    '浅蓝
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_医嘱状态)) = 7 Then
                        '已启用
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 7), "yyyy-MM-dd HH:mm")
                        If strCurr < strTime Then
                            '长嘱,启用后,启用时间未到这一种情况
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H4AAD00    '浅绿
                            blnDo = True
                        End If
                    End If
                    If Not blnDo Then
                        If Val(.TextMatrix(i, COL_医嘱状态)) <> 1 And Val(.TextMatrix(i, COL_ID)) <> 0 Then
                            '已通过校对(也包含后续的多个状态)
                            If Format(.TextMatrix(i, COL_上次执行), "YYYY-MM-DD") >= Format(strCurr, "YYYY-MM-DD") Then  '当天已发送的(长嘱可能发送到将来)
                                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HA08000               '海蓝
                            Else
                                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000    '深蓝
                            End If
                        End If
                    End If

                    '校对后术前术后医嘱红色显示
                    If .TextMatrix(i, COL_诊疗类别) = "Z" And (Val(.TextMatrix(i, COL_操作类型)) = 4 Or Val(.TextMatrix(i, COL_操作类型)) = 14) _
                       And InStr(",-1,1,2,4,", Val(.TextMatrix(i, COL_医嘱状态))) = 0 Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed    '红色
                    End If

                    '发送后转科医嘱红色显示
                    If .TextMatrix(i, COL_诊疗类别) = "Z" And Val(.TextMatrix(i, COL_操作类型)) = 3 And Val(.TextMatrix(i, COL_医嘱状态)) = 8 Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed    '红色
                    End If

                    '毒麻精药品标识:中药配方及组成味中药不处理
                    If .TextMatrix(i, COL_毒理分类) <> "" Then
                        If InStr(",麻醉药,毒性药,精神药,精神I类,精神II类,", .TextMatrix(i, COL_毒理分类)) > 0 Then
                            .Cell(flexcpFontBold, i, col_医嘱内容) = True
                            .Cell(flexcpFontBold, i, col_内容) = True
                        End If
                    End If

                    '皮试结果标识
                    If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "1" And .TextMatrix(i, COL_皮试) <> "" Then
                        j = GetSkinTestResult(Val(.TextMatrix(i, COL_诊疗项目ID)), .TextMatrix(i, COL_皮试))
                        .Cell(flexcpForeColor, i, COL_皮试) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, COL_皮试))
                    End If


                    '图标处理
                    '报告列及打印状态标识
                    Call SetAdviceReportIcon(i)

                    '自由录入
                    If Val(.TextMatrix(i, COL_诊疗项目ID)) = 0 And Val(.TextMatrix(i, COL_ID)) <> 0 Then
                        Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("自由").Picture
                    End If
                    '备用医嘱
                    If .TextMatrix(i, COL_频率) = "必要时" Or .TextMatrix(i, COL_频率) = "需要时" Then
                        Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("备用").Picture
                    End If
                    '紧急标志:一并给药只显示在第一行
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_标志)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("紧急").Picture
                        ElseIf Val(.TextMatrix(i, COL_标志)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("补录").Picture
                        End If

                        If Val(.TextMatrix(i, COL_医嘱状态)) < 2 Then   '新开或暂存的医嘱
                            Select Case Val(.TextMatrix(i, COL_审核状态))
                                '0-无需审核，1-待审核，2-审核通过，3-审核未通过
                            Case 1
                                If .TextMatrix(i, COL_诊疗类别) = "K" And Val(.TextMatrix(i, COL_检查方法)) = 1 Then
                                    '用血医嘱审核图标单独显示(表明是有医生核对)
                                    Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("核对").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                                End If
                            Case 2
                                If Not (.TextMatrix(i, COL_诊疗类别) = "K" And Val(.TextMatrix(i, COL_检查方法)) = 1) Then
                                    Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核通过").Picture
                                End If
                            Case 3
                                Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核未通过").Picture
                            Case 4, 5
                                If gbln血库系统 = False Then
                                    Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                                End If
                            Case 7
                                Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待签发").Picture
                            Case Else
                            End Select
                            .Cell(flexcpPictureAlignment, i, COL_F标志) = 4
                        End If
                        '处方审查系统
                        If .TextMatrix(i, COL_处方审查状态) = "0" Then
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("待审核").Picture
                        ElseIf .TextMatrix(i, COL_处方审查状态) = "2" Or .TextMatrix(i, COL_处方审查结果) = "1" Then
                            '超时免审当作合格处理
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核通过").Picture
                        ElseIf .TextMatrix(i, COL_处方审查结果) = "2" Then
                            ' 不合格
                            Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("审核未通过").Picture
                        End If
                    End If

                    '未用医嘱标识
                    If Val(.TextMatrix(i, COL_执行标记)) = -1 Then
                        Set .Cell(flexcpPicture, i, COL_F标志) = frmIcons.imgFlag.ListImages("未用").Picture
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '灰色
                    End If


                    'Pass:根据审查结果显示警示灯
                    If mblnPass Then
                        If .TextMatrix(i, COL_警示) <> "" Then
                            Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, Val(.TextMatrix(i, COL_警示)))
                            .TextMatrix(i, COL_警示) = ""
                        End If
                    End If
                End If

                If bln给药途径 Or bln输血途径 Then
                    .RemoveItem i
                Else
                    '简洁模式，组合医嘱内容
                    If mvarCond.显示模式 = 0 And mvarCond.过滤模式 <> 3 Then
                        strFormat = .TextMatrix(i, col_医嘱内容)
                        If .TextMatrix(i, COL_诊疗类别) <> "Z" And Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 And InStr(strFormat, "重整医嘱") = 0 Then
                            '医嘱内容定义中包含了相关项时，不再重复组合
                            mrsDefine.Filter = "诊疗类别='" & .TextMatrix(i, COL_诊疗类别) & "'"
                            If Not (.TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "1") Then
                                strFormat = strFormat & .TextMatrix(i, COL_皮试)
                            End If

                            If Not (InStr("5,6,7", .TextMatrix(i, COL_诊疗类别)) = 0 And .TextMatrix(i, COL_频率) = "一次性") Then
                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[总量]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_总量)
                                    If strTmp <> "" Then strFormat = strFormat & ",共" & strTmp
                                End If

                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!医嘱内容, "[单量]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_单量)
                                    If strTmp <> "" Then strFormat = strFormat & ",每次" & strTmp
                                End If
                            End If
                        End If
                        .TextMatrix(i, col_内容) = strFormat


                        '合并用法列:用法 频率 天数(一并给药的在前面已处理)
                        If .TextMatrix(i, COL_诊疗类别) <> "Z" And Val(.TextMatrix(i, COL_诊疗项目ID)) <> 0 And InStr(strFormat, "重整医嘱") = 0 Then
                            
                            '简洁模式下除药品、手术项目外其他的医嘱不显示用法
                            If .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) = 0 Or _
                                InStr(",5,6,7,", "," & .TextMatrix(i, COL_诊疗类别) & ",") > 0 And Val(.TextMatrix(i, COL_相关ID)) <> 0 Then
                                strFormat = .TextMatrix(i, COL_用法)
                            Else
                                strFormat = ""
                            End If
                            
                            '检验 '检查 '输血 '手术 '护理等级 简洁模式下不显示频率
                            If .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_操作类型)) = 6 Or _
                                .TextMatrix(i, COL_诊疗类别) = "D" And Val(.TextMatrix(i, COL_相关ID)) = 0 Or _
                                .TextMatrix(i, COL_诊疗类别) = "K" And Val(.TextMatrix(i, COL_相关ID)) = 0 Or _
                                .TextMatrix(i, COL_诊疗类别) = "F" And Val(.TextMatrix(i, COL_相关ID)) = 0 Or _
                                .TextMatrix(i, COL_诊疗类别) = "H" And Val(.TextMatrix(i, COL_操作类型)) = 1 Then
                                strTmp = ""
                            Else
                                strTmp = .TextMatrix(i, COL_频率)
                            End If
                            If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                            strTmp = .TextMatrix(i, COL_天数)
                            If strTmp <> "" Then
                                strFormat = strFormat & IIF(strFormat <> "", ",", "") & "共" & strTmp & "天"
                            End If

                            .TextMatrix(i, COL_用法) = strFormat
                        End If
                        
                        '简洁模式下，临嘱终止时间显示为空。
                        If .TextMatrix(i, COL_期效) = "临嘱" Then .TextMatrix(i, COL_终止时间) = ""
                    End If
                    
                    If mvarCond.过滤模式 = 3 Then
                        '如果是报告页签下，内容 列 可能为空，重新赋值
                        .TextMatrix(i, col_内容) = .TextMatrix(i, col_医嘱内容)
                        If Val(.TextMatrix(i, COL_报告ID)) = 0 And .TextMatrix(i, COL_检查报告ID) = "" And Val(.TextMatrix(i, COL_RIS报告ID)) = 0 And Val(.TextMatrix(i, COL_LIS报告ID)) = 0 Then
                            .TextMatrix(i, COL_查阅状态) = "未出"
                        Else
                            .TextMatrix(i, COL_查阅状态) = "查阅"
                            If Val(.Cell(flexcpData, i, COL_查阅状态)) = 0 Then  '未读
                                .Cell(flexcpForeColor, i, COL_查阅状态, i, COL_查阅状态) = &HFF0000     '蓝色
                            ElseIf Val(.Cell(flexcpData, i, COL_查阅状态)) = 2 Then  '部分已读
                                .Cell(flexcpForeColor, i, COL_查阅状态, i, COL_查阅状态) = &HFF00FF     '紫色
                            Else
                                .Cell(flexcpForeColor, i, COL_查阅状态, i, COL_查阅状态) = &H80&      '暗红
                            End If
                            .Cell(flexcpFontUnderline, i, COL_查阅状态, i, COL_查阅状态) = True
                        End If
                        '增加过滤未出的报告和已出的报告
                        If .RowHidden(i) = False Then
                            If Not IIF(.TextMatrix(i, COL_查阅状态) = "未出", mvarCond.未出报告, mvarCond.已出报告) Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    End If
                    i = i + 1
                End If
            Loop
            
            '设置医嘱内容单元格的图标,电子签名标识，屏蔽打印,危急值
            For i = 1 To .Rows - 1
                Call SetAdviceIcon(i)
            Next
            
            '自动调整行高
            If mvarCond.显示模式 = 0 And mvarCond.过滤模式 <> 3 Then
                If InStr("2505,3345,1005,1335", .ColWidth(COL_用法)) > 0 Then .ColWidth(COL_用法) = IIF(mlngFontSize = 9, 2505, 3345)   '用户未改该列宽时才设置
                .AutoSize col_内容, COL_用法
                .ColWidth(COL_开始时间) = IIF(mlngFontSize = 9, 1130, 1510)
            Else
                If InStr("2505,3345,1005,1335", .ColWidth(COL_用法)) > 0 Then .ColWidth(COL_用法) = IIF(mlngFontSize = 9, 1005, 1335)
                .AutoSize col_医嘱内容, COL_用法
                .ColWidth(COL_开始时间) = IIF(mlngFontSize = 9, 1530, 2040)
            End If

            '固定列图标对齐:设置为中对齐,不然擦边框时可能有问题
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '电子签名图标对齐
            .Cell(flexcpPictureAlignment, .FixedRows, col_医嘱内容, .Rows - 1, col_医嘱内容) = 0
            Call SetTag一并给药
            Call Set标本状态
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
    Call SetAdviceColVisible
    '只有临嘱时才用红色表线
    vsAdvice.GridColor = IIF(mvarCond.过滤模式 = "2", &H8080FF, vsAdvice.GridColorFixed)
    
    imgColSel.Visible = (mvarCond.显示模式 = 1 And mvarCond.过滤模式 <> 3)
    
    Call LocatedDefaultAdviceRow(lng医嘱ID)
    
    Screen.MousePointer = 0
    LoadAdvice = True
    If Not mfrmParent Is Nothing Then
        '新版护士站调用时，默认设置一次颜色，afterrowcolchange中无法设置。
        If mfrmParent.Name = "frmInNurseRoutine" Then
            If vsAdvice.Col >= vsAdvice.FixedCols Then
                vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_开始时间)
            End If
        End If
    End If
    '自动刷新医嘱提醒区域
    If blnRefreshNotify Then RaiseEvent RequestRefresh(True)
    Exit Function
errH:
    vsAdvice.Redraw = True
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetAdviceIcon(ByVal lngRow As Long)
'功能：根据当前行的内容设置医嘱内容的图标标识
'说明：注意是单行设置，不是一组设置
    Dim int图标数 As Integer '医嘱内容上面的图标个数
    
    int图标数 = 1
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_审核标记)) = 2 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = img16.ListImages("停嘱申请").Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = img16.ListImages("停嘱申请").Picture
    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_签名否)) = 1 And Val(vsAdvice.TextMatrix(lngRow, COL_屏蔽打印)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = img16dbl.ListImages(1).Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = img16dbl.ListImages(1).Picture
    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_签名否)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgSign.ListImages("签名").Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgSign.ListImages("签名").Picture
    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_屏蔽打印)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = img16.ListImages("屏蔽打印").Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = img16.ListImages("屏蔽打印").Picture
    Else
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = Nothing
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = Nothing
        int图标数 = 0
    End If
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_高危药品)) > 0 Then
        If vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) Is Nothing Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgQuestion.ListImages("高危药品").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgQuestion.ListImages("高危药品").Picture
            int图标数 = 1
        Else
            If vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) <> frmIcons.imgQuestion.ListImages("高危药品").Picture Then
                pictmp.Cls
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("高危药品").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
                int图标数 = 2
            End If
        End If
    End If
    
    '危急值图标
    If Val(vsAdvice.TextMatrix(lngRow, COL_危急值ID)) > 0 Then
        If int图标数 = 0 Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgQuestion.ListImages("危急值").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgQuestion.ListImages("危急值").Picture
        ElseIf int图标数 = 1 Then
            pictmp.Cls
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("危急值").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
            int图标数 = 2
        ElseIf int图标数 = 2 Then
            pictmp.Cls
            pictmp.Width = 720
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, 480, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("危急值").Picture, 480, 0, 240, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
            pictmp.Width = 480
            int图标数 = 3
        End If
    End If
    
    '易跌倒图标
    If Val(vsAdvice.TextMatrix(lngRow, COL_易跌倒)) > 0 Then
        If int图标数 = 0 Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = frmIcons.imgQuestion.ListImages("易跌倒").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = frmIcons.imgQuestion.ListImages("易跌倒").Picture
        ElseIf int图标数 = 1 Then
            pictmp.Cls
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, pictmp.Width / 2, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
            int图标数 = 2
        ElseIf int图标数 = 2 Then
            pictmp.Cls
            pictmp.Width = 720
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, 480, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, 480, 0, 240, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
            pictmp.Width = 480
            int图标数 = 3
        ElseIf int图标数 = 3 Then
            pictmp.Cls
            pictmp.Width = 960
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容), 0, 0, 720, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("易跌倒").Picture, 720, 0, 240, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_医嘱内容) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_内容) = pictmp.Image
            pictmp.Width = 480
            int图标数 = 4
        End If
    End If
End Sub

Private Function RowIs配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否中药配方行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='7' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs配方行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否检验组合行
'说明：指定行为显示行,且类别="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From 病人医嘱记录 Where Rownum=1 And 诊疗类别='C' And 相关ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs检验行 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPrice(ByVal lngRow As Long) As Boolean
'功能：读取指定医嘱的计价,并根据当前的诊疗收费 关系进行更新
    Dim rs诊疗项目 As New ADODB.Recordset
    Dim rs收费细目 As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim str医嘱IDs As String, str收费细目IDs As String, str诊疗收费 As String
    Dim strSQL As String, i As Long, j As Long
    Dim bln配方行 As Boolean, bln检验行 As Boolean, blnLoad As Boolean
    Dim lng病人科室ID As Long, lng执行科室ID As Long
    Dim dblPrice As Double, lng材料ID As Long
    Dim lng医嘱ID As Long, lng相关ID As Long
    Dim strPriceType As String

    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIs配方行(lngRow)
            bln检验行 = RowIs检验行(lngRow)
        End If
        
        lng医嘱ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
        lng相关ID = Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
                                    
        blnLoad = True
        
        '药品、卫材的计价
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "4" Then
            '卫材计价
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,NULL as 检查方法,0 as 执行标记,0 as 费用性质,0 as 收费方式," & _
                " A.收费细目ID,1 as 住院包装,C.计算单位,1 as 数量,Decode(Nvl(C.是否变价,0),1,Nvl(B.单价,D.缺省价格),D.现价) as 单价,A.执行科室ID,0 as 从项,C.类别 as 收费类别" & _
                " From 病人医嘱记录 A,病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1] And A.ID=B.医嘱ID(+) And A.收费细目ID=C.ID And Nvl(A.执行性质,0) Not IN(0,5)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "5", "6", "7") & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
                
                blnLoad = False
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            '中,西成药:可能按规格下医嘱,计算1个住院包装的单价
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,NULL as 检查方法,0 as 执行标记,0 as 费用性质,0 as 收费方式," & _
                " C.ID as 收费细目ID,B.住院包装,B.住院单位 as 计算单位,1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.住院包装 as 单价," & _
                " A.执行科室ID,0 as 从项,C.类别 as 收费类别" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.诊疗项目ID=B.药名ID And B.药品ID=C.ID And Nvl(A.执行性质,0) Not IN(0,5)" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "5", "6", "7") & _
                " And (A.收费细目ID is NULL Or A.收费细目ID=B.药品ID)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3) And D.收费细目ID=C.ID" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
                
                '仅一并给药(如果是)的第一成药行才显示给药途径的计价
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_相关ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
        ElseIf bln配方行 Then
            '中草药:一定对应有规格记录且填写了收费细目ID
            strSQL = "Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,NULL as 标本部位,NULL as 检查方法,0 as 执行标记,0 as 费用性质,0 as 收费方式," & _
                " C.ID as 收费细目ID,B.住院包装,B.住院单位 as 计算单位,1 as 数量,Decode(Nvl(C.是否变价,0),1,-NULL,D.现价)*B.住院包装 as 单价," & _
                " A.执行科室ID,0 as 从项,C.类别 as 收费类别" & _
                " From 病人医嘱记录 A,药品规格 B,收费项目目录 C,收费价目 D" & _
                " Where A.诊疗类别='7' And A.相关ID=[1]" & _
                GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "5", "6", "7") & _
                " And A.收费细目ID=B.药品ID And A.收费细目ID=C.ID And C.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3)" & _
                " And D.收费细目ID=C.ID And Nvl(A.执行性质,0) Not IN(0,5)" & _
                " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))"
        End If
        
        '读取现有计价(取最新价格)：除药品、卫材外的计价,包含相关医嘱计价
        '不计价,手工计价的医嘱不读取
        '用Union方式可以利用索引
        If blnLoad Then
            '不是新开的医嘱，根据病人医嘱计价提取
            If InStr(",1,2,-1,", vsAdvice.TextMatrix(lngRow, COL_医嘱状态)) = 0 Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0) as 费用性质,Nvl(B.收费方式,0) as 收费方式," & _
                    " B.收费细目ID,1 as 住院包装,C.计算单位,B.数量,Decode(C.是否变价,1,B.单价,Sum(D.现价)) as 单价," & _
                    " Nvl(B.执行科室ID,A.执行科室ID) as 执行科室ID,Nvl(B.从项,0) as 从项,C.类别 as 收费类别" & _
                    " From 病人医嘱记录 A,病人医嘱计价 B,收费项目目录 C,收费价目 D" & _
                    " Where A.诊疗类别 Not IN('4','5','6','7') And A.ID=B.医嘱ID" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "5", "6", "7") & _
                    " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5) And B.收费细目ID=C.ID And B.收费细目ID=D.收费细目ID" & _
                    " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                    " And (A.ID=[1]" & IIF(lng相关ID <> 0, " Or A.ID=[2]", "") & " Or A.相关ID=[1])" & _
                    " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0),Nvl(B.收费方式,0)," & _
                    " B.收费细目ID,C.类别,C.计算单位,B.数量,C.是否变价,B.单价,Nvl(B.执行科室ID,A.执行科室ID),Nvl(B.从项,0)"
            Else
            '新开的医嘱，根据诊疗收费 关系提取(非药变价显示为0)
            '门诊留观病人，适用科室采用门诊的
            '几种对应的计价：
            '   1.加收的费用，只在主项目上面加收，目前只有床旁或术中这种情况
            '   2.基本的费用，但是具体的检查部位和检查方法的
            '   3.基本的费用，非检查部位和方法的(注意检验标本填写在标本部位中)
                lng材料ID = 0 '检验试管费用,只收取试管对应的卫材费用
                If vsAdvice.TextMatrix(lngRow, COL_试管编码) <> "" Then
                    lng材料ID = GetTubeMaterial(vsAdvice.TextMatrix(lngRow, COL_试管编码))
                End If
                
                str诊疗收费 = "Select * From (" & _
                    "Select C.诊疗项目ID,C.收费项目ID,C.检查部位,C.检查方法,C.费用性质,C.收费数量,C.固有对照,C.从属项目,C.收费方式,c.适用科室id" & _
                    " ,Max(Nvl(c.适用科室id, 0)) Over(Partition By c.诊疗项目id, c.检查部位, c.检查方法, c.费用性质) As Top" & _
                    " From 诊疗收费关系 C,病人医嘱记录 A Where (A.ID=[1]" & IIF(lng相关ID <> 0, " Or A.ID=[2]", "") & " Or A.相关ID=[1]) And A.诊疗项目ID+0=C.诊疗项目ID" & _
                    "   And (a.相关id Is Null And a.执行标记 In (1, 2) And c.费用性质 = 1 Or" & vbNewLine & _
                    "   a.标本部位 = c.检查部位 And a.检查方法 = c.检查方法 And Nvl(c.费用性质, 0) = 0 Or" & vbNewLine & _
                    "   (a.检查方法 Is Null or a.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=a.诊疗项目ID And Z.操作类型='4')) And Nvl(c.费用性质, 0) = 0 And c.检查部位 Is Null And c.检查方法 Is Null)" & _
                    "      And (C.适用科室ID is Null or C.适用科室ID = Nvl(A.执行科室ID,[4]) And C.病人来源 = " & IIF(mlng病人性质 = 1, 1, 2) & ")" & _
                    " ) Where Nvl(适用科室id, 0) = Top"
                
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0) as 费用性质,Nvl(B.收费方式,0) as 收费方式," & _
                    " B.收费项目ID as 收费细目ID,1 as 住院包装,C.计算单位,B.收费数量 as 数量,Decode(C.是否变价,1,Sum(D.缺省价格),Sum(D.现价)) as 单价," & _
                    " A.执行科室ID,Nvl(B.从属项目,0) as 从项,C.类别 as 收费类别" & _
                    " From 病人医嘱记录 A,(" & str诊疗收费 & ") B,收费项目目录 C,收费价目 D" & _
                    " Where A.诊疗类别 Not IN('4','5','6','7') And A.医嘱状态 IN(-1,1,2) And A.诊疗项目ID+0=B.诊疗项目ID" & _
                    GetPriceGradeSQL(mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级, "C", "D", "5", "6", "7") & _
                    " And (A.相关ID is Null And A.执行标记 IN(1,2) And B.费用性质=1" & _
                    "       Or A.标本部位=B.检查部位 And A.检查方法=B.检查方法 And Nvl(B.费用性质,0)=0" & _
                    "       Or (A.检查方法 is Null or a.诊疗类别 = 'E' And Exists(Select 1 From 诊疗项目目录 Z Where Z.id=a.诊疗项目ID And Z.操作类型='4')) And Nvl(B.费用性质,0)=0 And B.检查部位 is Null And B.检查方法 is Null)" & _
                    " And Nvl(A.计价特性,0)=0 And Nvl(A.执行性质,0) Not IN(0,5) And B.收费项目ID=C.ID And B.收费项目ID=D.收费细目ID" & _
                    " And ((Sysdate Between D.执行日期 and D.终止日期) or (Sysdate>=D.执行日期 And D.终止日期 is NULL))" & _
                    " And (C.撤档时间 is NULL Or C.撤档时间=To_Date('3000-01-01','YYYY-MM-DD')) And C.服务对象 IN(" & IIF(mlng病人性质 = 1, 1, 2) & ",3)" & _
                    " And (Nvl(B.收费方式,0)=1 And C.类别='4' And B.收费项目ID=[3] Or Not(Nvl(B.收费方式,0)=1 And C.类别='4' And [3]<>0))" & _
                    " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null) And (A.ID=[1]" & IIF(lng相关ID <> 0, " Or A.ID=[2]", "") & " Or A.相关ID=[1])" & _
                    " Group by A.ID,A.相关ID,A.序号,A.诊疗类别,A.诊疗项目ID,A.标本部位,A.检查方法,A.执行标记,Nvl(B.费用性质,0),Nvl(B.收费方式,0)," & _
                    " B.收费项目ID,C.类别,C.计算单位,B.收费数量,C.是否变价,A.执行科室ID,Nvl(B.从属项目,0)"
            End If
        End If
        strSQL = strSQL & " Order by 序号,费用性质,从项,收费类别"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱计价", "H病人医嘱计价")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng医嘱ID, lng相关ID, lng材料ID, mlng病区ID, mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
        '显示计价内容
        If Not rsTmp.EOF Then
            '确定显示行数
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '获取诊疗项目,收费细目信息
            For i = 1 To rsTmp.RecordCount
                If InStr("," & str医嘱IDs & ",", "," & rsTmp!ID & ",") = 0 Then str医嘱IDs = str医嘱IDs & "," & rsTmp!ID
                If InStr("," & str收费细目IDs & ",", "," & rsTmp!收费细目ID & ",") = 0 Then str收费细目IDs = str收费细目IDs & "," & rsTmp!收费细目ID
                rsTmp.MoveNext
            Next
            str医嘱IDs = Mid(str医嘱IDs, 2)
            str收费细目IDs = Mid(str收费细目IDs, 2)
                        
            If mblnMoved Then
            '通过DBLink连接的远程历史库特殊处理：使用f_Num2list后无法利用索引,加driving_site避免把远程大表复制到当前服务器
                
                strSQL = "Select /*+driving_site(a)*/ B.ID,B.类别,C.名称 as 类别名称,B.名称,B.标本部位" & _
                      " From H病人医嘱记录 A,诊疗项目目录 B,诊疗项目类别 C" & _
                      " Where A.诊疗项目ID=B.ID And B.类别=C.编码 And A.ID "
                If InStr(str医嘱IDs, ",") > 0 Then              '少数SQL是这种，只好不用绑定变量
                    strSQL = strSQL & " In(" & str医嘱IDs & ")"
                Else
                    strSQL = strSQL & " = [1]"
                End If
            Else
              strSQL = "Select /*+cardinality(d,10)*/ B.ID,B.类别,C.名称 as 类别名称,B.名称,B.标本部位" & _
                  " From 病人医嘱记录 A,诊疗项目目录 B,诊疗项目类别 C,Table(f_Num2list([1])) D" & _
                  " Where A.ID = D.Column_Value And A.诊疗项目ID=B.ID And B.类别=C.编码"
            End If
            Set rs诊疗项目 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str医嘱IDs)
            
            strSQL = "Select A.ID,A.类别,B.名称 as 类别名称,A.编码," & _
                " A.名称,A.规格,A.产地,A.费用类型,A.是否变价" & _
                " From 收费项目目录 A,收费项目类别 B,Table(f_Num2list([1])) D" & _
                " Where A.类别=B.编码 And A.ID = D.Column_Value"
            strSQL = "Select /*+ Rule*/ A.ID,A.类别,A.类别名称,A.编码,Nvl(B.名称,A.名称) as 名称," & _
                " A.规格,A.产地,A.费用类型,A.是否变价,C.跟踪在用" & _
                " From (" & strSQL & ") A,收费项目别名 B,材料特性 C" & _
                " Where A.ID=C.材料ID(+) And A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[2]"
            Set rs收费细目 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str收费细目IDs, IIF(gbyt药品名称显示 = 0, 1, 3))
            
            '显示每行内容
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs诊疗项目.Filter = "ID=" & rsTmp!诊疗项目ID
                rs收费细目.Filter = "ID=" & rsTmp!收费细目ID
                
                '计价医嘱
                If rsTmp!诊疗类别 = "4" Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "卫生材料-" & rs诊疗项目!名称
                ElseIf InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "药品医嘱-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "给药途径-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                    .TextMatrix(i, COLPrice("计价医嘱")) = "输血途径-" & rs诊疗项目!名称
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "采集方法-" & rs诊疗项目!名称
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "中药煎法-" & rs诊疗项目!名称
                    Else
                        .TextMatrix(i, COLPrice("计价医嘱")) = "中药用法-" & rs诊疗项目!名称
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "检验项目-" & rs诊疗项目!名称
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        '部位及方法
                        .TextMatrix(i, COLPrice("计价医嘱")) = "检查部位-" & NVL(rsTmp!标本部位) & "(" & NVL(rsTmp!检查方法) & ")"
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "附加手术-" & rs诊疗项目!名称
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, COLPrice("计价医嘱")) = "麻醉项目-" & rs诊疗项目!名称
                    End If
                Else
                    If NVL(rsTmp!费用性质, 0) = 1 Then
                        '床旁或术中加收费用
                        .TextMatrix(i, COLPrice("计价医嘱")) = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称 & "(" & Decode(NVL(rsTmp!执行标记, 0), 1, "床旁", 2, "术中", "") & "加收)"
                    Else
                        .TextMatrix(i, COLPrice("计价医嘱")) = rs诊疗项目!类别名称 & "医嘱-" & rs诊疗项目!名称
                    End If
                End If
                
                '类别
                .TextMatrix(i, COLPrice("类别")) = rs收费细目!类别名称
                '收费项目:规格/产地
                .TextMatrix(i, COLPrice("收费项目")) = rs收费细目!名称
                If Not IsNull(rs收费细目!产地) Then
                    .TextMatrix(i, COLPrice("收费项目")) = .TextMatrix(i, COLPrice("收费项目")) & "(" & rs收费细目!产地 & ")"
                End If
                If Not IsNull(rs收费细目!规格) Then
                    .TextMatrix(i, COLPrice("收费项目")) = .TextMatrix(i, COLPrice("收费项目")) & " " & rs收费细目!规格
                End If
                
                '计算单位:药嘱药品为住院单位,非药嘱药品为售价单位
                .TextMatrix(i, COLPrice("单位")) = NVL(rsTmp!计算单位)
                '计价数量:药嘱药品为1,非药嘱药品为对应售价数
                .TextMatrix(i, COLPrice("计价数量")) = FormatEx(rsTmp!数量, 5)
                
                '执行科室
                lng执行科室ID = NVL(rsTmp!执行科室ID, 0)
                If rs收费细目!类别 = "4" And NVL(rs收费细目!跟踪在用, 0) = 1 Or _
                    InStr(",5,6,7,", rs收费细目!类别) > 0 And InStr(",5,6,7,", rs诊疗项目!类别) = 0 Then
                    lng病人科室ID = mlng科室ID
                    lng执行科室ID = Get收费执行科室ID(mlng病人ID, mlng主页ID, rs收费细目!类别, rs收费细目!ID, 4, lng病人科室ID, 0, 2, lng执行科室ID, , , 2)
                End If
                
                '单价处理
                If InStr(",5,6,7,", rs收费细目!类别) > 0 Then
                    If NVL(rs收费细目!是否变价, 0) = 1 Then
                        '求药品时价
                        If InStr(",5,6,7,", rs诊疗项目!类别) > 0 Then
                            '药嘱药品计算一个住院包装的住院时价
                            .TextMatrix(i, COLPrice("单价")) = CalcDrugPrice(rs收费细目!ID, lng执行科室ID, NVL(rsTmp!住院包装, 1), , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级)
                            .TextMatrix(i, COLPrice("单价")) = Format(Val(.TextMatrix(i, COLPrice("单价"))) * NVL(rsTmp!住院包装, 0), gstrDecPrice)
                        Else
                            '非药嘱药品计算相对售价数量的售价实价
                            .TextMatrix(i, COLPrice("单价")) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, NVL(rsTmp!数量, 0), , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                        End If
                    Else
                        '药嘱药品为住院单价,非药药品为售价
                        .TextMatrix(i, COLPrice("单价")) = Format(NVL(rsTmp!单价), gstrDecPrice)
                    End If
                ElseIf rs收费细目!类别 = "4" And NVL(rs收费细目!跟踪在用, 0) = 1 And NVL(rs收费细目!是否变价, 0) = 1 Then
                    '时价卫材的单价和药品一样计算
                    .TextMatrix(i, COLPrice("单价")) = Format(CalcDrugPrice(rs收费细目!ID, lng执行科室ID, NVL(rsTmp!数量, 0), , , , mstr药品价格等级, mstr卫材价格等级, mstr普通项目价格等级), gstrDecPrice)
                Else
                    .TextMatrix(i, COLPrice("单价")) = Format(NVL(rsTmp!单价), gstrDecPrice)
                End If
                
                '执行科室
                If lng执行科室ID <> 0 Then
                    .TextMatrix(i, COLPrice("执行科室")) = Sys.RowValue("部门表", lng执行科室ID, "名称")
                End If
                
                '显示医保费用类型
                If Val(rsTmp!收费细目ID & "") <> 0 Then
                    strPriceType = GetPriceType(mlng病人ID, Val(rsTmp!收费细目ID & ""), mint险类, mlng病人性质 = 1)
                End If
                '费用类型
                If strPriceType = "" Then
                    .TextMatrix(i, COLPrice("费用类型")) = NVL(rs收费细目!费用类型)
                Else
                    .TextMatrix(i, COLPrice("费用类型")) = strPriceType
                End If

                
                '从属项目
                .TextMatrix(i, COLPrice("从项")) = IIF(NVL(rsTmp!从项, 0) = 0, "", "√")
                
                '收费方式
                .TextMatrix(i, COLPrice("收费方式")) = getChargeMode(Val(NVL(rsTmp!收费方式, 0)))
                
                dblPrice = dblPrice + Format(Val(.TextMatrix(i, COLPrice("计价数量"))) * Val(.TextMatrix(i, COLPrice("单价"))), "0.00000")
                
                rsTmp.MoveNext
            Next
        End If
        
        '合计行
        If .Rows > 2 Then
            .MergeCol(COLPrice("计价医嘱")) = True
            .MergeCol(COLPrice("类别")) = True
            
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, COLPrice("计价医嘱"), .Rows - 1, COLPrice("单位")) = "合计"
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("计价医嘱"), .Rows - 1, COLPrice("单位")) = 4
            .Cell(flexcpText, .Rows - 1, COLPrice("计价数量"), .Rows - 1, COLPrice("单价")) = Format(dblPrice, gstrDecPrice)
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("计价数量"), .Rows - 1, COLPrice("单价")) = 7
            .MergeCells = flexMergeRestrictRows
            .MergeRow(.Rows - 1) = True
            
        End If
        
        .Row = 1: .Col = 0
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    ShowPrice = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowSendList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的发送记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln配方行 As Boolean, bln检验行 As Boolean
    Dim strKey As String, lngKey As Long
    Dim rs执行 As ADODB.Recordset
    Dim str发送号 As String, strTab As String
    Dim bln状态说明 As Boolean
    Dim lng输血 As Long
    Dim j As Long
    
    On Error GoTo errH
        lng输血 = -1
    With vsAppend
        '记录原定位行
        lngKey = -1
        If .Row >= .FixedRows Then
            strKey = .TextMatrix(.Row, COLSend("发送号")) & "," & .TextMatrix(.Row, COLSend("医嘱ID")) & "," & .TextMatrix(.Row, COLSend("收费项目"))
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_期效) = "长嘱" And 1 <> mlng病人性质 Then
            strTab = "住院费用记录"
        Else
            If GetAdviceFeeKind(Val(vsAdvice.TextMatrix(lngRow, COL_ID))) = 2 Then  '住院医生站的临嘱可发送到门诊
                strTab = "住院费用记录"
            Else
                strTab = "门诊费用记录"
            End If
        End If
    
        .Redraw = False
        If .FixedRows = 2 Then .RemoveItem 0
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If mbln叮嘱发送执行 And Val(vsAdvice.TextMatrix(lngRow, COL_医嘱状态)) = 4 And Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID)) = 0 Then
            Call SetExecShow(False, mblnShowExec)
           .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "E" Then
            bln配方行 = RowIs配方行(lngRow)
            bln检验行 = RowIs检验行(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.执行状态,0),0,'未执行',1,'完全执行',2,'部份执行')"
        strExe2 = "Decode(Nvl(B.执行状态,0),0,'未执行',1,'执行完成',2,'拒绝执行',3,'正在执行')"
        strState = "Decode(A.执行状态,9,'收费异常',Decode(A.记录性质,1,Decode(A.记录状态,0,'收费划价',1,'已收费',3,'已退费'),2,Decode(A.记录状态,0,'记帐划价',1,'已记帐',3,'已销帐'),'未计费'))"
        
        '药嘱对应的药品计价按住院包装显示,非药嘱对应的药品计价按零售单位显示
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            If Not RowIn一并给药(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '成药部份:填写了发送记录,但可能无对应费用(如自备药,但医嘱有规格)
            strSub = "Select A.*,B.住院包装,B.住院单位" & _
                " From " & strTab & " A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL And A.收费类别 IN('5','6','7')" & _
                " And A.收费细目ID=B.药品ID And A.医嘱序号=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, strTab, "H" & strTab)
            ElseIf zlDatabase.DateMoved(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, strTab, "H" & strTab)
            End If
                
            strSQL = _
                " Select B.医嘱ID,C.相关ID,C.标本部位,C.检查方法,B.发送时间,B.NO,B.记录性质,B.门诊记帐,A.收费细目ID," & _
                " Nvl(A.住院单位,D.住院单位) as 单位," & _
                " Nvl(A.数次/Nvl(A.住院包装,1),B.发送数次/Nvl(D.剂量系数,1)/Nvl(D.住院包装,1)) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID,Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态," & _
                " B.首次时间,B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费'," & strState & ") as 计费状态," & _
                " B.发送人,B.状态说明,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别," & _
                " B.执行状态 as 执行状态ID,B.计费状态 as 计费状态ID,A.记录状态,NVL(B.完成时间,A.执行时间) as 完成时间,NVL(B.完成人,A.执行人) as 完成人,B.执行说明,B.接收时间,B.接收人,B.报到时间" & _
                " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,药品规格 D" & _
                " Where B.医嘱ID=C.ID And C.收费细目ID=D.药品ID(+)" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And A.医嘱序号(+)=B.医嘱ID" & _
                " And C.ID=[1]"

            '在一并给药的首行才显示给药途径的发送
            If lngRow = lngBegin Then
                '给药途径部份:填写了发送记录(叮嘱无),但不一定有费用
                strSub = "Select A.*,B.住院包装,B.住院单位" & _
                    " From " & strTab & " A,药品规格 B" & _
                    " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                    " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[2]"
                If mblnMoved Then
                    strSub = Replace(strSub, strTab, "H" & strTab)
                ElseIf zlDatabase.DateMoved(mvInDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, strTab, "H" & strTab)
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select B.医嘱ID,C.相关ID,C.标本部位,C.检查方法,B.发送时间,B.NO,B.记录性质,B.门诊记帐,A.收费细目ID," & _
                    " Decode(Nvl(Instr('567',A.收费类别),0),0,Decode(A.收费类别,'4',F.计算单位,D.计算单位),Nvl(A.住院单位,E.住院单位)) as 单位," & _
                    " Nvl(A.数次/Nvl(A.住院包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.住院包装,1)) as 发送数次," & _
                    " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID,Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态," & _
                    " B.首次时间,B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费'," & strState & ") as 计费状态," & _
                    " B.发送人,B.状态说明,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别," & _
                    " B.执行状态 as 执行状态ID,B.计费状态 as 计费状态ID,A.记录状态 ,NVL(B.完成时间,A.执行时间) as 完成时间,NVL(B.完成人,A.执行人) as 完成人,B.执行说明,B.接收时间,B.接收人,B.报到时间" & _
                    " From (" & strSub & ") A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D,药品规格 E,收费项目目录 F" & _
                    " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID And C.收费细目ID=E.药品ID(+) And C.收费细目ID=F.ID(+)" & _
                    " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID" & _
                    " And C.ID=[2]"
            End If
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        Else
            '其它医嘱(包括卫材、配方及检查，手术一组医嘱):填写了发送记录(叮嘱无),但不一定有费用
            '中药自备药也是无对应费用(但医嘱有规格)
            strSub = _
                " Select A.*,B.住院包装,B.住院单位" & _
                " From " & strTab & " A,药品规格 B" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=[1]"
            strSub = strSub & " Union ALL " & _
                " Select A.*,B.住院包装,B.住院单位" & _
                " From " & strTab & " A,药品规格 B,病人医嘱记录 C" & _
                " Where A.记录状态 IN(0,1,3) And A.价格父号 is NULL" & _
                " And A.收费细目ID=B.药品ID(+) And A.医嘱序号=C.ID" & _
                " And C.相关ID=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, strTab, "H" & strTab)
            ElseIf zlDatabase.DateMoved(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, strTab, "H" & strTab)
            End If
            
            strSQL = _
                " Select * From 病人医嘱记录 Where ID=[1]" & _
                " Union ALL " & _
                " Select * From 病人医嘱记录 Where 相关ID=[1]"
            strSQL = _
                " Select B.医嘱ID,C.医嘱内容,C.相关ID,C.标本部位,C.检查方法,B.发送时间,B.NO,B.记录性质,B.门诊记帐,A.收费细目ID," & _
                " Decode(Nvl(Instr('567',A.收费类别),0),0,Decode(A.收费类别,'4',F.计算单位,D.计算单位),Nvl(A.住院单位,E.住院单位)) as 单位," & _
                " Nvl(Nvl(A.付数,1)*A.数次/Nvl(A.住院包装,1),B.发送数次/Nvl(E.剂量系数,1)/Nvl(E.住院包装,1)) as 发送数次," & _
                " Nvl(A.执行部门ID,B.执行部门ID) as 执行部门ID,Decode(Nvl(Instr(',4,5,6,7,',A.收费类别),0),0," & strExe2 & "," & strExe1 & ") as 执行状态," & _
                " B.首次时间,B.末次时间,Decode(Nvl(B.计费状态,0),-1,'无需计费',0,'未计费'," & strState & ") as 计费状态," & _
                " B.发送人,B.状态说明,B.发送号,B.记录序号 as 发送序号,A.序号 as 费用序号,C.诊疗项目ID,C.诊疗类别," & _
                " B.执行状态 as 执行状态ID,B.计费状态 as 计费状态ID,A.记录状态,B.完成时间,B.完成人,B.执行说明,B.接收时间,B.接收人,B.报到时间" & _
                " From (" & strSub & ") A,病人医嘱发送 B,(" & strSQL & ") C,诊疗项目目录 D,药品规格 E,收费项目目录 F" & _
                " Where B.医嘱ID=C.ID And C.诊疗项目ID=D.ID" & IIF(mbln叮嘱发送执行, "(+)", "") & " And C.收费细目ID=E.药品ID(+) And C.收费细目ID=F.ID(+)" & _
                " And A.NO(+)=B.NO And A.记录性质(+)=B.记录性质 And 0+A.医嘱序号(+)=B.医嘱ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
                strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
            End If
        End If
        
        strSQL = "Select  A.发送序号,A.费用序号," & _
            " A.医嘱ID,A.相关ID,A.诊疗类别,F.名称 as 类别名称," & IIF(mbln叮嘱发送执行, IIF(InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0, "D.名称", "Nvl(d.名称, a.医嘱内容)"), "D.名称") & " as 诊疗项目,A.标本部位,A.检查方法,A.发送时间,A.NO,A.记录性质,A.门诊记帐," & _
            " Nvl(G.名称,B.名称)||Decode(B.产地,NULL,NULL,'('||B.产地||')')||Decode(B.规格,NULL,NULL,' '||B.规格) as 收费项目," & _
            " A.单位,A.发送数次 as 数量,C.名称 as 执行科室,A.执行状态,A.首次时间,A.末次时间,A.计费状态,A.发送人,A.状态说明,A.发送号," & _
            " A.执行部门ID,A.执行状态ID,A.计费状态ID,A.记录状态,D.操作类型,H.跟踪在用,A.完成时间,a.完成人,a.执行说明,a.接收时间,a.接收人,a.报到时间" & _
            " From (" & strSQL & ") A,收费项目目录 B,部门表 C,诊疗项目目录 D,诊疗项目类别 F,收费项目别名 G,材料特性 H" & _
            " Where A.收费细目ID=B.ID(+) And A.执行部门ID=C.ID(+) And A.诊疗项目ID=D.ID" & IIF(mbln叮嘱发送执行, "(+)", "") & " And A.诊疗类别=F.编码(+)" & _
            " And A.收费细目ID=H.材料ID(+) And A.收费细目ID=G.收费细目ID(+) And G.码类(+)=1 And G.性质(+)=" & IIF(gbyt药品名称显示 = 0, 1, 3) & _
            " Order by A.发送号 Desc,A.诊疗类别,A.发送序号,A.费用序号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)))
        
        If Not rsTmp.EOF Then
            strSQL = "Select Max(a.执行时间) As 执行时间, a.医嘱id, a.发送号 From 病人医嘱执行 A, 病人医嘱发送 B, 病人医嘱记录 C" & vbNewLine & _
                        "Where a.医嘱id = b.医嘱id And a.发送号 = b.发送号 And b.医嘱id=c.id and (c.id=[1] or c.相关id=[1])" & vbNewLine & _
                        "Group By a.医嘱id, a.发送号"
            Set rs执行 = zlDatabase.OpenSQLRecord(strSQL, Me.Name, IIF(Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)) = 0, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))))
            
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                If InStr(str发送号 & ",", "," & NVL(rsTmp!发送号, 0) & ",") = 0 Then
                    str发送号 = str发送号 & "," & NVL(rsTmp!发送号, 0)
                End If
                .TextMatrix(i, COLSend("发送号")) = NVL(rsTmp!发送号, 0)
                .TextMatrix(i, COLSend("发送时间")) = Format(NVL(rsTmp!发送时间), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COLSend("发送时间")) = Format(NVL(rsTmp!发送时间), "yyyy-MM-dd HH:mm:ss")
                
                '发送医嘱
                If rsTmp!诊疗类别 = "4" Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "卫生材料-" & rsTmp!诊疗项目
                ElseIf InStr(",5,6,7,", rsTmp!诊疗类别) > 0 Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "药品医嘱-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "给药途径-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" Then
                    .TextMatrix(i, COLSend("发送医嘱")) = "输血途径-" & rsTmp!诊疗项目
                ElseIf rsTmp!诊疗类别 = "E" And (bln配方行 Or bln检验行) Then
                    If bln检验行 Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "采集方法-" & rsTmp!诊疗项目
                    ElseIf Not IsNull(rsTmp!相关ID) Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "中药煎法-" & rsTmp!诊疗项目
                    Else
                        .TextMatrix(i, COLSend("发送医嘱")) = "中药用法-" & rsTmp!诊疗项目
                    End If
                ElseIf Not IsNull(rsTmp!相关ID) Then
                    If rsTmp!诊疗类别 = "C" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "检验项目-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "D" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "检查部位-" & NVL(rsTmp!标本部位) & "(" & NVL(rsTmp!检查方法) & ")"
                    ElseIf rsTmp!诊疗类别 = "F" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "附加手术-" & rsTmp!诊疗项目
                    ElseIf rsTmp!诊疗类别 = "G" Then
                        .TextMatrix(i, COLSend("发送医嘱")) = "麻醉项目-" & rsTmp!诊疗项目
                    End If
                Else
                    .TextMatrix(i, COLSend("发送医嘱")) = rsTmp!类别名称 & "医嘱-" & rsTmp!诊疗项目
                End If
               
                .TextMatrix(i, COLSend("单据号")) = NVL(rsTmp!NO)
                .TextMatrix(i, COLSend("收费项目")) = NVL(rsTmp!收费项目)
                .TextMatrix(i, COLSend("发送数次")) = FormatEx(NVL(rsTmp!数量), 5) & NVL(rsTmp!单位)
                .TextMatrix(i, COLSend("计费状态")) = NVL(rsTmp!计费状态)
                .TextMatrix(i, COLSend("执行状态")) = NVL(rsTmp!执行状态)
                .TextMatrix(i, COLSend("执行科室")) = NVL(rsTmp!执行科室)
                .TextMatrix(i, COLSend("首次时间")) = Format(NVL(rsTmp!首次时间), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("末次时间")) = Format(NVL(rsTmp!末次时间), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("接受时间")) = Format(NVL(rsTmp!接收时间), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("接受人")) = NVL(rsTmp!接收人)
                .TextMatrix(i, COLSend("到场时间")) = Format(NVL(rsTmp!报到时间), "yyyy-MM-dd HH:mm")
                
                .TextMatrix(i, COLSend("发送人")) = NVL(rsTmp!发送人)
                .TextMatrix(i, COLSend("状态说明")) = NVL(rsTmp!状态说明)
                If rsTmp!状态说明 & "" <> "" Then
                    bln状态说明 = True
                End If
                '隐藏列,用于执行处理
                .TextMatrix(i, COLSend("医嘱ID")) = rsTmp!医嘱ID
                .TextMatrix(i, COLSend("相关ID")) = NVL(rsTmp!相关ID)
                .TextMatrix(i, COLSend("记录性质")) = NVL(rsTmp!记录性质, 0)
                .TextMatrix(i, COLSend("门诊记帐")) = Val("" & rsTmp!门诊记帐)
                .TextMatrix(i, COLSend("记录状态")) = NVL(rsTmp!记录状态, 0)
                .TextMatrix(i, COLSend("诊疗类别")) = NVL(rsTmp!诊疗类别)
                .TextMatrix(i, COLSend("操作类型")) = NVL(rsTmp!操作类型)
                .TextMatrix(i, COLSend("跟踪在用")) = NVL(rsTmp!跟踪在用, 0)
                .TextMatrix(i, COLSend("完成时间")) = Format(NVL(rsTmp!完成时间), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("执行时间")) = Format(NVL(rsTmp!完成时间), "yyyy-MM-dd HH:mm")
                rs执行.Filter = "医嘱ID=" & rsTmp!医嘱ID & " And 发送号=" & NVL(rsTmp!发送号, 0)
                If Not rs执行.EOF Then .TextMatrix(i, COLSend("最后执行时间")) = Format(NVL(rs执行!执行时间), "yyyy-MM-dd HH:mm")
                
                .TextMatrix(i, COLSend("执行人")) = NVL(rsTmp!完成人)
                .TextMatrix(i, COLSend("执行说明")) = NVL(rsTmp!执行说明)
                .TextMatrix(i, COLSend("输血类型")) = Val(NVL(rsTmp!检查方法))  '输血类医嘱：检查方法存储，0-备血，1-用血
                .Cell(flexcpData, i, COLSend("计费状态")) = CStr(rsTmp!计费状态ID)
                .Cell(flexcpData, i, COLSend("执行状态")) = Val(NVL(rsTmp!执行状态ID, 0))
                .Cell(flexcpData, i, COLSend("执行科室")) = Val("" & rsTmp!执行部门ID)
                
                '定位原先行
                If NVL(rsTmp!发送号, 0) & "," & NVL(rsTmp!医嘱ID) & "," & NVL(rsTmp!收费项目) = strKey Then
                    lngKey = i
                End If
                
                If Val("" & rsTmp!执行部门ID) = mlng病区ID Or Val("" & rsTmp!执行部门ID) = mlng科室ID Or vsAdvice.TextMatrix(vsAdvice.Row, COL_执行性质) = "离院带药" Then
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = &HD0FFFF
                    If lngKey = -1 Then
                        '如果不定位到以前的行，则自动定位到本病区或科室执行的项目
                        lngKey = i
                    End If
                End If
                If vsAdvice.TextMatrix(lngRow, COL_诊疗类别) = "K" And rsTmp!诊疗类别 & "" = "K" Then
                    If gbln血库系统 Then
                        lng输血 = i
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        
        If lng输血 <> -1 Then
            '输血医嘱其它诊疗项目的信息
            strSQL = "select b.名称 as 诊疗项目,a.申请量 as 数量,b.计算单位 as 单位,a.诊疗项目id from 输血申请项目 a,诊疗项目目录 b where a.诊疗项目id=b.id and a.医嘱id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!诊疗项目ID & "") <> Val(vsAdvice.TextMatrix(lngRow, COL_诊疗项目ID)) Then
                    .AddItem ""
                    For j = .FixedCols To .Cols - 1
                        .TextMatrix(.Rows - 1, j) = .TextMatrix(lng输血, j)
                    Next
                    .Cell(flexcpData, .Rows - 1, COLSend("发送时间")) = .Cell(flexcpData, lng输血, COLSend("发送时间"))
                    .Cell(flexcpData, .Rows - 1, COLSend("计费状态")) = .Cell(flexcpData, lng输血, COLSend("计费状态"))
                    .Cell(flexcpData, .Rows - 1, COLSend("执行状态")) = .Cell(flexcpData, lng输血, COLSend("执行状态"))
                    .Cell(flexcpData, .Rows - 1, COLSend("执行科室")) = .Cell(flexcpData, lng输血, COLSend("执行科室"))
                    .TextMatrix(.Rows - 1, COLSend("发送医嘱")) = "输血医嘱-" & rsTmp!诊疗项目
                    .TextMatrix(.Rows - 1, COLSend("发送数次")) = FormatEx(NVL(rsTmp!数量), 5) & NVL(rsTmp!单位)
                    .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, lng输血, .FixedCols)
                Else
                    .TextMatrix(lng输血, COLSend("发送数次")) = FormatEx(NVL(rsTmp!数量), 5) & NVL(rsTmp!单位)
                End If
                rsTmp.MoveNext
            Next
        End If
        
        If str发送号 <> "" Then
            .AddItem "", 0
            .FixedRows = 2
            .Cell(flexcpText, 0, 0, 0, .Cols - 1) = " 共发送 " & UBound(Split(str发送号, ",")) & " 次"
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
            .MergeRow(0) = True
            
            .Row = IIF(lngKey = -1, .FixedRows, lngKey + 1): .Col = COLSend("发送医嘱")
        Else
            .Row = IIF(lngKey = -1, .FixedRows, lngKey): .Col = COLSend("发送医嘱")
        End If
        .MergeCells = flexMergeFree
        .MergeCol(COLSend("发送号")) = True
        .MergeCol(COLSend("发送时间")) = True
        .MergeCol(COLSend("单据号")) = True
        .MergeCol(COLSend("发送医嘱")) = True
        .MergeCol(COLSend("收费项目")) = True
        .MergeCol(COLSend("首次时间")) = True
        .MergeCol(COLSend("末次时间")) = True
        .MergeCol(COLSend("发送人")) = True
        .MergeCol(COLSend("状态说明")) = True
        
        .ColHidden(COLSend("状态说明")) = Not bln状态说明
        .Redraw = True
        Call vsAppend_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    ShowSendList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadExecList(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long) As Boolean
'功能：读取指定医嘱的执行情况表
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strPre As String
    Dim rs血库 As ADODB.Recordset
    Dim bln输血 As Boolean
    Dim int血袋数 As Integer
    
    On Error GoTo errH
    
    '检验项目一并执行时，执行情况登记到第一个项目上。分散单独执行时，登记到各个项目上。
    strSQL = "Select A.要求时间,A.执行时间,A.本次数次,D.计算单位,A.执行摘要,A.执行人,A.登记时间,A.登记人,DECODE(NVL(A.执行结果,1),0,'未执行',1,'完成',2,'拒绝',3,'外出') As 执行结果,a.核对人,a.核对时间,d.操作类型,d.类别,a.说明,a.记录来源 as 来源" & _
        " From 病人医嘱执行 A,病人医嘱发送 B,病人医嘱记录 C,诊疗项目目录 D" & _
        " Where A.医嘱ID=[1] And A.发送号=[2]" & _
        " And A.医嘱ID=B.医嘱ID And A.发送号=B.发送号 And B.医嘱ID=C.ID And C.诊疗项目ID=D.ID" & IIF(mbln叮嘱发送执行, "(+)", "") & _
        " Order by A.登记时间 Desc"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱执行", "H病人医嘱执行")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, lng发送号)
    With vsExec
        strPre = .Cell(flexcpData, .Row, 0)
        .Redraw = flexRDNone
        .Rows = vsExec.FixedRows
        .Rows = vsExec.FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            '输血医嘱处理流程变动 70823
            If gbln血库系统 And Val(rsTmp!操作类型 & "") = 8 And rsTmp!类别 = "E" Then
                strSQL = "select zl_Get_输血执行次数(相关id) as 数量 from 病人医嘱记录 where id = [1]"
                Set rs血库 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
                If Not rs血库.EOF Then int血袋数 = Val(rs血库!数量 & "")
                bln输血 = True
            End If
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Format(rsTmp!要求时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 1) = Format(rsTmp!执行时间, "yyyy-MM-dd HH:mm")
                If bln输血 Then
                    .TextMatrix(i, 2) = FormatEx(Val(rsTmp!本次数次 & "") * int血袋数, 0) & " 袋"
                Else
                    .TextMatrix(i, 2) = FormatEx(rsTmp!本次数次, 5) & " " & NVL(rsTmp!计算单位)
                End If
                .TextMatrix(i, 3) = NVL(rsTmp!执行摘要)
                .TextMatrix(i, 4) = NVL(rsTmp!执行人)
                .TextMatrix(i, 5) = Format(rsTmp!登记时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 6) = NVL(rsTmp!登记人)
                .TextMatrix(i, 7) = rsTmp!执行结果 & ""
                .TextMatrix(i, 8) = NVL(rsTmp!核对人)
                .TextMatrix(i, 9) = Format(rsTmp!核对时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 10) = NVL(rsTmp!说明)
                .TextMatrix(i, 11) = IIF(1 = Val(rsTmp!来源 & ""), "移动端", "PC端")
                
                .Cell(flexcpData, i, 0) = Format(rsTmp!要求时间, "yyyy-MM-dd HH:mm:ss")
                .Cell(flexcpData, i, 1) = Format(rsTmp!执行时间, "yyyy-MM-dd HH:mm:ss")
       
                If .Cell(flexcpData, i, 0) = strPre Then .Row = i
                rsTmp.MoveNext
            Next
            rsTmp.MoveFirst
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
    End With
    With vsAppend
        If Not (.TextMatrix(.Row, COLSend("诊疗类别")) = "E" And .TextMatrix(.Row, COLSend("操作类型")) = "1" And Mid(gstr医嘱核对, 2, 1) = "1" Or _
            (.TextMatrix(.Row, COLSend("诊疗类别")) = "E" And .TextMatrix(.Row, COLSend("操作类型")) = "8" Or .TextMatrix(.Row, COLSend("诊疗类别")) = "K") And Mid(gstr医嘱核对, 1, 1) = "1") Then
            
            vsExec.ColHidden(8) = True
            vsExec.ColHidden(9) = True
        Else
            vsExec.ColHidden(8) = False
            vsExec.ColHidden(9) = False
        End If
    End With
    LoadExecList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowSignList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的签名记录
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSignList = True: Exit Function
        End If
        
        strSQL = "Select A.签名ID,A.操作类型,B.签名时间,B.签名人,B.时间戳," & _
            " Decode(A.操作类型,1,'新开医嘱',3,'校对医嘱',4,'作废医嘱',8,'停止医嘱','其它操作') as 签名类型" & _
            " From 病人医嘱状态 A,医嘱签名记录 B Where A.医嘱ID=[1] And A.签名ID=B.ID Order by B.签名时间"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱状态", "H病人医嘱状态")
            strSQL = Replace(strSQL, "医嘱签名记录", "H医嘱签名记录")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!签名ID)
                .TextMatrix(i, 0) = rsTmp!签名类型
                .Cell(flexcpData, i, 0) = Val(rsTmp!操作类型)
                .TextMatrix(i, 1) = Format(rsTmp!签名时间, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!签名人
                .TextMatrix(i, 3) = Format(NVL(rsTmp!时间戳), "yyyy-MM-dd HH:mm:ss")
                Set .Cell(flexcpPicture, i, 0) = frmIcons.imgSign.ListImages("签名").Picture
                rsTmp.MoveNext
            Next
        End If
        .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, 0) = 0
        .Row = 1
        .Redraw = True
    End With
    ShowSignList = True
    Exit Function
errH:
    vsAppend.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBillAppend(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'功能：显示指定行医嘱的单据附项内容
'返回：blnExist=医嘱是否存在单据附项内容
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    
    blnExist = False
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    strSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order by 排列"
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱附件", "H病人医嘱附件")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp!项目 & "：" & NVL(rsTmp!内容)
                lngIdx = .Find(rsTmp!项目 & "：", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp!项目 & "：")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            '光标定位在第一个输入附项
            rsTmp.MoveFirst
            lngIdx = .Find(rsTmp!项目 & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp!项目 & "：")
            
            Call SetRTFFont(1)
        End With
        blnExist = True
    End If
    
    ShowBillAppend = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub SetRTFFont(bytKind As Byte)
    If bytKind = 0 Or bytKind = 1 Then
        With rtfAppend
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 2 Then
        With rtfInfo
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 3 Then
        With rtfOther
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
    If bytKind = 0 Or bytKind = 4 Then
        With rtfSche
            .SelStart = 0
            .SelLength = Len(.Text)
            .SelFontSize = mlngFontSize
            .SelLength = 0
        End With
    End If
End Sub

Private Function ShowAdvicePlan(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'功能：显示指定行医嘱的执行安排信息
'返回：blnExist=医嘱是否存在执行安排信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    blnExist = False
    rtfInfo.Text = "": rtfInfo.SelStart = 0
    
    On Error GoTo errH
    
    With vsAdvice
        If InStr("D,F,G,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Or _
            .TextMatrix(lngRow, COL_诊疗类别) = "E" And InStr(",0,6,", "," & .TextMatrix(lngRow, COL_操作类型) & ",") > 0 Then
            
            If .TextMatrix(lngRow, COL_诊疗类别) = "E" And Val(.TextMatrix(lngRow, COL_操作类型)) = 6 Then
                strSQL = "Select a.安排时间,a.执行间,a.执行说明 From 病人医嘱发送 a,病人医嘱记录 b " & _
                        "Where a.医嘱ID = b.ID And b.相关ID=[1] And (a.执行说明 is Not Null Or a.安排时间 is Not Null) And Rownum=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            Else
                strSQL = "Select 安排时间,执行间,执行说明 From 病人医嘱发送 Where 医嘱ID=[1] And (执行说明 is Not Null Or 安排时间 is Not Null)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            End If
            If Not rsTmp.EOF Then
                strSQL = ""
                
                If Not IsNull(rsTmp!安排时间) Then
                    strSQL = strSQL & vbCrLf & "安排时间：" & Format(rsTmp!安排时间, "yyyy-MM-dd HH:mm")
                End If
                If Not IsNull(rsTmp!执行间) Then
                    strSQL = strSQL & vbCrLf & "执行间：" & rsTmp!执行间
                End If
                strSQL = strSQL & vbCrLf & NVL(rsTmp!执行说明)
                
                rtfInfo.Text = Mid(strSQL, 3)
                
                Call SetRTFFont(2)
                blnExist = True
            End If
        End If
    End With
    ShowAdvicePlan = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowOtherAppend(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱的审核信息
'说明：只检查审核状态通过和未通过的医嘱
'返回：是否存在审核信息
    Dim strSQL As String
    Dim int类型 As Integer
    Dim rsTmp As ADODB.Recordset
    Dim str操作员 As String
    Dim str时间 As String
    Dim str未用原因 As String '输血医嘱特有
    
    On Error GoTo errH

    str操作员 = "审核人：": str时间 = "审核时间："
    With vsAdvice
        If gbln血库系统 And .TextMatrix(lngRow, COL_诊疗类别) = "K" Then
            If Val(.TextMatrix(lngRow, COL_执行标记)) = -1 Then '读取标记未用的原因
                strSQL = "Select 操作人员,操作时间,操作说明 From 病人医嘱状态 Where 医嘱id = [1] And 操作类型 = [2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), 17)
                If Not rsTmp.EOF Then
                    str未用原因 = "未用原因：" & rsTmp!操作说明
                    str未用原因 = str未用原因 & "(操作员：" & rsTmp!操作人员 & "  操作时间：" & Format(rsTmp!操作时间 & "", "YYYY-MM-DD HH:MM:SS") & ")"
                End If
            End If
        End If
        
        Select Case .TextMatrix(lngRow, COL_审核状态)
            Case 2
                If gbln血库系统 And .TextMatrix(lngRow, COL_诊疗类别) = "K" Then  '输血医嘱处理流程变动 70823
                    int类型 = 15 '血库审核通过
                    str操作员 = "血库审核人："
                    str时间 = "血库审核时间："
                Else
                    int类型 = 11
                End If
            Case 3
                int类型 = 12
            Case 4
                int类型 = 11
            Case 5
                int类型 = 14
                str操作员 = "血库接收人："
                str时间 = "血库接收时间："
        End Select
        rtfOther.Text = ""
        strSQL = "Select 操作人员,操作时间 From 病人医嘱状态 Where 医嘱id = [1] And 操作类型 = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), int类型)
    End With
    
    If Not rsTmp.EOF Then
        strSQL = ""
        Do While Not rsTmp.EOF
            strSQL = IIF(strSQL = "", "", strSQL & vbCrLf) & str操作员 & rsTmp!操作人员 & vbCrLf & _
                str时间 & Format(rsTmp!操作时间 & "", "YYYY-MM-DD HH:MM:SS")
            rsTmp.MoveNext
        Loop
        If str未用原因 <> "" Then
            strSQL = IIF(strSQL = "", "", strSQL & vbCrLf) & str未用原因
        End If
        rtfOther.Text = strSQL
        Call SetRTFFont(3)
        ShowOtherAppend = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ShowCompoundInfo(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'功能：显示指定行的输液配药内容
'返回：blnExist=医嘱是否存在发送到输液配置中心的药品
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    blnExist = False
    On Error GoTo errH
    
    '只有护士站才调用本函数
    If gstr输液配置中心 <> "" Then
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
            strSQL = "Select 1 From 输液配药记录 Where 医嘱id = [1] and nvl(操作状态,0)<>12 And Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)))
            If Not rsTmp.EOF Then '非输液类的也允许发送到配置中心，但不产生配药记录，临嘱和长嘱根据输液配中心的参数来决定是否产生配药记录
                blnExist = True
            End If
        End If
    End If
    ShowCompoundInfo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadRollList(ByVal lngRow As Long) As Boolean
'功能：显示指定行医嘱可以回退的内容在内存中
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vRoll As TYPE_AdviceRoll
    Dim lng医嘱ID As Long
    
    ReDim marrRollList(0)
    If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
        LoadRollList = True: Exit Function
    End If
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_相关ID)) <> 0 Then
        lng医嘱ID = Val(vsAdvice.TextMatrix(lngRow, COL_相关ID))
    Else
        lng医嘱ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
    End If
    
    '可回退医嘱操作和发送,医嘱本身的操作优先(如发送后自动停止)
    '临嘱不可回退自动停止,回退发送时自动回退停止
    strSQL = " And (A.ID=[1] Or A.相关ID=[1])"
    strSQL = _
        " Select Distinct 0 as 发送号,B.操作人员 as 人员,B.操作时间 as 时间,B.操作类型," & _
        " Decode(B.操作类型,4,'作废医嘱',5,'重整医嘱',6,'暂停医嘱',7,'启用医嘱',8,'停止医嘱',9,'确认停止',10,'皮试结果',13,'停嘱申请') as 内容" & _
        " From 病人医嘱记录 A,病人医嘱状态 B" & _
        " Where A.ID=B.医嘱ID" & strSQL & _
        " And (Nvl(A.医嘱期效,0)=0 And B.操作类型 Not IN(1,2,3)" & _
            " Or Nvl(A.医嘱期效,0)=1 And B.操作类型 Not IN(1,2,3,8))" & _
        " Union ALL" & _
        " Select Distinct B.发送号,B.发送人 as 人员,B.发送时间 as 时间,0 as 操作类型,'发送医嘱' as 内容" & _
        " From 病人医嘱记录 A,病人医嘱发送 B" & _
        " Where A.ID=B.医嘱ID" & strSQL & _
        " Order by 时间 Desc,发送号"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng医嘱ID)
    If Not rsTmp.EOF Then
        ReDim marrRollList(rsTmp.RecordCount)
        For i = 1 To rsTmp.RecordCount
            With vRoll
                .操作类型 = rsTmp!操作类型
                .发送号 = rsTmp!发送号
                .操作时间 = Format(rsTmp!时间, "yyyy-MM-dd HH:mm:ss")
                .操作人员 = rsTmp!人员
                .操作内容 = "操作人:" & rsTmp!人员 & ",时间:" & Format(rsTmp!时间, "yyyy-MM-dd HH:mm") & ",内容:" & rsTmp!内容
            End With
            marrRollList(i) = vRoll '第0的个不算
            rsTmp.MoveNext
        Next
    End If
    
    LoadRollList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RollFirstEnabled() As Boolean
'功能：判断第一个回退项是否可以回退
    Dim vRoll As TYPE_AdviceRoll
    Dim blnEnabled As Boolean
    
    If UBound(marrRollList) < 1 Then Exit Function
    vRoll = marrRollList(1)
    
    '出院或已诊病人不允许回退操作
    If mintPState = ps出院 Or mintPState = ps已诊 Then Exit Function
    
    '预出院病人仅可以回退出院医嘱发送
    If mintPState = ps预出 Then
        blnEnabled = False
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "Z" _
            And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_操作类型))) > 0 Then
            If vRoll.操作类型 = 0 And vRoll.发送号 <> 0 Then
                blnEnabled = True
            End If
        End If
        If Not blnEnabled Then Exit Function
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗项目ID)) = 0 And InStr(vRoll.操作内容, "发送医嘱") > 0 Then
        Exit Function
    End If
    
    '医生只能回退自已的作废、停止,暂停、启用,重整；发送操作
    If mint场合 <> 1 Then
        If Not ((vRoll.操作类型 = 0 Or InStr("45678", vRoll.操作类型) > 0 Or vRoll.操作类型 = 13) And vRoll.操作人员 = UserInfo.姓名) Then
            Exit Function
        ElseIf mint场合 = 2 Then
            If InStr("," & mstr前提IDs & ",", "," & vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID) & ",") = 0 Or vRoll.操作人员 <> UserInfo.姓名 Then
                 Exit Function
            End If
        ElseIf mint场合 = 0 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_前提ID)) <> 0 Then Exit Function
        End If
    End If
    
    RollFirstEnabled = True
End Function

Private Function LoadBillList() As Boolean
'功能：显示指定行的医嘱发送可以打印的诊疗单据在菜单上
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objMenu As CommandBarPopup
    Dim strTmp As String
    Dim blnBlood As Boolean, intBloodState As Integer '新开时打印输血医嘱
    
    If mcbsMain Is Nothing Then LoadBillList = True: Exit Function
    Set objPopup = mcbsMain.FindControl(, conMenu_Report_ClinicBill, False, True)
    If objPopup Is Nothing Then LoadBillList = True: Exit Function
    objPopup.Visible = True
    
    objPopup.CommandBar.Controls.DeleteAll
    If mcbsMain Is Nothing Then LoadBillList = True: Exit Function
    Set objMenu = mcbsMain.FindControl(, conMenu_EditPopup, False, True)
    If objMenu Is Nothing Then LoadBillList = True: Exit Function
    For i = objMenu.CommandBar.Controls.Count To 1 Step -1
        If objMenu.CommandBar.Controls(i).ID > conMenu_Report_ClinicBill * 100# And objMenu.CommandBar.Controls(i).ID < conMenu_Report_ClinicBill * 100# + 100 Then
            objMenu.CommandBar.Controls(i).Delete
        End If
    Next
    
    '输血医嘱申请单打印模式=0，新开后可以打印
    intBloodState = 8
    If mint申请单打印模式 = 0 And vsAdvice.TextMatrix(vsAdvice.Row, COL_诊疗类别) = "K" Then
        blnBlood = True
        strTmp = ",-1,4,"
        If InStr(",1,2,3,", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) & ",") > 0 Then
            intBloodState = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态))
        End If
    Else
        strTmp = ",-1,1,2,4,"
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 _
       Or InStr(strTmp, "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_医嘱状态)) & ",") > 0 Then
        LoadBillList = True: Exit Function
    End If

        
    On Error GoTo errH
    strTmp = ""
    With vsAdvice
        If mint申请单打印模式 = 1 Or blnBlood = True Then
            '需要区分是输血申请单还是用血通知单
            If gbln血库系统 = True Then
                strTmp = " Union All " & vbNewLine & _
                " Select '-17', Decode(类别, 1, '输血申请单', '取血通知单') 名称, '', '', 类别" & vbNewLine & _
                " From (Select Decode(a.操作类型, '8', Nvl(a.执行分类, 0), 0)+1 类别" & vbNewLine & _
                "       From 诊疗项目目录 a, 病人医嘱记录 b, 病人医嘱记录 c" & vbNewLine & _
                "       Where Instr(',8,9,', ',' || a.操作类型 || ',') > 0 And a.Id = b.诊疗项目id And b.诊疗类别 = 'E' And b.相关id = c.Id And" & vbNewLine & _
                "             c.Id = [1] And c.申请序号 Is Not Null And c.诊疗类别 = 'K' And c.医嘱状态 = [2])"
            Else
                strTmp = " Union All " & _
                " Select '-17','输血申请单','','',0 From 病人医嘱记录 A Where  a.ID=[1] And A.申请序号 is not null And A.诊疗类别 = 'K' And A.医嘱状态=[2]"
            End If
        End If
        strSQL = "Select Distinct D.编号,D.名称,D.说明,B.NO,0 类别" & _
            " From 病人医嘱记录 A,病人医嘱发送 B,病历单据应用 C,病历文件列表 D" & _
            " Where C.诊疗项目ID=A.诊疗项目ID And a.ID=b.医嘱ID " & _
            " And C.应用场合=2 And C.病历文件ID=D.ID And D.种类=7 And (a.ID=[1] or A.相关ID=[1])" & _
            " And (A.申请序号 is null Or A.诊疗类别 <>'K')" & _
            strTmp & _
            " Order by 编号"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Decode(Val(.TextMatrix(.Row, COL_相关ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_相关ID))), intBloodState)
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
    End With
     '如果只有一个诊疗单据可用，则直接加入到医嘱菜单里
    If rsTmp.RecordCount = 1 Then
        objPopup.Visible = False
        objPopup.Category = "已判断"
        Set objPopup = objMenu
    End If
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, IIF(rsTmp.RecordCount = 1, "打印:", "") & rsTmp!名称)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                '中药的煎法用法单据号和中药不一样，界面上显示的中药用法，所以把单据的NO拼进去
                '如果小于0表示使用病区固定报表
                If Val(rsTmp!编号 & "") < 0 Then
                    objControl.Parameter = "ZL1_INSIDE_1254_" & Abs(Val(rsTmp!编号 & "")) & IIF(Val(rsTmp!类别 & "") = 0, "", "_" & Val(rsTmp!类别 & "")) '对应的自定义报表编号
                Else
                    objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" & "|" & rsTmp!NO '对应的自定义报表编号
                End If
                'If i > 1 Then objControl.Enabled = False '一个项目只能设置一个诊疗单据
            End With
            rsTmp.MoveNext
        Next
    End If
    
    LoadBillList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadBillListOld() As Boolean
'功能：显示指定行的医嘱发送可以打印的诊疗单据在菜单上(发送菜单用)
'      有可能长嘱打印前几次发送的单据，还是要选择发送记录No来打印
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim strTmp As String
    
    If mcbsMain Is Nothing Then LoadBillListOld = True: Exit Function
    Set objPopup = mcbsMain.FindControl(, conMenu_Report_ClinicBill, False, True) '可能故意隐藏了
    If objPopup Is Nothing Then LoadBillListOld = True: Exit Function
    
    objPopup.CommandBar.Controls.DeleteAll
    
    If tbcAppend.Selected.Tag <> "发送" Then LoadBillListOld = True: Exit Function
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 _
        Or Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("发送号"))) = 0 Then
        LoadBillListOld = True: Exit Function
    End If
    
    On Error GoTo errH
    
    With vsAppend
        '排除申请单下达的输血医嘱，单独处理
        '本身就是医嘱发送后才调用，不用再次判断申请单打印模式
'        If mint申请单打印模式 = 1 Then
            If gbln血库系统 = True Then
                strTmp = " Union All " & vbNewLine & _
                " Select '-17', Decode(类别, 1, '输血申请单', '取血通知单') 名称, '', 类别" & vbNewLine & _
                " From (Select Decode(a.操作类型, '8', Nvl(a.执行分类, 0), 0)+1 类别" & vbNewLine & _
                "       From 诊疗项目目录 a, 病人医嘱记录 b, 病人医嘱记录 c" & vbNewLine & _
                "       Where Instr(',8,9,', ',' || a.操作类型 || ',') > 0 And a.Id = b.诊疗项目id And b.诊疗类别 = 'E' And b.相关id = c.Id And" & vbNewLine & _
                "             c.Id = [3] And c.申请序号 Is Not Null And c.诊疗类别 = 'K' And c.医嘱状态 = 8)"
            Else
                strTmp = " Union All " & _
                " Select '-17','输血申请单','',0 From 病人医嘱记录 A Where  a.ID=[3] And A.申请序号 is not null And A.诊疗类别 = 'K' And A.医嘱状态=8"
            End If
'        End If
        strSQL = "Select Distinct D.编号,D.名称,D.说明,0 类别" & _
            " From 病人医嘱发送 A,病人医嘱记录 B,病历单据应用 C,病历文件列表 D" & _
            " Where A.发送号=[1] And A.NO=[2]" & _
            " And A.医嘱ID=B.ID And B.诊疗项目ID=C.诊疗项目ID" & _
            " And C.应用场合=2 And C.病历文件ID=D.ID And D.种类=7" & _
            " And (b.申请序号 is null Or b.诊疗类别 <>'K')" & _
            strTmp & _
            " Order by 编号"
        If mblnMoved Then
            strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
            strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, COLSend("发送号"))), .TextMatrix(.Row, COLSend("单据号")), Val(.TextMatrix(.Row, COLSend("医嘱ID"))))
    End With
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, rsTmp!名称)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                '如果小于0表示使用病区固定报表
                If Val(rsTmp!编号 & "") < 0 Then
                    objControl.Parameter = "ZL1_INSIDE_1254_" & Abs(Val(rsTmp!编号 & "")) & IIF(Val(rsTmp!类别 & "") = 0, "", "_" & Val(rsTmp!类别 & "")) '对应的自定义报表编号
                Else
                    objControl.Parameter = "ZLCISBILL" & Format(rsTmp!编号, "00000") & "-1" '对应的自定义报表编号
                End If
                'If i > 1 Then objControl.Enabled = False '一个项目只能设置一个诊疗单据
            End With
            rsTmp.MoveNext
        Next
    End If
    
    LoadBillListOld = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAppend_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnExec As Boolean
    Dim blnBloldExec As Boolean '输血执行
    If Not Me.Visible Then Exit Sub
    If NewRow = OldRow Then Exit Sub
    
    With vsAppend
        If NewCol >= .FixedCols And NewRow >= .FixedRows Then
            If .Redraw <> flexRDNone Then
                If tbcAppend.Selected.Tag = "发送" And vsAppend.Cols = COLSend.Count Then
                    If mint场合 = 1 And Val(.TextMatrix(NewRow, COLSend("发送号"))) <> 0 And (InStr(",5,6,7,", .TextMatrix(NewRow, COLSend("诊疗类别"))) = 0 Or (mbln叮嘱发送执行 And .TextMatrix(NewRow, COLSend("诊疗类别")) = "")) Then
                        If Not (.TextMatrix(NewRow, COLSend("诊疗类别")) = "Z" And Val(.TextMatrix(NewRow, COLSend("操作类型"))) <> 0) Then
                            If Val(.Cell(flexcpData, NewRow, COLSend("执行科室"))) = mlng病区ID Or Val(.Cell(flexcpData, NewRow, COLSend("执行科室"))) = mlng科室ID Or vsAdvice.TextMatrix(vsAdvice.Row, COL_执行性质) = "离院带药" Then
                                blnExec = True
                            End If
                            
                            If gbln血库系统 And .TextMatrix(NewRow, COLSend("诊疗类别")) = "E" And Val(.TextMatrix(NewRow, COLSend("操作类型"))) = 8 Then
                                '新流程的用血医嘱调用输血执行登记(血液医嘱的检查方法=1-用血)
                                blnBloldExec = IsUseBloodAdvice
                                blnExec = True
                            End If
                        End If
                    End If
                    vsAppend.Enabled = False '控件位置变化,行位置变化,避免鼠标点击连续生效
                    '主要处理血库执行和其他医嘱执行，执行显示状态保持一致(之前是输血执行，切换到的其他医嘱的处理)
                    If blnExec = True And blnBloldExec = False And picBlood.Tag = "可见" Then
                        If Not mobjFrmBlood Is Nothing Then mblnShowExec = mobjFrmBlood.IsShowExec
                    End If
                    Call SetExecShow(blnExec, mblnShowExec, blnBloldExec)
                    vsAppend.Enabled = True
                    Me.Refresh
                    
                    '读取执行列表
                    If mblnShowExec And blnExec And blnBloldExec = False Then
                        Call LoadExecList(Val(.TextMatrix(NewRow, COLSend("医嘱ID"))), Val(.TextMatrix(NewRow, COLSend("发送号"))))
                    Else
                        vsExec.Rows = vsExec.FixedRows
                        vsExec.Rows = vsExec.FixedRows + 1
                        vsExec.Row = vsExec.FixedRows
                    End If
                    '输血执行列表读取
                    If blnExec = True And blnBloldExec = True Then
                        If Not mobjFrmBlood Is Nothing Then
                            Call mobjFrmBlood.zlRefresh(Me, glngSys, p住院医嘱发送, Val(.TextMatrix(NewRow, COLSend("医嘱ID"))), mlng医护科室ID, GetInsidePrivs(p住院医嘱发送), 2, mlng病区ID, mblnMoved, mlngFontSize)
                        End If
                    End If
                End If
                    
                '显示可打印的诊疗单据
                Call LoadBillListOld
            End If
        End If
    End With
End Sub

Private Sub vsAppend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    With vsAppend
        If Button = 2 And tbcAppend.Selected.Tag = "发送" Then
            If mcbsMain Is Nothing Then Exit Sub
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    Set objPopup = mcbsMain.FindControl(, conMenu_Report_ClinicBill, False, True) '可能故意隐藏了
                    If Not objPopup Is Nothing Then
                        '如果没有数据，或者数据大于1，则从新更新单个发送记录的单据，因为已有多个单据的话是选择医嘱才会出现的
                        If objPopup.CommandBar.Controls.Count = 0 Or objPopup.CommandBar.Controls.Count > 1 Then Call LoadBillListOld
                        If objPopup.CommandBar.Controls.Count > 0 Then
                            'ShowPopup不会触发InitCommandsPopup事件
                            mintBillPrint = 1
                            objPopup.CommandBar.ShowPopup
                        End If
                    End If
                End If
            End If
        ElseIf Button = 2 And tbcAppend.Selected.Tag = "签名" Then
            If mcbsMain Is Nothing Then Exit Sub
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    Set objPopup = mcbsMain.FindControl(, conMenu_Tool_Sign, False, True) '可能故意隐藏了
                    If Not objPopup Is Nothing Then
                        If objPopup.CommandBar.Controls.Count > 0 Then
                            'ShowPopup不会触发InitCommandsPopup事件
                            objPopup.CommandBar.ShowPopup
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAppend_GotFocus()
    vsAppend.BackColorSel = &HFFCC99
    
    '因为绑定相同,获取焦点时会丢失绑定,Resize会恢复
    picAppend.Tag = "不执行"
    tbcAppend.Height = tbcAppend.Height + 30
    picAppend.Tag = ""
    tbcAppend.Height = tbcAppend.Height - 30
End Sub

Private Sub vsAppend_LostFocus()
    vsAppend.BackColorSel = &HFFEBD7
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HFFEBD7
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'说明:PASS 中的 “RowIn一并给药” 与此方法相同,修改此方法也需要同步修改 PASS同名方法
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub vsColumn_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsColumn.RowData(Row)
        If Val(vsColumn.TextMatrix(Row, 0)) <> 0 Then
            vsAdvice.ColWidth(lngCol) = vsAdvice.ColData(lngCol)
            vsAdvice.ColHidden(lngCol) = False
        Else
            vsAdvice.ColWidth(lngCol) = 0
            vsAdvice.ColHidden(lngCol) = True
        End If
    End If
End Sub

Private Sub vsColumn_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsColumn
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub

Private Sub vsColumn_LostFocus()
    vsColumn.Visible = False
End Sub

Private Sub vsColumn_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsColumn.Cell(flexcpForeColor, Row, 1) = vsColumn.BackColorFixed Then Cancel = True
End Sub

Private Sub FuncAdviceReCharge(Index As Integer)
'功能：费用销帐申请和审核
'参数：Index=冲销子功能索引(0,1)
    Dim blnOK As Boolean, lng医嘱ID As Long
    Dim strCommon As String, intAtom As Integer
    
    '调用费用部件功能
    On Error Resume Next
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Sub
    If gobjInExse Is Nothing Then
        Set gobjInExse = CreateObject("zl9InExse.clsInExse")
        If gobjInExse Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
    End With
        
    '部件调用合法性设置
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", intAtom)
    
    If Index = conMenu_Edit_ChargeDelApply Then
        '医生站只有销帐申请功能，且只能申请药品
        blnOK = gobjInExse.CallReCharge(mfrmParent, gcnOracle, gstrDBUser, glngSys, 0, IIF(mint场合 = 1, 0, 2), mlng病区ID, GetInsidePrivs(1133), mlng病人ID, , lng医嘱ID)
    ElseIf Index = conMenu_Edit_ChargeDelAudit Then
        blnOK = gobjInExse.CallReCharge(mfrmParent, gcnOracle, gstrDBUser, glngSys, 1, 0, mlng病区ID, GetInsidePrivs(1133), mlng病人ID)
    End If
    
    Call GlobalDeleteAtom(intAtom)
    
    If blnOK Then RaiseEvent RequestRefresh(False)
End Sub

Private Sub FuncApplyModi()
'功能：修改申请单
    Dim strSQL As String, rsTmp As Recordset
    With vsAdvice
        '先判断是否是自定义申请单
        strSQL = "Select 文件ID From 医嘱申请单文件 Where 医嘱ID=[1] And RowNum<2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_相关ID)) = 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_相关ID))))
        If rsTmp.RecordCount > 0 Then
            FuncApplyCustom 1, Val(rsTmp!文件ID)
        Else
                        If Val(.TextMatrix(.Row, COL_医嘱状态)) <> 1 Then
                MsgBox "不允许修改已发送的申请。", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(.TextMatrix(.Row, COL_操作类型)) = 6 And .TextMatrix(.Row, COL_诊疗类别) = "E" Then
                Call FuncApplyLIS(Val(.TextMatrix(.Row, COL_申请序号)))
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                Call FuncApplyPACS(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_申请序号)))
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)) = 1 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_检查方法)) = 1 Then
                    Call FuncApplyBlood(4)
                Else
                    Call FuncApplyBlood(1)
                End If
            ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                Call FuncApplyOperation(1)
            ElseIf Val(.TextMatrix(.Row, COL_操作类型)) = 7 And .TextMatrix(.Row, COL_诊疗类别) = "Z" Then
                Call FuncApplyConsultation(1)
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncApplyView()
'功能：查看申请单
    Dim lng医嘱ID As Long
    Dim lngNo As Long
    Dim strSQL As String, rsTmp As Recordset
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_申请序号))
        
        If lng医嘱ID <> 0 And lngNo <> 0 Then
            strSQL = "Select 文件ID From 医嘱申请单文件 Where 医嘱ID=[1] And RowNum<2"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_相关ID)) = 0, lng医嘱ID, Val(.TextMatrix(.Row, COL_相关ID))))
            If rsTmp.RecordCount > 0 Then
                FuncApplyCustom 2, Val(rsTmp!文件ID)
            Else
                If .TextMatrix(.Row, COL_诊疗类别) = "K" Then
                    Call FuncApplyBlood(2)
                ElseIf .TextMatrix(.Row, COL_诊疗类别) = "F" Then
                    Call FuncApplyOperation(2)
                ElseIf Val(.TextMatrix(.Row, COL_操作类型)) = 7 And .TextMatrix(.Row, COL_诊疗类别) = "Z" Then
                    Call FuncApplyConsultation(2)
                ElseIf .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                    '检查
                    If Val(Mid(gstrInUseApp, 1, 1)) = 1 Then
                        Call ShowApply检查(Me, lngNo)
                    End If
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncApplyPACS(ByVal lng医嘱ID As Long, ByRef lng申请序号 As Long)
'功能：调用检查申请单
'参数：lng医嘱ID=修改申请单时当前行的医嘱ID,lng申请序号 =当前修改行的申请序号

    Dim bln中医 As Boolean
    Dim str类型 As String
    Dim blnSucceed As Boolean
    Dim strMsg As String
    Dim lngNo As Long
    
    If CheckAdviceAddModi(IIF(lng医嘱ID = 0, 0, 1)) = False Then Exit Sub
    
    If lng医嘱ID & "_" & lng申请序号 = "0_0" Then
        If Not FuncPathAdd() Then Exit Sub
    End If
    
    '诊断检查
    If InStr(mstr检查入院诊断, "D") > 0 Then
        bln中医 = Sys.DeptHaveProperty(mlng科室ID, "中医科")
        str类型 = IIF(bln中医, "2,12", "2")
        If Not ExistsDiagNoses(mlng病人ID, mlng主页ID, str类型) Then
            strMsg = "病人的入院诊断还没有输入，请先输入病人的入院诊断再下达相关医嘱。"
        End If
        If strMsg <> "" Then
            If InStr(";" & mMainPrivs & ";", ";首页整理;") > 0 Then
                vsAdvice.Refresh
                MsgBox strMsg & vbCrLf & vbCrLf & "请按 [确定] 进入诊断输入界面。", vbInformation, gstrSysName
                blnSucceed = True
                RaiseEvent EditDiagnose(Me, mlng病人ID, mlng主页ID, mlng科室ID, str类型, blnSucceed)
                vsAdvice.Refresh
                If Not blnSucceed Then Exit Sub
            Else
                vsAdvice.Refresh
                MsgBox strMsg, vbInformation, gstrSysName
                vsAdvice.Refresh: Exit Sub
            End If
        End If
    End If
    lngNo = ApplyInPacs(Me, lng申请序号, mlng病人ID, mlng主页ID, Val(mbyt婴儿), mlng病人性质, lng医嘱ID, mlng医护科室ID, mlng科室ID, mlng病区ID, mobjVBA, mobjScript, mrsDefine, mclsMipModule, , mlng前提ID)
    If lngNo <> 0 Then Call LoadAdvice
    
    If mlng路径状态 = 1 And Not gobjPath Is Nothing And lngNo <> 0 Then
        Call FuncPathSet(lngNo)
    End If
End Sub

Private Sub FuncApplyLIS(ByVal lng申请序号 As Long)
'功能：调用检验申请产生申请单和检验医嘱
'参数：lng申请序号=修改申请单时的申请序号
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean, strSQL As String
    Dim strResult As String, strDiag As String, strDept As String, strErr As String
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset '注意此变量不要乱用,在LisInfoTrans方法被赋值
    Dim rsTemp As ADODB.Recordset
    Dim lng医嘱ID As Long, lng相关ID As Long, lng序号 As Long
    Dim lng执行科室ID As Long, lng采集科室ID As Long, lng检验项目ID As Long, lng采集项目ID As Long
    Dim str检验计价性质 As String, str采集计价性质 As String, str检验执行性质 As String, str采集执行性质 As String
    Dim str检验项目 As String, str采集方法 As String, str标本 As String, str紧急 As String
    Dim strCurDate As String, str医嘱内容 As String, str医嘱IDs As String, blnCancel As Boolean
    Dim strDelIDs As String, arrDelID() As String
    Dim Y As Long, j As Long
    Dim str嘱托 As String, str附项 As String
    Dim bln中医 As Boolean, blnSucceed As Boolean
    Dim str类型 As String
    Dim arrAppend As Variant
    Dim lng开单科室ID As Long
    Dim lng附项序号 As Long
    Dim str诊断 As String
    Dim lng假医嘱ID As Long '避免医嘱ID序列值的浪费，在最后提交事务时产生真的医嘱ID
    Dim str医嘱ID As String, str相关ID As String
    Dim varID As Variant
    Dim strTmp As String
    Dim bln提醒对码 As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim strItems As String
    Dim strTabAdvice As String
    Dim blnCheckItem As Boolean '医保管控监测
    Dim rsPrice As ADODB.Recordset
    Dim str摘要 As String, strMsg As String
    Dim dat开始执行时间 As Date
    Dim str开始执行时间 As String
    Dim dat当前时间 As Date
    Dim datTurn As Date
    Dim rsLISInfo As ADODB.Recordset
    Dim lng申请组号 As Long
    
    If lng申请序号 = 0 Then
        If Not FuncPathAdd() Then Exit Sub
    End If
    
    If CheckAdviceAddModi(IIF(lng申请序号 = 0, 0, 1), , datTurn) = False Then Exit Sub
    
    Set rsPati = GetPatiInfo(mlng病人ID, mlng主页ID)
    If rsPati.RecordCount = 0 Then
        MsgBox "未能正确读取病人信息！", vbInformation, gstrSysName
        Exit Sub
    End If
    If lng申请序号 <> 0 Then
        strDiag = GetAdviceDiag(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    End If
    lng开单科室ID = Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2)
    strDept = Sys.RowValue("部门表", lng开单科室ID, "名称")
    Call InitObjLis(p住院医生站)
    If gobjLIS Is Nothing Then Exit Sub
    Call CreatePlugInOK(p住院医嘱下达, mint场合)
    
    On Error GoTo errH
     
    '返回已选择的检验项目格式如下: 采诊科室ID1,执行科室ID1,申请时间1,诊疗项目编码1,标本1,紧急医嘱1,采集方式诊疗项目ID 1;采诊科室ID2,执行科室ID2,申请时间2,诊疗项目编码2,标本2,紧急医嘱2,采集方式诊疗项目ID 2;.....
    strResult = gobjLIS.ShowLisApplicationForm(mfrmParent, lng申请序号, mlng病人ID, mbyt婴儿, mlng主页ID, rsPati!姓名, "" & rsPati!性别, "" & rsPati!年龄, IIF(mlng病人性质 = 1, 1, 2), _
        Val("" & rsPati!门诊号), Val("" & rsPati!住院号), Val("" & rsPati!健康号), strDiag, UserInfo.姓名, UserInfo.部门ID, UserInfo.部门名, lng开单科室ID, strDept, blnCancel, strErr)
     
    If strErr <> "" Then
        MsgBox "检验接口内部错误：" & strErr, vbInformation, gstrSysName
    ElseIf blnCancel Then
        Exit Sub    '取消，退出
    Else
        arrSQL = Array()
        '修改申请单时，先删除旧的医嘱
        If lng申请序号 <> 0 Then
            str医嘱IDs = GetAdivceBy申请序号(lng申请序号)
            For i = 0 To UBound(Split(str医嘱IDs, ","))
                '调用删除前外挂接口
                On Error Resume Next
                If Not gobjPlugIn Is Nothing Then
                    If gobjPlugIn.AdviceDeletBefor(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(Split(str医嘱IDs, ",")(i)), mint场合) = False Then
                        If err.Number = 0 Then Exit Sub
                    End If
                    Call zlPlugInErrH(err, "AdviceDeletBefor")
                End If
                If err.Number <> 0 Then err.Clear
                On Error GoTo errH
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Delete(" & Split(str医嘱IDs, ",")(i) & ",1)"
                strDelIDs = strDelIDs & "," & Split(str医嘱IDs, ",")(i)
            Next
            strDelIDs = Mid(strDelIDs, 2)
        End If
        
        If strResult <> "" Then
            '诊断检查
            If InStr(mstr检查入院诊断, "C") > 0 Then
                bln中医 = Sys.DeptHaveProperty(mlng科室ID, "中医科")
                str类型 = IIF(bln中医, "2,12", "2")
                If Not ExistsDiagNoses(mlng病人ID, mlng主页ID, str类型) Then
                    strMsg = "病人的入院诊断还没有输入，请先输入病人的入院诊断再下达相关医嘱。"
                End If
                If strMsg <> "" Then
                    If InStr(";" & mMainPrivs & ";", ";首页整理;") > 0 Then
                        vsAdvice.Refresh
                        MsgBox strMsg & vbCrLf & vbCrLf & "请按 [确定] 进入诊断输入界面。", vbInformation, gstrSysName
                        blnSucceed = True
                        RaiseEvent EditDiagnose(Me, mlng病人ID, mlng主页ID, mlng科室ID, str类型, blnSucceed)
                        vsAdvice.Refresh
                        If Not blnSucceed Then Exit Sub
                    Else
                        vsAdvice.Refresh
                        MsgBox strMsg, vbInformation, gstrSysName
                        vsAdvice.Refresh: Exit Sub
                    End If
                End If
            End If
            
            bln提醒对码 = True
            
            If mint险类 <> 0 Then
                If gclsInsure.GetCapability(support实时监控, mlng病人ID, mint险类) Then
                    blnCheckItem = True
                End If
            End If

            If strDiag <> "" Then
                str诊断 = GetDiag诊断描述(strDiag)
                If str诊断 <> "" Then
                    str诊断 = "申请单诊断<Split2>0<Split2><Split2>" & str诊断
                End If
            End If
             
            dat当前时间 = zlDatabase.Currentdate()
            strCurDate = "To_Date('" & Format(dat当前时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
            lng序号 = GetMaxAdviceNO(mlng病人ID, mlng主页ID, mbyt婴儿)
            lng申请组号 = -1
            '在该方法中对rsLISInfo, rsTmp赋值
            Call LisInfoTrans(strResult, rsLISInfo, rsTmp)
                        
            '只产生临嘱
            For i = 1 To rsLISInfo.RecordCount
                
                If lng申请组号 <> Val(rsLISInfo!组号 & "") Then
                    lng申请组号 = Val(rsLISInfo!组号 & "")
                    lng申请序号 = Get申请序号
                End If
        
                lng假医嘱ID = lng假医嘱ID + 1
                str相关ID = "<FAKEID>" & lng假医嘱ID & "</FAKEID>"
                lng相关ID = lng假医嘱ID
                lng采集科室ID = Val(rsLISInfo!采集科室ID & "")
                lng执行科室ID = Val(rsLISInfo!执行科室ID & "")
                str开始执行时间 = rsLISInfo!开始执行时间 & ""
                str标本 = rsLISInfo!标本 & ""
                str附项 = rsLISInfo!附项 & ""
                str嘱托 = rsLISInfo!嘱托 & ""
                str紧急 = rsLISInfo!紧急 & ""
                lng采集项目ID = Val(rsLISInfo!采集项目ID & "")
                lng检验项目ID = Val(rsLISInfo!检验项目ID & "")
                                
                dat开始执行时间 = CDate(str开始执行时间)
                str开始执行时间 = "To_Date('" & Format(dat开始执行时间, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
                
                '判断是否是补录医嘱
                If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Or (mbyt婴儿 > 0 And datTurn <> CDate(0)) Then
                    str紧急 = "2"
                ElseIf DateDiff("n", dat开始执行时间, dat当前时间) > gint补录间隔 Then
                    str紧急 = "2"
                End If
                    
                'a.先产生检验医嘱 申请单开出来的的检验医嘱只有一个检验项目ID
                rsTmp.Filter = "ID=" & lng检验项目ID
                str检验项目 = rsTmp!名称 & ""
                str检验计价性质 = Val("" & rsTmp!计价性质)
                str检验执行性质 = IIF("" & rsTmp!执行科室 = "", "NULL", "" & rsTmp!执行科室)
                str医嘱内容 = str检验项目 & IIF("" = rsLISInfo!时间内容 & "", "", "(" & rsLISInfo!时间内容 & ")")
                lng序号 = lng序号 + 1
                str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(lng检验项目ID) & "||2")
                blnCancel = CheckLISAppAdvice(2, mlng病人ID, mlng主页ID, mint险类, "C", lng检验项目ID, lng开单科室ID, UserInfo.姓名, lng执行科室ID, Val(rsTmp!执行科室 & ""), str摘要 & "||0||0|| ||0")
                If Not blnCancel Then Exit Sub
                
                lng假医嘱ID = lng假医嘱ID + 1
                str医嘱ID = "<FAKEID>" & lng假医嘱ID & "</FAKEID>"
                lng医嘱ID = lng假医嘱ID
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
                    str医嘱ID & "," & str相关ID & "," & lng序号 & ",2," & mlng病人ID & "," & _
                    mlng主页ID & "," & mbyt婴儿 & ",1,1,'C'," & _
                    lng检验项目ID & ",Null,Null,Null,1," & _
                    "'" & str医嘱内容 & "',Null," & "'" & str标本 & "','一次性',Null," & _
                    "Null,Null,Null," & str检验计价性质 & "," & lng执行科室ID & _
                    "," & str检验执行性质 & "," & str紧急 & "," & str开始执行时间 & ",Null," & mlng科室ID & "," & _
                    lng开单科室ID & ",'" & UserInfo.姓名 & "'," & strCurDate & ",NULL," & ZVal(mlng前提ID) & "," & _
                    "NULL,0,Null," & IIF(str摘要 = "", "Null", "'" & str摘要 & "'") & ",'" & UserInfo.姓名 & "'" & _
                    ",Null,Null,Null,Null," & lng申请序号 & ",null,null,null,null,null,'" & rsLISInfo!时间ID & "')"
                
                strItems = strItems & "," & lng检验项目ID & ":" & lng执行科室ID
                
                If blnCheckItem Then
                    strTabAdvice = _
                        "select " & lng医嘱ID & " as ID," & lng序号 & " as 序号," & lng相关ID & " as 相关ID,'C' as 诊疗类别," & lng检验项目ID & " as 管码项目ID," & _
                        lng检验项目ID & " as 诊疗项目ID,-null as 收费细目ID, 1 As 总量, 0 As 单量,'" & str标本 & "' as 标本部位,'' As 检查方法," & _
                        "0 as 执行标记," & Val("" & rsTmp!计价性质) & " as 计价特性, 0 As 附加手术," & Val("" & rsTmp!执行科室) & " As 执行性质," & lng执行科室ID & " as 执行科室id from dual"
                End If

                'b.再产生采集方法医嘱
                rsTmp.Filter = "ID=" & lng采集项目ID
                str采集方法 = rsTmp!名称 & ""
                str采集计价性质 = Val("" & rsTmp!计价性质)
                str采集执行性质 = "" & rsTmp!执行科室
                str医嘱内容 = AdviceMakeText(str检验项目, str采集方法, str标本)
                If "" <> rsLISInfo!时间内容 & "" Then str医嘱内容 = str医嘱内容 & "(" & rsLISInfo!时间内容 & ")"
                lng序号 = lng序号 + 1
                str摘要 = ""
                str摘要 = gclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", CStr(lng采集项目ID) & "||2")
                blnCancel = CheckLISAppAdvice(2, mlng病人ID, mlng主页ID, mint险类, "E", lng采集项目ID, lng开单科室ID, UserInfo.姓名, lng采集科室ID, Val(rsTmp!执行科室 & ""), str摘要 & "||0||0|| ||0")
                If Not blnCancel Then Exit Sub
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_病人医嘱记录_Insert(" & _
                    str相关ID & ",Null," & lng序号 & ",2," & mlng病人ID & "," & _
                    mlng主页ID & "," & mbyt婴儿 & ",1,1,'E'," & _
                    lng采集项目ID & ",Null,Null,Null,1," & _
                    "'" & str医嘱内容 & "','" & str嘱托 & "'," & "'" & str标本 & "','一次性',Null," & _
                    "Null,Null,Null," & str采集计价性质 & "," & lng采集科室ID & _
                    "," & str采集执行性质 & "," & str紧急 & "," & str开始执行时间 & ",Null," & mlng科室ID & "," & _
                    lng开单科室ID & ",'" & UserInfo.姓名 & "'," & strCurDate & ",NULL," & ZVal(mlng前提ID) & "," & _
                    "NULL,0,Null," & IIF(str摘要 = "", "Null", "'" & str摘要 & "'") & ",'" & UserInfo.姓名 & "'" & _
                    ",Null,Null,Null,Null," & lng申请序号 & ",null,null,null,null,null,'" & rsLISInfo!时间ID & "')"
                
                strItems = strItems & "," & lng采集项目ID & ":" & lng采集科室ID
                
                If blnCheckItem Then
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & lng相关ID & " as ID," & lng序号 & " as 序号,-null as 相关ID,'E' as 诊疗类别," & lng检验项目ID & " as 管码项目ID," & _
                        lng采集项目ID & " as 诊疗项目ID,-null as 收费细目ID, 1 As 总量, 0 As 单量,'" & str标本 & "' as 标本部位,'' As 检查方法," & _
                        "0 as 执行标记," & Val("" & rsTmp!计价性质) & " as 计价特性, 0 As 附加手术," & Val("" & rsTmp!执行科室) & " As 执行性质," & lng采集科室ID & " as 执行科室id from dual"
                End If
                
                '医保对码检查
                If gint医保对码 = 2 Then bln提醒对码 = True
                strMsg = CheckAdviceInsure(mint险类, bln提醒对码, mlng病人ID, mlng病人性质, "", Mid(strItems, 2), Left(str医嘱内容, 50), mlng病区ID)
                If strMsg <> "" Then
                    If gint医保对码 = 1 Then
                        vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "要继续保存医嘱吗？", Me)
                        If vMsg = vbNo Or vMsg = vbCancel Then Exit Sub
                        If vMsg = vbIgnore Then bln提醒对码 = False
                    ElseIf gint医保对码 = 2 Then
                        MsgBox strMsg & vbCrLf & vbCrLf & "请先和相关人员联系处理，否则医嘱将不允许保存。", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strMsg = ""
                End If
                
                '医保管控实时监测：首次输入(经过)或者更改时检查
                If blnCheckItem Then
                    If MakePriceRecord申请单("12", mlng病人ID, mlng主页ID, strTabAdvice, strItems, rsPati!费别 & "", lng开单科室ID, rsPrice) Then
                        If Not gclsInsure.CheckItem(mint险类, 1, 0, rsPrice) Then
                            MsgBox "医保监测检查未通(执行Insure.CheckItem接口)，本次下达的LIS申请单不能保存。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If str附项 <> "" And str诊断 <> "" Then
                    str附项 = str诊断 & "<Split1>" & str附项
                ElseIf str附项 = "" And str诊断 <> "" Then
                    str附项 = str诊断
                End If
                
                '单据申请附项，有外键，所以先产生医嘱
                If str附项 <> "" Then
                    arrAppend = Split(str附项, "<Split1>")
                    For j = 0 To UBound(arrAppend)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_病人医嘱附件_Insert(" & str相关ID & "," & _
                            "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                            j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                            IIF(j = 0, ",1", "") & ")"
                        lng附项序号 = j + 1
                    Next
                End If
                If strDiag <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_病人诊断医嘱_Insert(" & str相关ID & ",'" & strDiag & "')"
                End If
                rsLISInfo.MoveNext
            Next
        End If
        
        '用序列产生真实的医嘱ID
        If lng假医嘱ID > 0 Then
            For j = 1 To lng假医嘱ID
                Y = zlDatabase.GetNextID("病人医嘱记录")
                If j = 1 Then
                    str医嘱IDs = ""
                    str医嘱IDs = Y
                Else
                    str医嘱IDs = str医嘱IDs & "," & Y
                End If
            Next
            varID = Split(str医嘱IDs, ",")
            
            For i = 0 To UBound(arrSQL)
                strTmp = arrSQL(i)
                
                If InStr(strTmp, "<FAKEID>") > 0 Then
                    j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                    strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
                    
                    If InStr(strTmp, "<FAKEID>") > 0 Then '最多替换两次
                        j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                        strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
                    End If
                    arrSQL(i) = strTmp
                End If
            Next
        End If
        
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        '临床路径判断
         If mlng路径状态 = 1 And Not gobjPath Is Nothing And lng申请序号 <> 0 Then
             Call FuncPathSet(lng申请序号)
         End If
         Call ZLHIS_CIS_001(mclsMipModule, mlng病人ID, mstr姓名, mstr住院号, , IIF(mlng病人性质 = 1, 1, 2), mlng主页ID, mlng病区ID, , mlng科室ID, "", , mstr床号, _
               lng相关ID, str紧急, 1, "E", "", UserInfo.姓名, Format(dat开始执行时间, "yyyy-MM-dd HH:MM:00"), lng开单科室ID, "", , , "")
    
         '刷新医嘱
         Call RefreshData
       
        '调用删除后外挂接口
        On Error Resume Next
        arrDelID = Split(strDelIDs, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                If Not gobjPlugIn Is Nothing Then
                    Call gobjPlugIn.AdviceDeleted(glngSys, p住院医嘱下达, mlng病人ID, mlng主页ID, Val(arrDelID(i)), mint场合)
                    Call zlPlugInErrH(err, "AdviceDeleted")
                End If
            End If
        Next
        If err.Number <> 0 Then err.Clear
        On Error GoTo 0
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetAdivceBy申请序号(ByVal lng申请序号 As Long) As String
'功能：根据申请序号获取所有检查医嘱ID串（采集医嘱ID）
    Dim i As Long, strTmp As String
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_申请序号)) = lng申请序号 Then
                If Val(.TextMatrix(i, COL_操作类型)) = 6 And .TextMatrix(i, COL_诊疗类别) = "E" Then
                    strTmp = strTmp & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
        GetAdivceBy申请序号 = Mid(strTmp, 2)
    End With
End Function

Private Function AdviceMakeText(ByVal str检验 As String, ByVal str采集 As String, ByVal str标本 As String) As String
'功能：产生检验医嘱的医嘱内容
    Dim i As Long, strText As String, strField As String, blnDefine As Boolean
    
    If mobjVBA Is Nothing Then
        On Error Resume Next
        Set mobjVBA = CreateObject("ScriptControl")
        err.Clear: On Error GoTo 0
        
        If Not mobjVBA Is Nothing Then
            mobjVBA.Language = "VBScript"
            Set mobjScript = New clsScript
            mobjVBA.AddObject "clsScript", mobjScript, True
        End If
    End If
               
    '确定是否定义
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "诊疗类别='C'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(NVL(mrsDefine!医嘱内容)) = "" Then
            blnDefine = False
        End If
    End If
    
    If Not blnDefine Then
        strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
    Else
        strText = mrsDefine!医嘱内容
        If InStr(strText, "[检验项目]") > 0 Then
            strField = str检验
            strText = Replace(strText, "[检验项目]", """" & strField & """")
        End If
        If InStr(strText, "[检验标本]") > 0 Then
            strField = str标本
            strText = Replace(strText, "[检验标本]", """" & strField & """")
        End If
        If InStr(strText, "[采集方法]") > 0 Then
            strField = str采集
            strText = Replace(strText, "[采集方法]", """" & strField & """")
        End If
        
        '计算医嘱内容
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = str检验 & IIF(str标本 <> "", "(" & str标本 & ")", "")
        End If
        err.Clear: On Error GoTo 0
    End If
        
    AdviceMakeText = strText
End Function

Private Sub GetAdvicesSameSend(ByVal lng发送号 As Long, ByRef strLIS As String, ByRef strALL As String, Optional ByVal str诊疗类别 As String = "C")
'功能：根据发送号获取一起发送的医嘱的主ID
'参数：strLIS 出参，检验医嘱ID串，strALL所有医嘱ID串
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIDs As String, i As Long
    Dim str检验IDs As String
    
    strSQL = "Select b.诊疗类别, Nvl(b.相关id,b.Id) As id" & vbNewLine & _
        "From 病人医嘱发送 A, 病人医嘱记录 B" & vbNewLine & _
        "Where a.医嘱id = b.Id And a.发送号 =[1] And b.病人id =[2] And b.主页id =[3]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng发送号, mlng病人ID, mlng主页ID)
    For i = 1 To rsTmp.RecordCount
        If InStr("," & strIDs & ",", "," & rsTmp!ID & ",") = 0 Then
            strIDs = strIDs & "," & rsTmp!ID
            If rsTmp!诊疗类别 & "" = str诊疗类别 Then
                str检验IDs = str检验IDs & "," & rsTmp!ID
            End If
        End If
        rsTmp.MoveNext
    Next
    
    strALL = Mid(strIDs, 2)
    strLIS = Mid(str检验IDs, 2)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PrintLisReport(ByVal lngPatiDeptID As Long, objFrm As Object)
    'LIS病区检验报告打印
    Dim objLisWork As Object
    Set objLisWork = CreateObject("zl9LisWork.clsLISImg")
    On Error GoTo hErr
    If Not objLisWork Is Nothing Then
        Call objLisWork.ShowPatientRptPrint(gcnOracle, glngSys, lngPatiDeptID, mMainPrivs, objFrm)
    End If
    Exit Sub
hErr:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub PrintBloodReport(ByVal lngAdviceID As Long, objFrm As Object)
    '输血执行单打印
    If InitObjBlood(True) = True Then
        Call gobjPublicBlood.ShowBloodInstantRptPrint(objFrm, lngAdviceID)
    End If
End Sub

Private Function CheckPatiIsAduit() As Boolean
'功能：检查病人是否开始审核
    Dim rsTmp As Recordset, strSQL As String
    Dim int审核标志 As Integer
    
    
    
    If mblnBatch Then CheckPatiIsAduit = True: Exit Function
    strSQL = "Select a.审核标志 From 病案主页 a" & _
                " Where a.病人ID=[1] And a.主页ID=[2]"
    On err GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "病人审核检查", mlng病人ID, mlng主页ID)
    If rsTmp.RecordCount > 0 Then
        If Val("" & rsTmp!审核标志) >= 1 And gbyt病人审核方式 = 1 Then
            MsgBox "该病人的费用正在审核或已经审核，不允许操作医嘱和费用。", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPatiIsAduit = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置医嘱清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    mlngFontSize = IIF(bytSize = 0, 9, 12)
    
    '对于vsFlexGrid控件在使用个性化设置时会加大列宽，因此在窗体初次加载是不设置字体,主要是getForm方法引起
    If Not Me.Visible Then
        vsAdvice.FontSize = mlngFontSize
        vsAppend.FontSize = mlngFontSize
        vsExec.FontSize = mlngFontSize
        vsfAdivceDetail.FontSize = mlngFontSize
        If Not mfrmCompoundMedicine Is Nothing Then
            mfrmCompoundMedicine.vsSend.FontSize = mlngFontSize
            mfrmCompoundMedicine.vsExec.FontSize = mlngFontSize
        End If
    End If
    
    If mvarCond.显示模式 = 0 Then
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_内容)
    Else
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_医嘱内容)
    End If
    
    Call Grid.SetFontSize(vsAppend, mlngFontSize)
    If tbcAppend.Selected.Tag = "发送" Then '第一行的字体要特殊处理
        vsAppend.Cell(flexcpFontSize, 0, 0, 0, vsAppend.Cols - 1) = mlngFontSize
    End If
    
    Call Grid.SetFontSize(vsExec, mlngFontSize)
    Call Grid.SetFontSize(vsfAdivceDetail, mlngFontSize)
    
    If Not mfrmCompoundMedicine Is Nothing Then
        Call Grid.SetFontSize(mfrmCompoundMedicine.vsSend, mlngFontSize)
        Call Grid.SetFontSize(mfrmCompoundMedicine.vsExec, mlngFontSize)
    End If
    
    '血液执行和血液明细窗体
    If Not mobjFrmBloodList Is Nothing Then
        If mobjFrmBloodList.Visible = True Then Call mobjFrmBloodList.SetFontSize(mlngFontSize)
    End If
    
    If Not mobjFrmBlood Is Nothing Then
        If mobjFrmBlood.Visible = True Then Call mobjFrmBlood.SetFontSize(mlngFontSize)
    End If
    
    Call SetRTFFont(0)
End Sub

Private Function CheckBabyEdit(ByVal lngBaby As Long) As Integer
'功能：检查母婴分离是否允许编辑
'返回：0，允许编辑，1=婴儿科室不允许编辑病人医嘱，2=病人科室不允许编辑婴儿医嘱
'参数：lngBaby婴儿序号
    CheckBabyEdit = 0
    If mlng婴儿科室ID <> 0 And mstr婴儿 <> "" Then
        If (mlng婴儿科室ID = mlng医护科室ID Or mlng婴儿病区ID = mlng医护科室ID) And lngBaby = 0 Then
            CheckBabyEdit = 1
        ElseIf (mlng科室ID = mlng医护科室ID Or mlng病区ID = mlng医护科室ID) And lngBaby > 0 Then
            CheckBabyEdit = 2
        End If
    End If
End Function

Private Function CheckDelAdivceOfPathItem(ByVal lng医嘱ID As Long) As Boolean
'功能：检查医嘱对应的路径项目是否允许删除，如果是必须执行的项目所对应的医嘱，则需要弹出原因选择并更新变异原因，
'       添加过变异原因的不再添加
'返回：True-可以删除该医嘱，false-不可删除
'参数:lng医嘱ID
    Dim blnCancel As Boolean, blnMust As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, rsAdvice As ADODB.Recordset
    Dim strReason As String
    Dim vPoint As PointAPI
    Dim strTemp As String
    Dim arrTmp As Variant
    Dim arrSQL As Variant
    Dim i As Long

    '1.检查路径项目
    strSQL = "Select  c.Id as 执行Id, c.分类,c.变异原因,d.执行方式,c.天数,c.阶段ID,c.路径记录ID,c.项目ID " & _
             " From 病人路径医嘱 B, 病人路径执行 C, 临床路径项目 D" & vbNewLine & _
             "Where b.病人医嘱Id=[1] And b.路径执行id = c.Id And d.Id = c.项目id And d.执行方式 in (1,2,4)"

    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "检查路径医嘱", lng医嘱ID)

    If rsTmp.RecordCount < 1 Then
        CheckDelAdivceOfPathItem = True
        Exit Function    '非 必须生成的路径医嘱
    End If
    '2.检查医嘱能否删除
    '该路径项目存在已校对但未作废的其他医嘱，提示并禁止删除    医嘱状态 ：3-已校对
    strSQL = "Select a.病人医嘱ID,b.医嘱状态 " & vbNewLine & _
             "From 病人路径医嘱 A, 病人医嘱记录 B" & vbNewLine & _
             "Where a.路径执行id = [1] And a.病人医嘱id = b.Id  And b.医嘱状态>1 and b.医嘱状态<>4"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "检查路径医嘱", rsTmp!执行Id)

    If rsAdvice.RecordCount > 0 Then
        MsgBox "删除医嘱所在的路径项目中存在已校对但未作废的医嘱，请先作废该医嘱后再执行此操作。", vbInformation, gstrSysName
        CheckDelAdivceOfPathItem = False
        Exit Function
    End If
    

    If mint场合 = 1 Then
        '对于已经过审核的医嘱，不允许修改删除。
        strSQL = "Select b.病人医嘱ID From 病人路径医嘱 B, 病人医嘱记录 C Where b.路径执行id = [1] And b.病人医嘱id = c.Id And c.开嘱医生 Like '%/%'"
        Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "检查路径医嘱", rsTmp!执行Id)
        If rsAdvice.RecordCount > 0 Then
            MsgBox "删除的医嘱中存在医生审查的医嘱，请取消审查后再执行此操作。", vbInformation, gstrSysName
            CheckDelAdivceOfPathItem = False
            Exit Function
        End If
    End If
    
    '根据执行方式 决定是否有必要添加变异原因
    blnMust = CheckPathItemIsMust(Val(rsTmp!执行方式 & ""), Val("" & rsTmp!天数), Val("" & rsTmp!路径记录id), Val("" & rsTmp!阶段id), Val("" & rsTmp!项目ID))
    If Not blnMust Then CheckDelAdivceOfPathItem = True: Exit Function
    
    '----------------------------
    '3.必须生成的项目填写变异原因
    For i = 1 To rsTmp.RecordCount
        If rsTmp!变异原因 & "" = "" Then
            strTemp = strTemp & rsTmp!执行Id & "," & rsTmp!分类 & ";"
        End If
        rsTmp.MoveNext
    Next
    
    If strTemp = "" Then
        CheckDelAdivceOfPathItem = True
        Exit Function
    Else
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    End If

    strSQL = "Select b.名称 as 分类,a.编码 as ID,a.编码,a.名称,a.简码 From 变异常见原因 a,变异常见原因 b" & _
             " Where a.性质=1 And a.末级=1 And a.上级=b.编码 And b.末级=0 " & _
             " Order by 分类,a.编码"
    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)

    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "变异常见原因", True, , , True, True, True, _
                                      vPoint.X, vPoint.Y, vsAdvice.RowHeight(vsAdvice.Row), blnCancel, False, True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "系统没有初始变异常见原因，请与系统管理员联系。", vbInformation, gstrSysName
        End If
        Exit Function
    Else
        strReason = rsTmp!ID
    End If

    If strReason <> "" Then
        arrSQL = Array()
        For i = 0 To UBound(Split(strTemp, ";"))
            arrTmp = Split(Split(strTemp, ";")(i), ",")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人路径生成_Update(" & arrTmp(0) & ",'" & arrTmp(1) & "',Null ,Null,Null,Null,Null,'" & strReason & "')"
        Next
        '不添加事务处理，若变异原因添加失败，医嘱不会删除，再次删除时，会重新添加变异原因后才可删除。
        For i = LBound(arrSQL) To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        CheckDelAdivceOfPathItem = True
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetSendCommandBar()
'功能：设置医嘱发送“按钮”样式--菜单栏/工具栏
'说明：
'      框架窗体的菜单(mcbsMain)必须修改，工具栏是否修改根据 frmDockInAdvice 是什么模式,是什么模式由 mblnInsideTools 区分；
'      mblnInsideTools  =True 修改cbsSub的工具栏，=False 修改mcbsMain的工具栏；
    Dim objControl As CommandBarControl
    Dim objCtlTmp As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim strPrivs As String
    Dim strPara As String
    Dim objCbs As Object
    Dim i As Long
    
    If mcbsMain Is Nothing Then Exit Sub
    
    On Error GoTo errH
    If gstr输液配置中心 <> "" Then
        strPrivs = GetInsidePrivs(p住院医嘱发送)
        If InStr(";" & strPrivs & ";", ";发送药疗临嘱;") = 0 Or InStr(";" & strPrivs & ";", ";发送药疗长嘱;") = 0 Then
            strPrivs = ""
        End If
    End If
    
    '菜单添加
    Set objMenuBar = mcbsMain.ActiveMenuBar.Controls(IIF(mblnInsideTools, 2, 3))
    For i = objMenuBar.CommandBar.Controls.Count To 1 Step -1
        If objMenuBar.CommandBar.Controls(i).ID = conMenu_Edit_Send Then
            objMenuBar.CommandBar.Controls(i).Delete: Exit For
        End If
    Next i
    strPara = zlDatabase.GetPara("来源病区", glngSys, p输液配置中心, "")
    With objMenuBar.CommandBar.Controls
        Set objControl = .Find(, conMenu_Edit_Audit)
        If Not objControl Is Nothing Then
            If strPrivs <> "" Then
                Set objMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "医嘱发送(&G)", objControl.Index + 1)
                Set objCtlTmp = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "发送所有医嘱(&G)")
                If InStr("," & strPara & ",", "," & mlng病区ID & ",") > 0 Or strPara = "" Then
                    objCtlTmp.Caption = "发送医嘱(不含输液)(&G)"
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送医嘱(仅输液)(&I)")
                Else
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送静脉营养药品(&I)")
                End If
                objControl.IconId = conMenu_Edit_Send
            Else
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "医嘱发送(&G)", objControl.Index + 1): objControl.ToolTipText = ""
            End If
        End If
    End With
    
    '工具栏添加
    Set objCbs = IIF(mblnInsideTools, cbsSub, mcbsMain)
 
    For i = objCbs(2).Controls.Count To 1 Step -1
        If objCbs(2).Controls(i).ID = conMenu_Edit_Send Then
            objCbs(2).Controls(i).Delete: Exit For
        End If
    Next i
     
    With objCbs(2).Controls
        Set objControl = .Find(, conMenu_Edit_Audit)
        If Not objControl Is Nothing Then
            If strPrivs <> "" Then
                Set objMenuBar = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "发送", objControl.Index + 1): objMenuBar.Style = xtpButtonIconAndCaption
                Set objCtlTmp = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "发送所有医嘱")
                
                If InStr("," & strPara & ",", "," & mlng病区ID & ",") > 0 Or strPara = "" Then
                    objCtlTmp.Caption = "发送医嘱(不含输液)"
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送医嘱(仅输液)")
                Else
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "发送静脉营养药品")
                End If
                
                objControl.IconId = conMenu_Edit_Send
            Else
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送", objControl.Index + 1): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "医嘱发送"
            End If
        End If
    End With
    
    '热键
    With objCbs.KeyBindings
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send   '发送所有医嘱
        .Add 0, vbKeyF2, conMenu_Edit_SendInfusion '发送输液药品医嘱
    End With
    
    If mblnInsideTools Then objCbs.RecalcLayout
    
    mcbsMain.RecalcLayout
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AddToolBarInDoctor()
'功能：设置工具栏按钮，对应于医嘱菜单下面的工具栏的按钮，先将其删掉再添加
    Dim objControl As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim objPopup As CommandBarPopup
    Dim varArr As Variant
    Dim strTmp As String
    Dim lngTmp As Long
    Dim objCbs As Object
    Dim lngIdx As Long
    Dim i As Long
    
    Dim intTmp As Integer
    Dim strName As String
    Dim lngID As Long
    
    
    If mcbsMain Is Nothing Then Exit Sub
    
    strTmp = "," & conMenu_Edit_NewItem & "," & conMenu_Edit_PacsApply & "," & conMenu_Edit_ApplyModi & "," & conMenu_Edit_ApplyView & "," & conMenu_Edit_ApplyDel & "," & _
        conMenu_Edit_Blankoff & "," & conMenu_Edit_Stop & "," & conMenu_Edit_TraReaction & "," & conMenu_Edit_SendBilling & "," & conMenu_Edit_Send & "," & conMenu_Edit_Untread & "," & _
        conMenu_Edit_Compend & "," & (conMenu_Edit_Compend * 10# + 2) & "," & conMenu_Edit_Compend * 10# + 3 & "," & conMenu_Edit_MarkMap & "," & conMenu_Edit_MarkKeyMap & "," & conMenu_Manage_ReportLisView & "," & conMenu_Manage_ReportPrint & "," & _
        conMenu_Edit_BatExecute & "," & conMenu_Edit_ChargeDelApply & "," & conMenu_Edit_MediAudit & "," & conMenu_Tool_SignNew & "," & conMenu_Edit_Audit & "," & conMenu_Edit_Price & "," & _
        conMenu_Edit_ReStop & "," & conMenu_Report_Reports & "," & conMenu_Manage_ThingAudit & "," & conMenu_Edit_ChargeOff & "," & conMenu_Edit_AdvicePrice & ","
    strTmp = strTmp & "," & conMenu_Edit_PacsApply & "," & (conMenu_Edit_PacsApply * 10# + 1) & "," & conMenu_Edit_LISApply & "," & (conMenu_Edit_LISApply * 10# + 1) & "," & conMenu_Edit_BloodApply & "," & (conMenu_Edit_BloodApply * 10# + 1)
    strTmp = strTmp & "," & conMenu_Edit_OperationApply & "," & (conMenu_Edit_OperationApply * 10# + 1) & "," & conMenu_Edit_ConsultationApply & "," & (conMenu_Edit_ConsultationApply * 10 + 1) & ","
    '工具栏添加
    If mblnInsideTools Then
        Set objCbs = cbsSub
        cbsSub(2).Visible = Not mblnHideFilter
    Else
        Set objCbs = mcbsMain
    End If

    '找到要添加的位置
    lngIdx = 0
    For Each objControl In objCbs(2).Controls '先求出前面的最后一个Control
        If Val(Left(objControl.ID, 1)) <> conMenu_FilePopup And Val(Left(objControl.ID, 1)) <> conMenu_ManagePopup Then
            If objControl.Index > 1 Then
                Set objControl = objCbs(2).Controls(objControl.Index - 1)
                lngIdx = objControl.Index
            Else
                lngIdx = 1
            End If
            Exit For
        End If
    Next
    
    '删除工具栏按钮
    For i = objCbs(2).Controls.Count To 1 Step -1
        If InStr(strTmp, "," & objCbs(2).Controls(i).ID & ",") > 0 Then
            objCbs(2).Controls(i).Delete
        Else
            If mblnInsideTools Then objCbs(2).Controls(i).Delete
        End If
    Next i

    With objCbs(2).Controls
        If mvarCond.过滤模式 <> 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "新开", lngIdx + 1): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "新开")
                    objControl.IconId = conMenu_Edit_NewItem
                .Add xtpControlButton, conMenu_Edit_Modify, "修改"
                .Add xtpControlButton, conMenu_Edit_Delete, "删除"
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
        End If
        
        If mint场合 = 0 Then '只有住院医生工作站调用时才有这几个按钮
            strTmp = ""
            intTmp = Val(Mid(gstrInUseApp, 1, 1))
            If intTmp = 1 Then strTmp = strTmp & ",检查申请:" & conMenu_Edit_PacsApply
            intTmp = Val(Mid(gstrInUseApp, 2, 1))
            If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",检验申请:" & conMenu_Edit_LISApply
            intTmp = Val(Mid(gstrInUseApp, 3, 1))
            If intTmp = 1 Then strTmp = strTmp & ",输血申请:" & conMenu_Edit_BloodApply
            intTmp = Val(Mid(gstrInUseApp, 4, 1))
            If intTmp = 1 Then strTmp = strTmp & ",手术申请:" & conMenu_Edit_OperationApply
            intTmp = Val(Mid(gstrInUseApp, 5, 1))
            If intTmp = 1 Then strTmp = strTmp & ",会诊申请:" & conMenu_Edit_ConsultationApply
            Get自定义申请单 2, mstr自定义申请单IDs
            If mstr自定义申请单IDs <> "" Then
                For i = 0 To UBound(Split(mstr自定义申请单IDs, "|"))
                    strTmp = strTmp & "," & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr自定义申请单IDs, "|")(i), ",")(0)
                Next
            End If
            strTmp = Mid(strTmp, 2)
            
            If strTmp <> "" Then
                If InStr(strTmp, ",") = 0 Then
                    strName = Split(strTmp, ":")(0)
                    lngID = Val(Split(strTmp, ":")(1))
                    Set objControl = .Add(xtpControlButton, lngID, strName, lngIdx + 1)
                        objControl.IconId = conMenu_Manage_Request
                        objControl.ToolTipText = strName
                        objControl.Style = xtpButtonIconAndCaption
                        objControl.BeginGroup = True
                                                If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    lngIdx = objControl.Index
                Else
                    varArr = Split(strTmp, ",")
                    For i = 0 To UBound(varArr)
                        strTmp = varArr(i)
                        strName = Split(strTmp, ":")(0)
                        lngID = Val(Split(strTmp, ":")(1))
                        
                        If i = 0 Then
                            Set objPopup = .Add(xtpControlSplitButtonPopup, lngID, strName, lngIdx + 1)
                                objPopup.IconId = conMenu_Manage_Request
                                objPopup.BeginGroup = True
                                objPopup.ToolTipText = strName
                                objPopup.Style = xtpButtonIconAndCaption
                                With objPopup.CommandBar.Controls
                                    Set objControl = .Add(xtpControlButton, lngID * 10# + 1, strName)
                                End With
                        Else
                            Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                        End If
                        If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                    Next
                    lngIdx = objPopup.Index
                End If
            End If
        End If
        
        If mvarCond.过滤模式 = 3 And mint场合 = 0 Then '只有住院医生工作站调用时才有这几个按钮
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "修改", lngIdx + 1)
                objControl.IconId = 3002
                objControl.ToolTipText = "修改申请"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "查看", objControl.Index + 1)
                objControl.IconId = 102
                objControl.ToolTipText = "查看申请"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "取消", objControl.Index + 1)
                objControl.IconId = 3004
                objControl.ToolTipText = "取消申请"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "作废", lngIdx + 1)
            objControl.BeginGroup = True
            objControl.Style = xtpButtonIconAndCaption
        lngIdx = objControl.Index
        
        If mint场合 = 0 Then
            If mvarCond.过滤模式 <> 3 Then  '只有住院医生工作站调用时才有这几个按钮
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "停止", objControl.Index + 1)
                    objControl.Style = xtpButtonIconAndCaption
                If gbln血库系统 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReaction, "输血反应", objControl.Index + 1)
                        objControl.IconId = 4113
                        objControl.Style = xtpButtonIconAndCaption
                End If
            End If
            
            If InStr(GetInsidePrivs(p住院医嘱下达), "发送门诊费用") = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "发送(&G)", objControl.Index + 1)
                    objControl.IconId = conMenu_Edit_Send
                    objControl.BeginGroup = True
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
            Else
                Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "发送", objControl.Index + 1)
                With objPopup.CommandBar.Controls
                    .Add xtpControlButton, conMenu_Edit_SendBilling, "住院记帐"
                    .Add xtpControlButton, conMenu_Edit_SendCharge, "门诊收费"
                End With
                objPopup.Style = xtpButtonIconAndCaption
                lngIdx = objPopup.Index
            End If
        End If
        
        If mint场合 = 1 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "发送", lngIdx + 1)
            objControl.Style = xtpButtonIconAndCaption
            If Val(zlDatabase.GetPara("发送前自动校对", glngSys, p住院医嘱发送, 0)) = 1 Then
                objControl.BeginGroup = True
            End If
            lngIdx = objControl.Index
        End If
        
        If mint场合 = 2 Then
            If InStr(GetInsidePrivs(p住院医嘱下达), "发送门诊费用") = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "临嘱发送(&G)", objControl.Index + 1): objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_Send
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            Else
                Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "临嘱发送", objControl.Index + 1): objPopup.BeginGroup = True
                With objPopup.CommandBar.Controls
                    .Add xtpControlButton, conMenu_Edit_SendBilling, "住院记帐"
                    .Add xtpControlButton, conMenu_Edit_SendCharge, "门诊收费"
                End With
                objPopup.Style = xtpButtonIconAndCaption
                lngIdx = objPopup.Index
            End If
        End If
        If mint场合 = 1 And mvarCond.过滤模式 <> 3 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "校对", lngIdx + 1)
            objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        End If
        
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Untread, "回退", lngIdx + 1)
            objPopup.Style = xtpButtonIconAndCaption
        lngIdx = objPopup.Index
        
        If mvarCond.过滤模式 = 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "查阅", lngIdx + 1): objPopup.BeginGroup = True
                objPopup.IconId = conMenu_Manage_Report
                objPopup.ToolTipText = "查阅报告"
                
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 1, "病历格式(&B)"): objControl.IconId = 102
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 6, "报表格式(&P)"): objControl.IconId = 102
                If gobjExchange Is Nothing And mint场合 <> 1 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "我已查阅(&R)")
                        objControl.BeginGroup = True
                    .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "自动标记(&A)"
                End If
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
            If gobjExchange Is Nothing Then
                If mint场合 <> 1 Then
                    Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend * 10# + 2, "打印报告", lngIdx + 1)
                        objPopup.IconId = 103
                        objPopup.Style = xtpButtonIconAndCaption
                        With objPopup.CommandBar.Controls
                            Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告"): objControl.IconId = 102
                            objControl.Style = xtpButtonIconAndCaption
                        End With
                    lngIdx = objPopup.Index
                Else
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "预览报告(&V)", lngIdx + 1)
                    objControl.IconId = 102
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
                End If
            End If
    
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "观片", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "观片处理"
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "关键图像", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "关键图像"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "结果", objControl.Index + 1): objControl.IconId = conMenu_Manage_ReportLisView
                objControl.ToolTipText = "浏览检验结果"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        Else
            If mint场合 = 1 Then '
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Price, "计价", lngIdx + 1)
                objControl.Style = xtpButtonIconAndCaption
                Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "确认停止", objControl.Index + 1)
                objControl.Style = xtpButtonIconAndCaption
                If Not mblnInsideTools Then
                    Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "执行单", objControl.Index + 1): objControl.IconId = 3205
                    objControl.Style = xtpButtonIconAndCaption
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "执行登记", objControl.Index + 1): objControl.IconId = 3587
                    objControl.Style = xtpButtonIconAndCaption
                End If
                Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "核对", objControl.Index + 1)
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            End If
    
            If mint场合 = 0 Then
                If Not mblnInsideTools Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "执行登记", lngIdx + 1): objControl.IconId = 3587
                        objControl.Style = xtpButtonIconAndCaption
                        lngIdx = objControl.Index
                End If
            End If
            
            If mint场合 = 1 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePrice, "对医嘱记帐", lngIdx + 1)
                    objControl.IconId = conMenu_Edit_Price
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
            End If
            If mint场合 <> 2 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "销帐", lngIdx + 1)
                    objControl.IconId = conMenu_Edit_ChargeOff
                    objControl.Style = xtpButtonIconAndCaption
            End If
                
            If mblnPass Then  '合理用药菜单
                Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objCbs(2).Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit, objControl.Index + 1)
            End If
            
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "签名", objControl.Index + 1): objControl.BeginGroup = True
                objControl.IconId = conMenu_Tool_Sign
                objControl.Style = xtpButtonIconAndCaption
        End If
    End With
    
    If mblnInsideTools Then objCbs.RecalcLayout
    
    mcbsMain.RecalcLayout
    
    If mint场合 = 1 And mvarCond.过滤模式 <> 3 Then
        Call SetSendCommandBar
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetRollAdviceIDs(ByVal lng医嘱ID As Long, ByVal bytMode As Byte, Optional ByVal lng操作类型 As Integer, Optional ByVal dat操作时间 As Date, Optional ByVal bln血库相关 As Boolean) As String
'功能：获取需要回退的医嘱记录集
'参数：lng医嘱ID 一组医嘱的组ID
'      bytMode-1.返回一组医嘱记录；2.批量返回所有医嘱记录集(多组)
'      lng操作类型-批量返回时传人
'      dat操作时间-批量返回时传人
'      bln血库相关 启用血库流程时调用传入，对应的操作为 批量回退作废操作

    Dim rsTmp       As ADODB.Recordset
    Dim strSQL      As String
    Dim strTmp As String
    
    On Error GoTo errH

    If bytMode = 1 Then
        strSQL = "Select f_List2str(Cast(Collect(ID || '') As t_Strlist)) As 医嘱ids  From 病人医嘱记录 Where ID =[1] Or 相关id =[1]"
    Else
        strSQL = "Select f_List2str(Cast(Collect(ID || '') As t_Strlist)) As 医嘱ids From 病人医嘱记录 Where" & _
                 " Not (医嘱状态 = 8 And 医嘱期效 = 1) And" & vbNewLine & _
                 "      ID In (Select 医嘱id" & vbNewLine & _
                 "             From 病人医嘱状态" & vbNewLine & _
                 "             Where (操作类型, 操作时间, 操作人员) In (Select 操作类型, 操作时间, 操作人员" & vbNewLine & _
                 "                                          From 病人医嘱状态" & vbNewLine & _
                 "                                          Where 医嘱id = [1] And 操作时间 = [2] And 操作类型 = [3]))"
                 
        If bln血库相关 Then
            strSQL = "Select f_List2str(Cast(Collect(ID || '') As t_Strlist)) As 医嘱ids From 病人医嘱记录 Where" & _
                 " 诊疗类别 = 'K' And 相关id Is Null And" & vbNewLine & _
                 "      ID In (Select 医嘱id" & vbNewLine & _
                 "             From 病人医嘱状态" & vbNewLine & _
                 "             Where (操作类型, 操作时间, 操作人员) In (Select 操作类型, 操作时间, 操作人员" & vbNewLine & _
                 "                                          From 病人医嘱状态" & vbNewLine & _
                 "                                          Where 医嘱id = [1] And 操作时间 = [2] And 操作类型 = [3]))"
        End If

    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID, dat操作时间, lng操作类型)
    If rsTmp.RecordCount = 1 Then GetRollAdviceIDs = rsTmp!医嘱ids & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckAdviceAddModi(Optional ByVal intType As Integer, Optional ByRef lng医嘱ID As Long, Optional ByRef datTurn As Date) As Boolean
'功能：新开和修改时检查是否允许修改或新增
'参数：intType=0-新增，1-修改
    Dim lngBabyEdit As Long
    Dim blnReturn As Boolean
    
    If mlng病人ID = 0 Then Exit Function
    If CheckDataMoved Then Exit Function
    '检查病人是否正在审核
    If Not CheckPatiIsAduit Then Exit Function
    With vsAdvice
        If intType = 1 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
            If lng医嘱ID = 0 Then Exit Function
        
            lngBabyEdit = CheckBabyEdit(Val(.TextMatrix(.Row, COL_婴儿ID)))
            If lngBabyEdit = 1 Then
                MsgBox "当前病人不在本科室，不允许编辑病人医嘱。", vbInformation, gstrSysName
                Exit Function
            ElseIf lngBabyEdit = 2 Then
                MsgBox "当前病人的婴儿不在本科室，不允许编辑婴儿医嘱。", vbInformation, gstrSysName
                Exit Function
            End If
            
            '医技下达的医嘱
            If Val(.TextMatrix(.Row, COL_前提ID)) <> mlng前提ID Then
                MsgBox "不能修改该医嘱,该医嘱是根据其他主医嘱产生的。", vbInformation, gstrSysName
                Exit Function
            End If
            
            '转科病人
            If CheckOtherDeptPatiOpt = False Then Exit Function
            
            '已校对或已废止
            If InStr(",4,8,9,", .TextMatrix(.Row, COL_医嘱状态)) > 0 Then
                MsgBox "当前选择的医嘱已经作废或停止，不能修改。", vbInformation, gstrSysName
                Exit Function
            ElseIf InStr(",1,2,", .TextMatrix(.Row, COL_医嘱状态)) = 0 Then
                MsgBox "当前选择的医嘱已经过校对，不能修改。", vbInformation, gstrSysName
                Exit Function
            End If
            
            '已签名的医嘱不能修改
            If Val(.TextMatrix(.Row, COL_签名否)) = 1 Then
                MsgBox "当前选择的医嘱已经签名，不能修改。请先取消签名。", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mint场合 = 1 Then '护士站调用
                '护士对于已经过审核的医嘱，不允许修改修改
                If .TextMatrix(.Row, COL_开嘱医生) Like "*/*" Then
                    MsgBox "当前选择的医嘱已经过医生审核，不能修改。", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                '无执业资格的医生只能删除修改未审核的医嘱。
                If Not mblnHaveAuditPriv Then
                    If HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_开嘱医生))) Then
                        MsgBox "你没有资格修改当前选择的医嘱，或者当前选择的医嘱已经过审核，不能修改。", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
    
    If mintPState = ps最近转出 Or mintPState = ps预出 Or mintPState = ps出院 Then
        If CheckPatiTurnLimit(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, datTurn, mintPState) = False Then Exit Function
    End If
    CheckAdviceAddModi = True
End Function

Private Sub FuncApplyBlood(ByVal intType As Long)
'功能：输血申请单
'参数：intType=0 新增，=1修改，=2查看 ,=4 核对
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long
    Dim lngNo As Long
    Dim bln用血 As Boolean
    Dim blnApply As Boolean
    
    If intType <> 2 Then
        If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
        If intType = 0 Then
            If Not FuncPathAdd() Then Exit Sub
        End If
        '检查是否满足中级以上专业技术职务
        If gbln输血申请中级以上 Then
            If UserInfo.专业技术职务 <> "主治医师" And UserInfo.专业技术职务 <> "主任医师" And UserInfo.专业技术职务 <> "副主任医师" Then
                MsgBox "启用了输血分级管理后，输血医嘱只有中级及以上专业技术职务医师才能下达。", vbInformation, "输血申请单"
                Exit Sub
            End If
        End If
        '修改时检查是否审核
        If intType = 1 Then
            If Not CanEditBloodAdvice(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_标志)) = 1, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_检查方法)) = 1) Then Exit Sub
        End If
    
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         lngRow = vsAdvice.Row
         bln用血 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_检查方法)) = 1
    End If
    
    If gbln血库系统 = True Then
        blnApply = frmApplyBloodNew.ShowMe(Me, mlng病人ID, mlng主页ID, IIF(mlng病人性质 = 1, 1, 0), intType, lngUpdateAdvice, mlng科室ID, mlng病区ID, Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2), mintPState, datTurn, mrsDefine, mclsMipModule, , , , , mbyt婴儿, , mlng前提ID, IIF(bln用血 = True, 1, 0))
    Else
        blnApply = frmApplyBlood.ShowMe(Me, mlng病人ID, mlng主页ID, IIF(mlng病人性质 = 1, 1, 0), intType, lngUpdateAdvice, mlng科室ID, mlng病区ID, Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2), mintPState, datTurn, mrsDefine, mclsMipModule, , , , , mbyt婴儿, , mlng前提ID)
    End If
    
    If blnApply = True Then
    
        If mlng路径状态 = 1 And Not gobjPath Is Nothing And (intType = 0 Or intType = 1) And lngUpdateAdvice <> 0 Then
            '获取输血申请序号
            lngNo = Sys.RowValue("病人医嘱记录", lngUpdateAdvice, "申请序号", "ID")
            If lngNo <> 0 Then Call FuncPathSet(lngNo)
        End If
        '刷新医嘱
        Call RefreshData
        '选择最后一行医嘱
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_医嘱内容
    End If
End Sub

Private Sub FuncApplyOperation(ByVal intType As Long)
'功能：手术申请单
'参数：intType=0 新增，=1修改，=2查看
    Dim lngUpdateAdvice As Long
    Dim datTurn As Date
    Dim lngRow As Long, strDefine As String
    Dim lng开嘱科室ID As Long
    Dim lngNo As Long
    
    If intType <> 2 Then
        If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
        '修改时检查是否审核
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)) = 2 Then
                MsgBox "申请单已经审核，不允许再修改。", vbInformation, "手术申请单"
                intType = 2
            End If
        End If
        If intType = 0 Then
            If Not FuncPathAdd() Then Exit Sub
        End If
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         lngRow = vsAdvice.Row
    End If
    
    lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2)
    If Not mrsDefine Is Nothing Then
        mrsDefine.Filter = "诊疗类别='F'"
        If Not mrsDefine.EOF Then strDefine = Trim(NVL(mrsDefine!医嘱内容))
    End If

    If frmApplyOperation.ShowMe(Me, 0, intType, mlng病人ID, mlng主页ID, IIF(mlng病人性质 = 1, 1, 0), lngUpdateAdvice, mlng科室ID, lng开嘱科室ID, strDefine, mlng病区ID, mintPState, datTurn, 0, mclsMipModule, , , mlng前提ID, mbyt婴儿) Then
        '刷新医嘱
        Call RefreshData
        '选择最后一行医嘱
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_医嘱内容
        
        If mlng路径状态 = 1 And Not gobjPath Is Nothing And lngUpdateAdvice <> 0 Then
            lngNo = Sys.RowValue("病人医嘱记录", lngUpdateAdvice, "申请序号", "ID")
            If lngNo <> 0 Then Call FuncPathSet(lngNo)
        End If
    End If
End Sub

Private Sub FuncApplyConsultation(ByVal intType As Long)
'功能：会诊申请单
'参数：intType=0 新增，=1修改，=2查看
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long, lngNo As Long
    Dim lng开嘱科室ID As Long

    If Not CheckWindow Then Exit Sub
    
    If intType <> 2 Then
        If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         lngRow = vsAdvice.Row
    End If
    
    lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2)
    Set mfrmEac = frmApplyConsultation
    If frmApplyConsultation.ShowMe(mfrmParent, lngUpdateAdvice, lngNo, intType, 0, mlng病人ID, mlng主页ID, mlng科室ID, lng开嘱科室ID, mlng病区ID, mintPState, datTurn, mclsMipModule, , , mlng前提ID, mbyt婴儿) Then
        '刷新医嘱
        Call RefreshData
        '选择最后一行医嘱
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_医嘱内容
    End If
    
End Sub

Private Sub zlPASSMap()
'功能:设置Pass VsAdvie及列映射
'注意:删除或修改下面列中数据时，请检查合理用药部件中的关联处理。
    Dim blnTmp As Boolean
    
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "合理用药监测", True)
    End If
    
    If gobjPass Is Nothing Then
        blnTmp = False
    Else
        blnTmp = gobjPass.PassType <> UNPASS
    End If
    
    mblnPass = blnTmp And Not mobjPassMap Is Nothing
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_住院医嘱清单
            .int场合 = mint场合
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .VSCOL = .GetVSCOL(COL_ID, COL_相关ID, COL_诊疗类别, _
                COL_诊疗项目ID, COL_收费细目ID, col_医嘱内容, COL_期效, COL_单量, COL_单量单位, COL_用法, COL_天数, , COL_开嘱时间, COL_开嘱医生, _
                COL_开始时间, COL_开嘱科室ID, COL_终止时间, COL_频率, , , , COL_警示, COL_序号, COL_医嘱状态, , , , , COL_执行性质, COL_标本部位, _
                , , , , , COL_总量, , COL_医生嘱托, COL_用药目的, COL_操作类型)
            Set .PassPati = .GetPatient()
            mblnPass = gobjPass.zlPassCheck(mobjPassMap)
        End With
    End If
End Sub

Private Sub zlPASSPati()
'功能:设置病人信息
    
    With mobjPassMap.PassPati
        .lng病人ID = mlng病人ID
        .lng主页ID = mlng主页ID
    End With
End Sub

Public Sub LocatedAdviceRow(ByVal lng医嘱ID As Long)
'功能：定位医嘱行
    Dim blnExist As Boolean
    Dim i As Long
    
    i = vsAdvice.FindRow(CStr(lng医嘱ID), , COL_ID)
    If i = -1 Then Exit Sub
    If vsAdvice.RowHidden(vsAdvice.Row) Then Exit Sub    '定位到了隐藏行的处理
    vsAdvice.Row = i
    Call vsAdvice.ShowCell(i, vsAdvice.FixedCols)
    Call ShowAdvicePlan(i, blnExist) '如果是安排类医嘱定位到安排情况页签
    If blnExist Then
        For i = 0 To tbcAppend.ItemCount - 1
            If tbcAppend(i).Tag = "安排" Then
                tbcAppend(i).Selected = True
                Exit For
            End If
        Next
    End If
End Sub

Private Sub SetCISMsg(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal lng医嘱ID As Long, ByVal lng紧急 As Long)
'功能：产生医嘱新停消息，在医嘱回退确认停止时调用
'参数：lng紧急 1 表示紧急医嘱 0 非紧急医嘱
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "select 1 From 业务消息清单 A Where a.病人id=[1] And a.就诊id=[2] And a.类型编码 ='ZLHIS_CIS_002' And a.优先程度=[3] And a.是否已阅=0 And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, IIF(lng紧急 = 1, 2, 1))
    If rsTmp.EOF Then
        strSQL = "Select a.病人性质 As 性质,a.出院科室id As 科室id, a.当前病区id As 病区id From 病案主页 A Where a.病人id =[1] And a.主页id =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        strSQL = "Zl_业务消息清单_Insert(" & lng病人ID & "," & lng主页ID & "," & rsTmp!科室ID & "," & rsTmp!病区ID & "," & IIF(rsTmp!性质 = 1, 1, 2) & ",'有新停止医嘱。','0010','ZLHIS_CIS_002'," & _
            lng医嘱ID & "," & IIF(lng紧急 = 1, 2, 1) & ",0,null," & rsTmp!病区ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetTag一并给药(Optional ByVal lngRow As Long)
'功能：在一并给药的医嘱前加标志
    Dim i As Long
    Dim lngBg As Long, lngEd As Long
    Dim j As Long
    Dim lngStart As Long, lngEnd As Long

    If mvarCond.过滤模式 = 3 Then Exit Sub

    With vsAdvice
        If lngRow = 0 Then
            lngStart = .FixedRows
            lngEnd = .Rows - 1
        Else
            lngStart = lngRow
            lngEnd = lngRow
        End If
        For i = lngStart To lngEnd
             lngBg = -1: lngEd = -1
             If RowIn一并给药(i, lngBg, lngEd) Then
                For j = lngBg To lngEd
                    If j = lngBg Then
                        .TextMatrix(j, COL_并) = "┏"
                    ElseIf j = lngEd Then
                        .TextMatrix(j, COL_并) = "┗"
                    Else
                        .TextMatrix(j, COL_并) = "┃"
                    End If
                Next
                If lngEd <> -1 Then
                   i = lngEd + 1
                End If
            End If
        Next
    End With
End Sub

Private Sub DefInSidePlugInBar(ByRef rsBar As ADODB.Recordset)
'功能：新版护士站调用时，内部工具条的改
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim i As Long
    Dim lngTmp As Long
    Dim blnGroup As Boolean
    
    If Not mblnInsideTools Then Exit Sub
    If rsBar Is Nothing Then Exit Sub
    rsBar.Filter = "BarType=2"
    If rsBar.RecordCount = 0 Then Exit Sub
    rsBar.Filter = "IsInTool=1 and BarType=2"
    '工具栏按钮
    rsBar.Sort = "序号 desc"
    Set objBar = cbsSub(2)
    lngTmp = -1
    With objBar.Controls
        If Not rsBar.EOF Then
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!功能名, lngTmp)
                    objControl.IconId = rsBar!图标ID
                    objControl.Parameter = rsBar!功能名
                    objControl.Style = xtpButtonIconAndCaption
                lngTmp = objControl.Index
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                rsBar.MoveNext
            Next
            objControl.BeginGroup = True
        End If
    End With
    rsBar.Filter = "IsInTool=0 and BarType=2"
    If Not rsBar.EOF Then
        rsBar.Sort = "序号"
        Set objPopup = objBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "扩展功能", , False)
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!功能ID, rsBar!菜单名)
                objControl.IconId = rsBar!图标ID
                objControl.Parameter = rsBar!功能名
                objControl.Style = xtpButtonIconAndCaption
                If Val(rsBar!IsGroup) = 1 Then
                    objControl.BeginGroup = True
                End If
                rsBar.MoveNext
            Next
        End With
    End If
        cbsSub.RecalcLayout
End Sub

Private Function ShowAdviceRISSch(ByVal lngRow As Long, ByRef blnExist As Boolean) As Boolean
'功能：显示指定行的预约信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    Dim i As Long
    
    blnExist = False
    rtfSche.Text = "": rtfSche.SelStart = 0
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_RIS预约ID)) = 0 Then Exit Function
    
    strSQL = "select 检查设备名称,To_Char(预约日期,'YYYY-MM-DD') as 预约日期," & vbNewLine & _
        "To_Char(预约开始时间,'YYYY-MM-DD HH24:MI:SS') as 预约开始时间," & vbNewLine & _
        "To_Char(预约结束时间,'YYYY-MM-DD HH24:MI:SS') as 预约结束时间," & vbNewLine & _
        "To_Char(预约开始时间段,'YYYY-MM-DD HH24:MI:SS') as 预约开始时间段," & vbNewLine & _
        "To_Char(预约结束时间段,'YYYY-MM-DD HH24:MI:SS') as 预约结束时间段,DECODE(是否调整,1,'已经预约调整','已经预约') as 预约状态" & vbNewLine & _
        "from RIS检查预约 Where 医嘱ID=[1]"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfSche
            For i = 0 To rsTmp.Fields.Count - 1
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp.Fields(i).Name & "：" & NVL(rsTmp.Fields(i).value)
                lngIdx = .Find(rsTmp.Fields(i).Name & "：", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp.Fields(i).Name & "：")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
            Next
            '光标定位在第一个
            lngIdx = .Find(rsTmp.Fields(0).Name & "：", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp.Fields(0).Name & "：")
            Call SetRTFFont(4)
        End With
        blnExist = True
    End If
    ShowAdviceRISSch = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceRISSch()
'功能：RIS医嘱预约
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    lngResult = -1
    If HaveRIS Then
        With vsAdvice
            If InStr(",1,3,8,", "," & .TextMatrix(.Row, COL_医嘱状态) & ",") > 0 Then
                lngResult = gobjRis.HISScheduling(2, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_诊疗项目ID)))
                If lngResult = 0 Then
                    '成功预约后更新状态
                    strSQL = "select min(预约ID) as ID from RIS检查预约 where 医嘱id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                    .TextMatrix(.Row, COL_RIS预约ID) = rsTmp!ID & ""
                End If
            Else
                MsgBox "医嘱状态为新开、校对、已发送时，才能预约！", vbInformation, gstrSysName
            End If
        End With
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncAdviceRISDel()
'功能：RIS医嘱取消预约
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngResult As Long
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_RIS预约ID)) <> 0 Then
            If Val(.TextMatrix(.Row, COL_医嘱状态)) = 8 Then
                strSQL = "Select Max(b.执行状态) As 结果 From 病人医嘱记录 A, 病人医嘱发送 B Where a.Id = b.医嘱id And (a.Id =[1] Or a.相关id=[1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                If Not rsTmp.EOF Then
                    If Val(rsTmp!结果 & "") = 0 Then
                        blnDo = True
                    Else
                        MsgBox "该医嘱已经被执行或者部分执行不能取消预约！", vbInformation, gstrSysName
                    End If
                End If
            Else
                blnDo = True
            End If
        End If
        If blnDo Then
            If HaveRIS Then
                lngResult = gobjRis.HISSchedulingEx(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_RIS预约ID)))
                If lngResult = 0 Then
                    '成功能取消更改状态
                    .TextMatrix(.Row, COL_RIS预约ID) = ""
                    Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)
                End If
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetAdviceReportIcon(ByVal lngRow As Long)
'功能：根据当前行的内容设置医嘱报告列的图标标识
'说明：注意是单行设置，不是一组设置
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_报告ID)) <> 0 Or _
                .TextMatrix(lngRow, COL_检查报告ID) <> "" Or _
                Val(.TextMatrix(lngRow, COL_RIS报告ID)) <> 0 Or _
                Val(.TextMatrix(lngRow, COL_LIS报告ID)) <> 0 Then
                
                
                If Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 0 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("报告").Picture
                ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 1 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("报告已阅").Picture
                ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 2 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("报告部分阅").Picture
                End If
            Else
                If Val(.TextMatrix(lngRow, COL_RIS预约ID)) <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F报告) = frmIcons.imgFlag.ListImages("预约").Picture
                End If
            End If
        End If
    End With
End Sub

Private Sub FuncPathSet(ByVal lng申请序号 As Long)
'功能:检查是否属于路径内项目
'True-路径内项目;False-非路径内项目
    Dim byt生成时间性质 As Byte
    Dim lng路径项目ID As Long, lng阶段Id As Long, lng天数 As Long
    Dim i As Long, k As Long, lng执行ID As Long
    Dim strSQL As String
    Dim str诊疗项目IDs As String, str路径项目分类 As String
    Dim strAdvices As String, str组ID As String, strList As String, strAdvicesOut As String
    Dim str分类 As String
    Dim str期效 As String
    Dim str开始日期 As String, strAddDate As String
    Dim dat日期 As Date, DatAddDate As Date
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rsPath As ADODB.Recordset
    Dim rsStep As ADODB.Recordset
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    Dim blnTrans As Boolean
    Dim blnPathOut As Boolean
    
    On Error GoTo errH:
    
    If Not (mint场合 = 0 Or mint场合 = 2) Then Exit Sub
    
    strSQL = "Select a.Id As 医嘱id,a.相关ID, Nvl(a.相关id, a.Id) As 组ID, a.诊疗项目id,a.诊疗类别, b.操作类型, a.开始执行时间, a.医嘱期效" & vbNewLine & _
            "From 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
            "Where a.诊疗项目id = b.Id And a.申请序号 = [1]" & vbNewLine & _
            "Order By a.序号"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng申请序号)
    If rsAdvice.RecordCount = 0 Then Exit Sub

    For i = 1 To rsAdvice.RecordCount
    '药品不含给药途径、用法、煎法，输血不含途径,检验不含采集方式，手术不含附加手术、麻醉，检查不含部位方法
        If i = 1 Then str组ID = rsAdvice!组ID & ""
        If str组ID <> rsAdvice!组ID & "" Then
            str组ID = rsAdvice!组ID & ""
            strAdvices = Mid(strAdvices, 2)
            str诊疗项目IDs = Mid(str诊疗项目IDs, 2)
            strList = strList & "&" & strAdvices & "|" & str诊疗项目IDs & "|" & str期效
            strAdvices = ""
            str诊疗项目IDs = ""
            str期效 = ""
        End If
        strAdvices = strAdvices & "," & rsAdvice!医嘱ID
        If Not (rsAdvice!诊疗类别 & "" = "E" And InStr(",2,3,4,6,8,", "," & rsAdvice!操作类型 & ",") > 0) And Not (InStr(",G,F,D,", "," & rsAdvice!诊疗类别 & ",") > 0 And NVL(rsAdvice!相关ID, 0) <> 0) Then
            str诊疗项目IDs = str诊疗项目IDs & "," & rsAdvice!诊疗项目ID
            If str开始日期 = "" Then str开始日期 = Format(rsAdvice!开始执行时间, "YYYY-MM-DD")
            If str期效 = "" Then str期效 = rsAdvice!医嘱期效
        End If
        rsAdvice.MoveNext
    Next
    strAdvices = Mid(strAdvices, 2)
    str诊疗项目IDs = Mid(str诊疗项目IDs, 2)
    If InStr(strList, strAdvices & "|" & str诊疗项目IDs & "|" & str期效) = 0 Then strList = strList & "&" & strAdvices & "|" & str诊疗项目IDs & "|" & str期效
    strList = Mid(strList, 2)
    arrTmp = Split(strList, "&")
    '获取路径当前阶段;当前日期
    Set rsPath = GetPatiPathInfo(mlng病人ID, mlng主页ID, str路径项目分类)
    If rsPath.RecordCount = 0 Then Exit Sub
    If rsPath!日期 = CDate(str开始日期) Then
        strSQL = "Select 当前阶段id From 病人临床路径 Where 病人ID = [1] And 主页ID=[2]" & vbNewLine & _
                        "Union All" & vbNewLine & _
                        "Select 当前阶段id From 病人合并路径 Where 病人ID = [1] And 主页ID=[2]"
    Else
        '根据医嘱的开始日期获取对应的路径阶段
        strSQL = "Select a.阶段id as 当前阶段ID " & vbNewLine & _
                "From 病人路径执行 A, 病人临床路径 B" & vbNewLine & _
                "Where b.Id = a.路径记录id And b.病人id = [1] And b.主页id = [2] And a.日期 = [3]" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select a.阶段id as 当前阶段ID  " & vbNewLine & _
                "From 病人路径执行 A, 病人合并路径 B" & vbNewLine & _
                "Where b.Id = a.合并路径记录id And b.病人id = [1] And b.主页id = [2] And a.日期 = [3]"
    End If
    Set rsStep = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, CDate(str开始日期))
    DatAddDate = zlDatabase.Currentdate
    strAddDate = "To_Date('" & Format(DatAddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrSQL = Array()
    
    '检验检查一个申请序号对应多组医嘱
    For k = LBound(arrTmp) To UBound(arrTmp)
        strAdvices = Split(arrTmp(k), "|")(0)
        str诊疗项目IDs = Split(arrTmp(k), "|")(1)
        str期效 = Split(arrTmp(k), "|")(2)
        If rsStep.RecordCount > 0 Then rsStep.MoveFirst    '当医嘱开始日期大于路径的当前阶段的当前日期时,rsSetp返回记录集为0,默认为路径外项目
        Do While Not rsStep.EOF
            If rsStep!当前阶段ID & "" <> "" Then
                lng路径项目ID = CheckPathInItem(mlng病人ID, mlng主页ID, str诊疗项目IDs, str分类, Val(rsStep!当前阶段ID & ""), False, CByte(str期效))
            End If
            If lng路径项目ID <> 0 Then Exit Do
            rsStep.MoveNext
        Loop
        '路径表单关联处理
        If lng路径项目ID = 0 Then
            blnPathOut = True
            strAdvicesOut = strAdvicesOut & "," & strAdvices
        Else
            '路径内项目
            If rsPath!日期 > CDate(str开始日期) Then
                Set rsTmp = GetPatiPathAppend(rsPath!路径记录id, CDate(str开始日期))
                If rsTmp.RecordCount > 0 Then
                    lng阶段Id = rsTmp!阶段id
                    lng天数 = rsTmp!天数
                    dat日期 = CDate(str开始日期)
                End If
                byt生成时间性质 = 1 '补录
            Else
                lng阶段Id = rsPath!当前阶段ID
                lng天数 = rsPath!当前天数
                dat日期 = rsPath!日期
                If rsPath!日期 = CDate(str开始日期) Then
                    byt生成时间性质 = 0
                Else
                    byt生成时间性质 = 2 '暂存
                End If
            End If
        
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人路径生成_Insert(0," & mlng病人ID & "," & mlng主页ID & ",NULL,0," & rsPath!路径记录id & "," & lng阶段Id & _
                                     ",To_Date('" & Format(dat日期, "yyyy-MM-dd") & "','YYYY-MM-DD')," & lng天数 & _
                                     ",'" & str分类 & "'," & lng路径项目ID & ",'" & strAdvices & "',Null,Null" & _
                                     ",'" & UserInfo.姓名 & "'," & strAddDate & ",NULL,1,Null,Null,Null,NUlL," & IIF(byt生成时间性质 = 1, 1, 0) & ",Null,Null,NUlL,Null,Null,Null,NUlL," & IIF(byt生成时间性质 = 0, "NULL", byt生成时间性质) & ")"
            
        End If
    Next
    
    If blnPathOut Then
        '路径外项目
        strAdvicesOut = Mid(strAdvicesOut, 2)
        If strAdvicesOut <> "" Then
            If rsPath!日期 > CDate(str开始日期) Then
                Set rsTmp = GetPatiPathAppend(rsPath!路径记录id, CDate(str开始日期))
                If rsTmp.RecordCount > 0 Then
                    lng阶段Id = rsTmp!阶段id
                    lng天数 = rsTmp!天数
                    dat日期 = CDate(str开始日期)
                End If
                byt生成时间性质 = 1 '补录
            Else
                lng阶段Id = rsPath!当前阶段ID
                lng天数 = rsPath!当前天数
                dat日期 = rsPath!日期
                If rsPath!日期 = CDate(str开始日期) Then
                    byt生成时间性质 = 0
                Else
                    byt生成时间性质 = 2 '暂存
                End If
            End If
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人路径生成_Insert(0," & mlng病人ID & "," & mlng主页ID & ",Null,0," & _
                                      rsPath!路径记录id & "," & lng阶段Id & ",To_Date('" & Format(dat日期, "yyyy-MM-dd") & "','YYYY-MM-DD')," & lng天数 & _
                                      ",'" & str路径项目分类 & "',Null" & ",'" & strAdvicesOut & "',Null,Null,'" & UserInfo.姓名 & "'," & strAddDate & ",'路径外项目'" & _
                                      ",1,Null,Null,Null,NUlL," & IIF(byt生成时间性质 = 1, 1, 0) & ",Null,Null,NUlL,Null,Null,Null,NUlL," & IIF(byt生成时间性质 = 0, "NULL", byt生成时间性质) & ")"
        End If
    End If
    '数据提交
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    '多个路径外项目一并处理
    If blnPathOut Then
        lng执行ID = GetPathOutItemID(Val(rsPath!路径记录id), DatAddDate)
        '强制刷新读取路径病人信息，因为界面可能切换病人
        Call gobjPath.zlRefresh(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, mintPState, False, True)
        Call gobjPath.zlExePathAppendItem(str路径项目分类, strAdvicesOut, lng执行ID, dat日期)
    End If
    
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FuncPathAdd() As Boolean
    Dim strSQL As String
    Dim str当前日期 As String
    Dim i As Long
    Dim lng疾病ID As Long, lng诊断ID As Long
    Dim bln中医 As Boolean
    Dim blnDo As Boolean, blnIsCancel As Boolean
    Dim blnIsSend As Boolean, blnYes As Boolean
    Dim rsTmp As ADODB.Recordset, rsPath As ADODB.Recordset
    Dim objDiagEdit As zlMedRecPage.clsDiagEdit
    
    '检查病人是否下达了出院医嘱
    If mstr婴儿 = "" And mlng路径状态 = 2 Then
        If CheckOutAdvice(mlng病人ID, mlng主页ID) Then
            MsgBox "该病人已经正常结束了路径并下达了出院医嘱，不能再新开医嘱。", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    
     '当前是新病人，未下过医嘱的，当前科室有可用的路径表单时，并且还未填写入院或门诊诊断的，提示先填写入院诊断。
    If mlng路径状态 = -1 And mlng病人性质 <> 1 Then
        If InStr(GetInsidePrivs(p临床路径应用), ";导入路径;") > 0 Then
            On Error GoTo errH
            strSQL = "select 1 From 病人医嘱记录 Where 病人ID=[1] and 主页ID=[2] and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
            If rsTmp.RecordCount = 0 Then
                If HavePath(mlng科室ID) Then
                    strSQL = "select 1 From 病人诊断记录 Where 诊断类型 In (1, 2, 11, 12) And 记录来源 = 3 And 病人ID=[1] and 主页ID=[2] and rownum<2"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
                    If rsTmp.RecordCount = 0 Then
                        If MsgBox("本科室有可用的临床路径表单，为了及时导入临床路径，请问是否填写入院诊断？", vbQuestion + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                            If objDiagEdit Is Nothing Then
                                Set objDiagEdit = New zlMedRecPage.clsDiagEdit
                                Call objDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng病人性质 = 1, 1260, 1261), mclsMipModule)
                            End If
                            If objDiagEdit.ShowDiagEdit(Me, 0, mlng病人ID, mlng主页ID, IIF(mlng病人性质 = 1, 1, 2), mlng科室ID, "", "", 0) Then
                                Set rsTmp = Get病种ID(mlng病人ID, mlng主页ID, mlng科室ID, bln中医)
                                If bln中医 Then
                                    rsTmp.Filter = "诊断类型 =12 OR 诊断类型 = 2 "
                                    For i = 1 To rsTmp.RecordCount
                                        lng疾病ID = Val("" & rsTmp!疾病id)
                                        lng诊断ID = Val("" & rsTmp!诊断id)
                                        Set rsPath = GetPathTable(lng疾病ID, lng诊断ID, mlng科室ID)
                                        If rsPath.RecordCount > 0 Then Exit For
                                        rsTmp.MoveNext
                                    Next
                                Else
                                    If rsTmp.RecordCount > 0 Then
                                        lng疾病ID = Val("" & rsTmp!疾病id)
                                        lng诊断ID = Val("" & rsTmp!诊断id)
                                    End If
                                    Set rsPath = GetPathTable(lng疾病ID, lng诊断ID, mlng科室ID)
                                End If
                                If Not rsPath Is Nothing Then
                                    If rsPath.RecordCount > 0 Then
                                        Call gobjPath.zlRefresh(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, mintPState, mblnMoved, True)
                                        Call gobjPath.zlImportPath
                                        RaiseEvent RequestRefresh(False)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    '路径中的病人，当天没有生成路径项目，则先调用生成
    If mlng路径状态 = 1 And mvarCond.婴儿 <= 0 Then
        blnDo = True
        If mint场合 = 2 Then
            blnDo = zlDatabase.GetPara("医技医嘱在路径表外", glngSys, p临床路径应用, 0) = 0
        End If
        '未评估时允许添加医嘱到昨天
        mblnNotEvaluete = Val(zlDatabase.GetPara("未评估时允许添加医嘱到昨天", glngSys, p临床路径应用, 1)) = 1
        
        If blnDo Then
            If CheckPathNotEvaluete(mlng病人ID, mlng主页ID, blnIsSend, str当前日期) = False Then
                If gobjPath Is Nothing Then
                    MsgBox "该病人当天当前阶段的路径项目未生成，不能新开医嘱。", vbInformation, gstrSysName
                ElseIf InStr(GetInsidePrivs(p临床路径应用), ";生成路径;") = 0 Then
                    MsgBox "该病人当天当前阶段的路径项目未生成，你没有生成路径的权限，不能新开医嘱。", vbInformation, gstrSysName
                Else
                    '之前可能没有进过路径页面，需要先通过刷新接口读取初始数据
                    Call gobjPath.zlRefresh(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, mintPState, mblnMoved, True)
                    Call gobjPath.zlExecPathSend(blnIsCancel)
                    Call LoadAdvice
                End If
                If Not blnIsCancel Then Exit Function
             Else
                If Not blnIsSend Then
                    If gobjPath Is Nothing Then
                        MsgBox "该病人当天当前阶段的路径项目未生成，不能新开医嘱。", vbInformation, gstrSysName
                        Exit Function
                    ElseIf InStr(GetInsidePrivs(p临床路径应用), ";生成路径;") = 0 Then
                        MsgBox "该病人当天当前阶段的路径项目未生成，你没有生成路径的权限，不能新开医嘱。", vbInformation, gstrSysName
                        Exit Function
                    Else
                        '如果启用了参数：未评估时允许添加医嘱到昨天，则提示，否则直接进行评估生成操作
                        If mblnNotEvaluete Then
                            blnYes = MsgBox("你要添加路径外项目到''" & str当前日期 & "'?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
                        End If
                        '如果选择否，则进行评估生成操作，选择是则允许新开路径外项目到 当前日期
                        If blnYes = False Then
                            '之前可能没有进过路径页面，需要先通过刷新接口读取初始数据
                            Call gobjPath.zlRefresh(mlng病人ID, mlng主页ID, mlng病区ID, mlng科室ID, mintPState, mblnMoved, True)
                            '没有生成，则返回false禁止新开操作
                            If Not gobjPath.zlExecPathSend Then
                                Call LoadAdvice
                                Exit Function
                            End If
                            Call LoadAdvice
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    FuncPathAdd = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub AddDataToVsf(ByVal rsData As ADODB.Recordset)
'功能：医嘱表格中加数据
    Dim i As Long
    With vsAdvice
        .Rows = rsData.RecordCount + 1
        For i = 1 To rsData.RecordCount
            .TextMatrix(i, COL_ID) = NVL(rsData!ID)
            .TextMatrix(i, COL_相关ID) = NVL(rsData!相关ID)
            .TextMatrix(i, COL_序号) = NVL(rsData!序号)
            .TextMatrix(i, COL_婴儿ID) = NVL(rsData!婴儿ID)
            .TextMatrix(i, COL_医嘱状态) = NVL(rsData!医嘱状态)
            .TextMatrix(i, COL_诊疗类别) = NVL(rsData!诊疗类别)
            .TextMatrix(i, COL_操作类型) = NVL(rsData!操作类型)
            .TextMatrix(i, COL_毒理分类) = NVL(rsData!毒理分类)
            .TextMatrix(i, COL_标志) = NVL(rsData!标志)
            .TextMatrix(i, COL_警示) = NVL(rsData!警示)
            .TextMatrix(i, COL_期效) = NVL(rsData!期效)
            .TextMatrix(i, COL_开始时间) = NVL(rsData!开始时间)
            .TextMatrix(i, COL_并) = NVL(rsData!并)
            .TextMatrix(i, col_医嘱内容) = NVL(rsData!医嘱内容)
            .TextMatrix(i, col_内容) = NVL(rsData!内容)
            .TextMatrix(i, COL_皮试) = NVL(rsData!皮试)
            .TextMatrix(i, COL_总量) = NVL(rsData!总量)
            .TextMatrix(i, COL_单量) = NVL(rsData!单量)
            .TextMatrix(i, COL_天数) = NVL(rsData!天数)
            .TextMatrix(i, COL_频率) = NVL(rsData!频率)
            .TextMatrix(i, COL_用法) = NVL(rsData!用法)
            .TextMatrix(i, COL_医生嘱托) = NVL(rsData!医生嘱托)
            .TextMatrix(i, COL_执行时间) = NVL(rsData!执行时间)
            .TextMatrix(i, COL_终止时间) = NVL(rsData!终止时间)
            .TextMatrix(i, COL_执行科室) = NVL(rsData!执行科室)
            .TextMatrix(i, COL_执行性质) = NVL(rsData!执行性质)
            .TextMatrix(i, COL_上次执行) = NVL(rsData!上次执行)
            .TextMatrix(i, COL_状态) = NVL(rsData!状态)
            .TextMatrix(i, COL_开嘱医生) = NVL(rsData!开嘱医生)
            .TextMatrix(i, COL_开嘱时间) = NVL(rsData!开嘱时间)
            .TextMatrix(i, COL_校对护士) = NVL(rsData!校对护士)
            .TextMatrix(i, COL_校对时间) = NVL(rsData!校对时间)
            .TextMatrix(i, COL_停嘱医生) = NVL(rsData!停嘱医生)
            .TextMatrix(i, COL_停嘱时间) = NVL(rsData!停嘱时间)
            .TextMatrix(i, COL_停嘱护士) = NVL(rsData!停嘱护士)
            .TextMatrix(i, COL_确认停嘱时间) = NVL(rsData!确认停嘱时间)
            .TextMatrix(i, COL_基本药物) = NVL(rsData!基本药物)
            .TextMatrix(i, COL_查阅状态) = NVL(rsData!查阅状态)
            .TextMatrix(i, COL_标本状态) = NVL(rsData!标本状态)
            .TextMatrix(i, COL_诊疗项目ID) = NVL(rsData!诊疗项目ID)
            .TextMatrix(i, COL_试管编码) = NVL(rsData!试管编码)
            .TextMatrix(i, COL_执行标记) = NVL(rsData!执行标记)
            .TextMatrix(i, COL_屏蔽打印) = NVL(rsData!屏蔽打印)
            .TextMatrix(i, COL_前提ID) = NVL(rsData!前提ID)
            .TextMatrix(i, COL_签名否) = NVL(rsData!签名否)
            .TextMatrix(i, COL_文件ID) = NVL(rsData!文件ID)
            .TextMatrix(i, COL_报告项) = NVL(rsData!报告项)
            .TextMatrix(i, COL_报告ID) = NVL(rsData!报告ID)
            .TextMatrix(i, COL_收费细目ID) = NVL(rsData!收费细目ID)
            .TextMatrix(i, COL_单量单位) = NVL(rsData!单量单位)
            .TextMatrix(i, COL_开嘱科室ID) = NVL(rsData!开嘱科室id)
            .TextMatrix(i, COL_审核状态) = NVL(rsData!审核状态)
            .TextMatrix(i, COL_申请序号) = NVL(rsData!申请序号)
            .TextMatrix(i, COL_审核标记) = NVL(rsData!审核标记)
            .TextMatrix(i, COL_高危药品) = NVL(rsData!高危药品)
            .TextMatrix(i, COL_标本部位) = NVL(rsData!标本部位)
            .TextMatrix(i, COL_用药目的) = NVL(rsData!用药目的)
            .TextMatrix(i, COL_检查报告ID) = NVL(rsData!检查报告ID)
            .TextMatrix(i, COL_处方审查状态) = NVL(rsData!处方审查状态)
            .TextMatrix(i, COL_处方审查结果) = NVL(rsData!处方审查结果)
            .TextMatrix(i, COL_RIS预约ID) = NVL(rsData!RIS预约ID)
            .TextMatrix(i, COL_RIS报告ID) = NVL(rsData!RIS报告ID)
            .TextMatrix(i, COL_LIS报告ID) = NVL(rsData!LIS报告ID)
            .TextMatrix(i, COL_RIS预约状态) = NVL(rsData!RIS预约状态)
            .TextMatrix(i, col_诊疗项目名称) = NVL(rsData!诊疗项目名称)
            .TextMatrix(i, COL_检查方法) = NVL(rsData!检查方法)
            .TextMatrix(i, COL_危急值ID) = NVL(rsData!危急值ID)
            .TextMatrix(i, COL_易跌倒) = Val(rsData!是否易至跌倒 & "")
            rsData.MoveNext
        Next
    End With
End Sub

Private Sub FuncAdviceRISPrintSch(ByVal lngFunID As Long)
'功能：RIS医嘱预约单打印
'参数：lngFunID 功能ID； lngFunID ＝conMenu_Tool_RisPrint－打印单个预约单，lngFunID ＝conMenu_Tool_RisPrintBat－批量打印
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strName As String
    
    On Error GoTo errH
    
    lngResult = -1
    If HaveRIS Then
        If lngFunID = conMenu_Tool_RisPrint Then
            With vsAdvice
                If Not .TextMatrix(.Row, COL_诊疗类别) = "D" Then
                    MsgBox "当前医嘱不是影像检查项目。", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .TextMatrix(.Row, COL_RIS预约ID) = 0 Then
                    MsgBox "当前影像检查医嘱没有被预约，不能打印。", vbInformation, gstrSysName
                    Exit Sub
                End If
                lngResult = gobjRis.HISPrintOneRisScheduleRpt(Val(.TextMatrix(.Row, COL_ID)))
            End With
        Else
            Call frmAdviceRisReport.ShowMe(Me, mlng病区ID)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsUseBloodAdvice() As Boolean
'功能：判断选中的医嘱是否是用血医嘱
    Dim i As Long
    Dim blnTrue As Boolean
    
    With vsAppend
        For i = .FixedRows To .Rows - 1
           If .TextMatrix(i, COLSend("诊疗类别")) = "K" Then
                blnTrue = (.TextMatrix(i, COLSend("输血类型")) = "1")
                Exit For
           End If
        Next
    End With
    IsUseBloodAdvice = blnTrue
End Function

Private Function HaveItemToRis(ByVal lng发送号 As Long, ByRef lng医嘱ID As Long) As Boolean
'功能：按发送号过滤本次发送的医嘱中是否有发到RIS去的医嘱没有
'说明：与RIS相关，在批量回退发送操作时调用，当本次发送中有>=2条医嘱时则禁止批量回，必须单独回退。因为RIS那边一次只能处理一条医嘱。
    Dim strSQL As String
    Dim strIDs As String, i As Long
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "Select a.id,a.诊疗项目id,0 as RISItem" & vbNewLine & _
        "From 病人医嘱记录 A, 病人医嘱发送 B, 诊疗项目目录 C" & vbNewLine & _
        "Where a.Id = b.医嘱id And a.诊疗项目id = c.Id And b.发送号 =[1] And" & vbNewLine & _
        "      (a.诊疗类别 In ('F', 'D') Or a.诊疗类别 = 'E' And Nvl(c.操作类型,'0') in ('5','0')) And a.相关id Is Null And a.医嘱期效 = 1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng发送号)
    Set rsTmp = zlDatabase.CopyNewRec(rsTmp)
    
    For i = 1 To rsTmp.RecordCount
        If InStr("," & strIDs & ",", "," & rsTmp!诊疗项目ID & ",") = 0 Then
            strIDs = strIDs & "," & rsTmp!诊疗项目ID
        End If
        rsTmp.MoveNext
    Next
    
    If HaveRIS(False) Then
        On Error Resume Next
        strTmp = gobjRis.HISIsRisItem(Mid(strIDs, 2))
        err.Clear: On Error GoTo errH
    End If
    If strTmp <> "" Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            If InStr("," & strTmp & ",", "," & rsTmp!诊疗项目ID & ",") > 0 Then
                rsTmp!RISItem = 1
            End If
            rsTmp.MoveNext
        Next
        rsTmp.Filter = "RISItem=1"
        If rsTmp.RecordCount = 1 Then
            lng医嘱ID = rsTmp!ID
        Else
            HaveItemToRis = True
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetAdviceReportTip(ByVal lngRow As Long) As String
'功能：获取鼠标悬浮提示字
    Dim strTmp As String
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_RIS报告ID)) <> 0 Then
            strTmp = "(RIS报告)"
        ElseIf Val(.TextMatrix(lngRow, COL_报告ID)) <> 0 Then
            strTmp = "(HIS报告)"
        ElseIf .TextMatrix(lngRow, COL_检查报告ID) <> "" Then
            strTmp = "(专业版PACS报告)"
        ElseIf Val(.TextMatrix(lngRow, COL_LIS报告ID)) <> 0 Then
            strTmp = "(三方LIS报告)"
        Else
            If Val(.TextMatrix(lngRow, COL_RIS预约ID)) <> 0 Then
                If Val(.TextMatrix(lngRow, COL_RIS预约状态)) = 0 Then
                    strTmp = "已经预约"
                Else
                    strTmp = "已经预约调整"
                End If
            End If
        End If
        If strTmp <> "" And Val(.TextMatrix(lngRow, COL_RIS预约ID)) = 0 Then
            If Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 0 Then
                strTmp = "报告未阅" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 1 Then
                strTmp = "报告已阅" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_查阅状态)) = 2 Then
                strTmp = "报告部分已阅" & strTmp
            End If
        End If
    End With
    GetAdviceReportTip = strTmp
End Function

Private Sub FuncApplyCustom(ByVal intType As Long, ByVal lng文件ID As Long)
'功能：自定义申请单
'参数：intType=0 新增，=1修改，=2查看
    Dim lng申请序号 As Long
    Dim datTurn As Date
    Dim lngRow As Long
    Dim lng开嘱科室ID As Long
    Dim lngNo As Long
    Dim objApplyCustom As New frmApplyCustom
    
    If intType <> 2 Then
        If mint场合 <> 2 Then If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
        '修改时检查是否审核
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_审核状态)) = 2 Then
                MsgBox "申请单已经审核，不允许再修改。", vbInformation, "申请单"
                intType = 2
            End If
        End If
        If intType = 0 Then
            If Not FuncPathAdd() Then Exit Sub
        End If
    End If
    
    If intType <> 0 Then
         lng申请序号 = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_申请序号))
         lngRow = vsAdvice.Row
    End If
    
    lng开嘱科室ID = Get开嘱科室ID(UserInfo.ID, mlng界面科室ID, mlng科室ID, 2)
    If objApplyCustom.ShowMe(mfrmParent, 0, intType, mlng病人ID, mlng主页ID, IIF(mlng病人性质 = 1, 1, 0), lng文件ID, lng申请序号, mlng科室ID, lng开嘱科室ID, mlng病区ID, mrsDefine, mintPState, datTurn, 0, mclsMipModule, mlng前提ID, mbyt婴儿, mint险类) Then
        If mlng路径状态 = 1 And Not gobjPath Is Nothing And lng申请序号 <> 0 Then
            lngNo = lng申请序号
            If lngNo <> 0 Then Call FuncPathSet(lngNo)
        End If
        '刷新医嘱
        Call RefreshData
        '选择最后一行医嘱
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_医嘱内容
    End If
End Sub

Private Sub FuncAdviceRISModi()
'功能：调整RIS预约
    Dim lng医嘱ID As Long
    Dim lng预约ID As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        lng预约ID = Val(.TextMatrix(.Row, COL_RIS预约ID))
    End With
    
    strSQL = "select 1 from 病人医嘱发送 a where a.医嘱id=[1] and nvl(a.执行状态,0) in (0,3) and nvl(a.执行过程,0)<=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        If HaveRIS(False) Then
            Call gobjRis.HISReSchedule(lng医嘱ID, lng预约ID)
        End If
    Else
        MsgBox "该项目已经执行，不允许再做调整。", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function MakeBillCharge(ByVal lng医嘱ID As Long) As Long
'功能：检查医嘱回退时自动产生销帐申请药品费用
'参数：是否是批量回退发送
'返回：0-继续回退操作，1-终止回退操作
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim str医嘱IDs As String
    Dim i As Long
    Dim strNO As String
    Dim datCur As Date
    Dim blnTran As Boolean
    Dim arrSQL As Variant
    Dim strMsg As String
    
    On Error GoTo errH
    
    '获取整组医嘱ID拼串
    strSQL = "Select a.id,b.no,a.相关ID,a.医嘱内容 From 病人医嘱记录 A,病人医嘱发送 B Where a.id=b.医嘱id and (a.Id = [1] Or a.相关id = [1]) and b.记录性质=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    For i = 1 To rsTmp.RecordCount
        str医嘱IDs = str医嘱IDs & "," & rsTmp!ID
        strNO = rsTmp!NO
        If IsNull(rsTmp!相关ID) Then
            strMsg = rsTmp!医嘱内容 & ""
        End If
        rsTmp.MoveNext
    Next
    str医嘱IDs = Mid(str医嘱IDs, 2)
    If str医嘱IDs = "" Then Exit Function
    
    '检查是否存在未审核的销帐申请
    strSQL = "Select 1 From 病人医嘱记录 A, 病人医嘱发送 B, 住院费用记录 C, 病人费用销帐 D" & vbNewLine & _
        " Where (a.Id = [1] Or a.相关id = [1]) And a.Id = b.医嘱id And b.医嘱id = c.医嘱序号 And c.Id = d.费用id And" & vbNewLine & _
        " c.记录状态 In (0, 1, 3) And d.状态 = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If Not rsTmp.EOF Then
        strMsg = "医嘱""" & strMsg & """项目存在未审核的销帐申请，请取消或审核销帐申请后再回退。"
        MsgBox strMsg, vbInformation, gstrSysName
        MakeBillCharge = 1
        Exit Function
    End If
    
    '判断是否可以销帐
    strSQL = "Select Min(ID) As 费用id, Max(a.病人病区id) As 申请科室id,a.医嘱序号,a.收费细目id, Sum(a.数次) As 数次" & vbNewLine & _
        "From 住院费用记录 A" & vbNewLine & _
        "Where a.No = [1] And a.记录性质 = 2 And a.收费类别 In ('5', '6')" & vbNewLine & _
        "And instr(','||[2]||',', ','||a.医嘱序号||',')>0" & vbNewLine & _
        "and nvl(a.执行状态,0)<>0" & vbNewLine & _
        "Group By a.收费细目id,a.医嘱序号" & vbNewLine & _
        "having Sum(a.数次)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, str医嘱IDs)
    
    If Not rsTmp.EOF Then
        strMsg = "医嘱""" & strMsg & """项目关联的药品已发药，禁止回退。" & vbCrLf & _
            "是否自动销帐申请药品费用？是则产生销帐申请，否则结束当前操作。"
            
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        
            arrSQL = Array()
            datCur = zlDatabase.Currentdate
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_病人费用销帐_insert(" & rsTmp!费用ID & "," & rsTmp!收费细目ID & "," & rsTmp!申请科室id & "," & rsTmp!数次 & ",'" & UserInfo.姓名 & "'," & _
                    "To_Date('" & Format(datCur, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1,0,'回退检查医嘱自动产生')"
                rsTmp.MoveNext
            Next
            gcnOracle.BeginTrans: blnTran = True
            For i = 0 To UBound(arrSQL)
                Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            gcnOracle.CommitTrans: blnTran = False
            
        End If
        MakeBillCharge = 1
    End If
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncCriticalAdvice(ByVal strPar As String, ByVal blnCheck As Boolean)
'功能：设置（关联/取消）危值医嘱关联
'参数：strPar-格式：危急值ID,医嘱ID(主医嘱ID)
'      blnCheck-true 取消关系，false 设置关系
    Dim lng危急值ID As Long
    Dim lng医嘱ID As Long
    Dim lng功能 As Long
    Dim strSQL As String
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim i As Long
    Dim lngOther危急值ID As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    lng功能 = IIF(blnCheck, 2, 1)
    lng危急值ID = Split(strPar, ",")(0)
    lng医嘱ID = Split(strPar, ",")(1)
    strSQL = "Zl_病人危急值医嘱_Update(" & lng功能 & "," & lng危急值ID & "," & lng医嘱ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If blnCheck Then
        '同一条医嘱可关联多个危急值，取消时要进一步判断是否还有关联
        strSQL = "select a.危急值ID,a.医嘱ID from 病人危急值医嘱 a where a.医嘱ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
        If Not rsTmp.EOF Then
            lngOther危急值ID = rsTmp!危急值ID & ""
        End If
    End If
    
    
    If RowIn一并给药(vsAdvice.Row, lngBegin, lngEnd) Then
        For i = lngBegin To lngEnd
            Set vsAdvice.Cell(flexcpPicture, i, col_医嘱内容) = Nothing
            Set vsAdvice.Cell(flexcpPicture, i, col_内容) = Nothing
            If blnCheck Then
                vsAdvice.TextMatrix(i, COL_危急值ID) = lngOther危急值ID
            Else
                vsAdvice.TextMatrix(i, COL_危急值ID) = lng危急值ID
            End If
            Call SetAdviceIcon(i)
        Next
    Else
        '更新界面表格图标
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_医嘱内容) = Nothing
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_内容) = Nothing
        If blnCheck Then
            vsAdvice.TextMatrix(vsAdvice.Row, COL_危急值ID) = lngOther危急值ID
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, COL_危急值ID) = lng危急值ID
        End If
        Call SetAdviceIcon(vsAdvice.Row)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCriticalAdvice(ByRef lng医嘱ID As Long) As ADODB.Recordset
'功能：根据当前选中行的医嘱查询出与之关联的危急值记录
'参数：出参 lng医嘱ID 即当前界面上选中医嘱的主医嘱ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_相关ID)) <> 0 Then
            lng医嘱ID = Val(.TextMatrix(.Row, COL_相关ID))
        Else
            lng医嘱ID = Val(.TextMatrix(.Row, COL_ID))
        End If
    End With
    
    strSQL = "select a.危急值ID,a.医嘱ID from 病人危急值医嘱 a where a.医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    
    Set GetCriticalAdvice = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get确认会诊(lng医嘱ID As Long) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "SELECT A.报到时间 FROM 病人医嘱发送 A where 医嘱ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(lng医嘱ID))
    If Not rsTmp.EOF Then
        Get确认会诊 = IIF(rsTmp!报到时间 & "" = "", False, True)
    Else
        Get确认会诊 = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Execute确认会诊(blnCancel As Boolean) As Boolean
'功能：是否确认医生参加会诊
    Dim strSQL As String, lng部门ID As Long
    
    If mlng病人ID = 0 Then Exit Function
    If MsgBox("是否" & IIF(blnCancel, "取消确认", "确认") & "医生参加了的会诊吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0 Then
        strSQL = "Zl_病人医嘱发送_会诊处理(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",0," & IIF(blnCancel = True, "3", "2") & ",'" & UserInfo.姓名 & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '要更新执行状态'可能要更新执行状态
    End If
    Execute确认会诊 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetCriticalData()
'功能：获取危急值记录
    Dim strSQL As String
    On Error GoTo errH
    If mbln危急值 Then
        strSQL = "select a.id,a.危急值描述 from 病人危急值记录 a where a.病人ID=[1] and a.主页ID=[2] order by a.报告时间 desc"
        Set mrs危急值 = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mlng病人ID, mlng主页ID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncViewLisRpt()
'功能：浏览检验报告
'说明：分两种模式，先判断本次就诊是否有PDF报告
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If mblnMoved Then
        strSQL = "select 1 from H病人医嘱记录 a,H病人医嘱报告 b,H医嘱报告内容 c where a.id=b.医嘱id and b.报告id=c.id and c.类型  in (0,2) and a.病人id=[1] and a.主页id=[2]"
    Else
        strSQL = "select 1 from 病人医嘱记录 a,病人医嘱报告 b,医嘱报告内容 c where a.id=b.医嘱id and b.报告id=c.id and c.类型  in (0,2) and a.病人id=[1] and a.主页id=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    
    If Not rsTmp.EOF Then
        '两个页签显示
        Call frmLisALL.ShowMe(mfrmParent, mlng病人ID, mlng主页ID, mlng科室ID, mlng病区ID, p住院医嘱下达, mMainPrivs)
    Else
        '以前的老模式
        Call InitObjLis(p住院医生站)
        If Not gobjLIS Is Nothing Then
            gobjLIS.PatientSampleBrowse mfrmParent, mlng病人ID, mMainPrivs, mlng科室ID, mlng病区ID, 2, mlng主页ID
        Else
            frmLisView.ShowMe mlng病人ID, p住院医嘱下达, mfrmParent
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FuncDrugRefcom()
'功能：弹出填写拒绝审核理由窗口调用合理用药部件接口
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strAdviceIDs As String
    Dim strErr As String
    
    On Error GoTo errH
    
    strSQL = "select 1 from 病人医嘱记录 a where a.病人id=[1] and a.主页id=[2] and a.医嘱状态=1 and a.诊疗类别 in ('5','6') and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    If Not rsTmp.EOF Then
        '有新开的药品医嘱
        Call gobjPass.ZLPharmReviewResultIn(mfrmParent, mlng病人ID, mlng主页ID, strErr)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetFormOperation() As String
'功能：获取窗体操作选择，该接口会在窗体卸载前调用，新版护士站 病人事务窗口
'返回：记录当前界面中控件选择状态
'说明：过滤栏当前医嘱行选择，XML结构方式
    

    'Private Type FilterCond
    '    婴儿 As Integer
    '    重整 As Boolean
    '    科内 As Boolean
    '    未记帐 As Boolean
    '    报告 As Integer     '0-全部，1－检查，2－检验，3－其他
    '    未出报告 As Boolean
    '    已出报告 As Boolean
    '    显示模式 As Integer '0-简洁，1－详细
    '    医嘱显示 As Integer '0-在用医嘱，1－所有医嘱
    '    过滤模式 As Integer '0-长嘱临嘱，1－长嘱，2－临嘱，3－报告
    '    开始时间 As Date
    '    结束时间 As Date
    '    是报告医嘱 As Boolean
    '    非报告医嘱 As Boolean
    '    未到终止时间 As Boolean '是否显示未到(执行终止时间)的医嘱
    'End Type
    'mvarCond
    'cboTime.ListIndex
    'lng医嘱ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    
    Dim strXML As String

    strXML = "<root>"
    strXML = strXML & "<ye>" & mvarCond.婴儿 & "</ye>"     '婴儿
    strXML = strXML & "<cz>" & IIF(mvarCond.重整, 1, 0) & "</cz>" '重整，Boolean 0/1
    strXML = strXML & "<kn>" & IIF(mvarCond.科内, 1, 0) & "</kn>"  '科内，Boolean 0/1
    strXML = strXML & "<wjz>" & IIF(mvarCond.未记帐, 1, 0) & "</wjz>"  '未记帐， Boolean 0/1
    strXML = strXML & "<bg>" & mvarCond.报告 & "</bg>"  '报告
    strXML = strXML & "<wcbg>" & IIF(mvarCond.未出报告, 1, 0) & "</wcbg>"  '未出报告 Boolean
    strXML = strXML & "<ycbg>" & IIF(mvarCond.已出报告, 1, 0) & "</ycbg>"  '已出报告 Boolean
    strXML = strXML & "<xsms>" & mvarCond.显示模式 & "</xsms>"  '显示模式
    strXML = strXML & "<yzxs>" & mvarCond.医嘱显示 & "</yzxs>"  '医嘱显示
    strXML = strXML & "<glms>" & mvarCond.过滤模式 & "</glms>"  '过滤模式
    strXML = strXML & "<kssj>" & Format(mvarCond.开始时间, "yyyy-MM-dd HH:mm:ss") & "</kssj>"   '    开始时间 As Date
    strXML = strXML & "<jssj>" & Format(mvarCond.结束时间, "yyyy-MM-dd HH:mm:ss") & "</jssj>"   '     结束时间 As Date
    strXML = strXML & "<sbgyz>" & IIF(mvarCond.是报告医嘱, 1, 0) & "</sbgyz>" '    是报告医嘱 As Boolean
    strXML = strXML & "<fbgyz>" & IIF(mvarCond.非报告医嘱, 1, 0) & "</fbgyz>" '    非报告医嘱 As Boolean
    strXML = strXML & "<wdzzsj>" & IIF(mvarCond.未到终止时间, 1, 0) & "</wdzzsj>" '    未到终止时间 As Boolean '是否显示未到(执行终止时间)的医嘱
    strXML = strXML & "<cbotime>" & cboTime.ListIndex & "</cbotime>" '时间范围下拉框索引值
    strXML = strXML & "<yzid>" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & "</yzid>" '界面选择的医嘱ID
    strXML = strXML & "</root>"
    

    GetFormOperation = strXML
End Function

Public Function RestoreFormOperation(ByVal strValue As String)
'功能：恢复窗体操作选择
'参数：strValue 前界面中控件选择状态
'Public Sub LocatedAdviceRow(ByVal lng医嘱ID As Long)
    Dim objXML As New zl9ComLib.clsXML
    Dim strTmp As String
    
    On Error Resume Next
    
    Call objXML.OpenXMLDocument(strValue)
    
    Call objXML.GetSingleNodeValue("ye", strTmp) '婴儿
    mvarCond.婴儿 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("cz", strTmp) '重整
    mvarCond.重整 = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("kn", strTmp) '科内
    mvarCond.科内 = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("wjz", strTmp) '未记帐
    mvarCond.未记帐 = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("bg", strTmp) '报告
    mvarCond.报告 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("wcbg", strTmp) '未出报告
    mvarCond.未出报告 = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("ycbg", strTmp) '已出报告
    mvarCond.已出报告 = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("xsms", strTmp) '显示模式
    mvarCond.显示模式 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("yzxs", strTmp) '医嘱显示
    mvarCond.医嘱显示 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("glms", strTmp) '过滤模式
    mvarCond.过滤模式 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("kssj", strTmp) '开始时间
    mvarCond.开始时间 = CDate(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("jssj", strTmp) '结束时间
    mvarCond.结束时间 = CDate(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("sbgyz", strTmp) '是报告医嘱
    mvarCond.是报告医嘱 = 1 = Val(strTmp): strTmp = ""
    
    
    Call objXML.GetSingleNodeValue("fbgyz", strTmp) '非报告医嘱
    mvarCond.非报告医嘱 = 1 = Val(strTmp): strTmp = ""
    
    
    Call objXML.GetSingleNodeValue("wdzzsj", strTmp) '未到终止时间
    mvarCond.未到终止时间 = 1 = Val(strTmp): strTmp = ""
        
    
    Call objXML.GetSingleNodeValue("cbotime", strTmp) '时间范围下拉框索引值
    cboTime.ListIndex = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("yzid", strTmp) '按开嘱时间
    mvarCond.医嘱ID = Val(strTmp): strTmp = ""
    
End Function

Private Sub Set标本状态()
'功能：对检验医嘱设置标本状态列，结果多LIS部件中返回
    Dim i As Long, str医嘱IDs As String, strMsg As String
    Dim rsAdvice As ADODB.Recordset
    Dim strIDAndRow As String, strTmp As String
    Dim lngRow As Long
    
    On Error GoTo errH
    
    If mvarCond.过滤模式 <> 3 Then Exit Sub
    Call InitObjLis(p住院医生站)
    If gobjLIS Is Nothing Then Exit Sub
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "6" And Val(.TextMatrix(i, COL_相关ID)) = 0 And Val(.TextMatrix(i, COL_医嘱状态)) = 8 Then
                str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                strIDAndRow = strIDAndRow & "," & Val(.TextMatrix(i, COL_ID)) & ";" & i & "<Tab>"
            End If
        Next
        If str医嘱IDs <> "" Then
            Set rsAdvice = gobjLIS.GetSampleType(Mid(str医嘱IDs, 2), strMsg)
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
            If Not rsAdvice Is Nothing Then
                rsAdvice.Filter = 0
                For i = 1 To rsAdvice.RecordCount
                    If InStr(strIDAndRow, "," & rsAdvice!医嘱ID & ";") > 0 Then
                        strTmp = Split(strIDAndRow, "," & rsAdvice!医嘱ID & ";")(1)
                        lngRow = Val(Split(strTmp, "<Tab>")(0))
                        .TextMatrix(lngRow, COL_标本状态) = rsAdvice!医嘱状态 & ""
                    End If
                    rsAdvice.MoveNext
                Next
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncViewPacsRpt()
'功能：浏览检检查报告
'说明：未处理阅读标记
    Dim blnAutoRead As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng医嘱ID As Long
    
    On Error GoTo errH
    Call CreateObjectPacs(mobjPublicPACS)
    If Not mobjPublicPACS Is Nothing Then
        
        strSQL = "select max(b.id) as 医嘱ID  from 病人医嘱报告 a,病人医嘱记录 b " & _
                " Where a.检查报告ID Is Not Null And a.医嘱ID = b.ID And b.病人id=[1] and b.主页id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        lng医嘱ID = Val(rsTmp!医嘱ID & "")
        
        Call mobjPublicPACS.zlDocShowReport(lng医嘱ID, , blnAutoRead, mfrmParent)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
