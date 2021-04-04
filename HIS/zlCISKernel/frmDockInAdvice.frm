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
            Name            =   "����"
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
            ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
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
            Name            =   "����"
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
         ToolTipText     =   "���ͣ��ʱ,�������������Զ���ʾ"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
            Key             =   "���δ�ӡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockInAdvice.frx":144CC
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDockInAdvice.frx":14A66
            Key             =   "ͣ������"
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
Public Event Activate() '���Ѽ���ʱ
Public Event RequestRefresh(ByVal RefreshNotify As Boolean) 'Ҫ��������ˢ��
Public Event StatusTextUpdate(ByVal Text As String) 'Ҫ�����������״̬������
Public Event ViewEPRReport(ByVal ����ID As Long, ByVal CanPrint As Boolean) 'Ҫ��鿴����
Public Event PrintEPRReport(ByVal ����ID As Long, ByVal Preview As Boolean) 'Ҫ���ӡ����
Public Event ViewPACSImage(ByVal ҽ��ID As Long) 'Ҫ����й�Ƭ
Public Event ExecLogNew(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ��� As Boolean) 'ִ������Ǽ�
Public Event ExecLogModi(ByVal ҽ��ID As Long, ByVal ���ͺ� As Long, ByVal ����ID As Long, ByVal ִ��ʱ�� As String, ��� As Boolean) 'ִ������޸�
Public Event EditDiagnose(ParentForm As Object, ByVal ����ID As Long, ByVal ��ҳID As Long, ByVal ����ID As Long, ByVal str���� As String, Succeed As Boolean) '�༭סԺ���
Public Event SetEditState(ByVal blnEditState As Boolean)    '�༭״̬ʱ���ò˵��Ϳ�ת�ƽ���Ĺ���
Public Event DoByAdvice(ByVal lngҽ��ID As Long, ByVal lng���ID As Long, ByVal lngWayID As Long, ByVal strTag As String)   'ҽ����ز�����lngWayID ����ID��Ŀǰֻ֧��  ��ҽ���Ƽ�,strTag ��չ����

Private mint���� As Integer  '���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
Private mMainPrivs As String '���������������е�Ȩ��,ע����ڲ�ģ��Ȩ��
Private mcbsMain As Object
Private WithEvents mfrmParent As Form
Attribute mfrmParent.VB_VarHelpID = -1
Private WithEvents mfrmEdit As frmInAdviceEdit 'ҽ���༭����
Attribute mfrmEdit.VB_VarHelpID = -1
Private WithEvents mfrmEac As frmApplyConsultation    '�������뵥����
Attribute mfrmEac.VB_VarHelpID = -1
Private WithEvents mfrmBilling As Form '���ʹ�����
Attribute mfrmBilling.VB_VarHelpID = -1
Private WithEvents mfrmCompoundMedicine As frmCompoundMedicine  '��Һ��ҩ��¼
Attribute mfrmCompoundMedicine.VB_VarHelpID = -1
Private mobjPublicPACS As Object             'PACSҵ���װ��������

Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mstrBillPrint As String '��ǰ��ӡ�����Ƶ��ݣ������š�NO����¼����

Private mintPState As TYPE_PATI_State '����״̬
Private mintִ��״̬ As Integer 'ҽ��վ��ִ��״̬
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng�������� As Long    '0-��ͨסԺ����,1-�������۲���,2-סԺ���۲���
Private mbytӤ�� As Byte
Private mint���� As Integer

Private mlng����ID As Long      '�����ת�����ˣ���Ϊԭ����ID
Private mlng����ID As Long      '�����ת�����ˣ���Ϊԭ����ID
Private mlngǰ��ID As Long
Private mlng����ҽ��ID As Long
Private mstrǰ��IDs As String
Private mlng�������ID As Long
Private mlngҽ������ID As Long
Private mstr���� As String
Private mstr�Ա� As String
Private mstrסԺ�� As String
Private mstr���� As String
Private mdat���� As Date '������ҳ.ҽ������ʱ��
Private mblnBatch As Boolean '��������ģʽ���̶���������ѡ���
Private mblnDirect As Boolean '�Ƿ�ֱ�ӵ��ù��ܣ�����ʾҽ���嵥������£�
Private mblnInsideTools As Boolean '�ڲ�������ģʽ
Private mblnHaveAuditPriv As Boolean
Private mblnSignVisible As Boolean  'ǩ�����ܰ�ť�ɼ���
Private mblnModalNew As Boolean '�¿������Ƿ�ģ̬

Private mvInDate As Date '��Ժ����
Private mblnMoved As Boolean
Private mstrӤ�� As String
Private mstrסԺҽ�� As String '���˵�סԺҽʦ
Private mstr���λ�ʿ As String '���˵����λ�ʿ
Private mint����״̬ As Integer
Private mlng·��״̬ As Long    '-1-δ���룬0-�����ϵ���������1-ִ���У�2-����������3-�������
Private mrsDefine As ADODB.Recordset    'ҽ�����ݶ���
Private mobjVBA As Object
Private mobjScript As clsScript
Private mlngӤ������ID As Long
Private mlngӤ������ID As Long
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�

Private mblnFirst As Boolean '�Ƿ��״ε���
Private mlngPlugInID As Long '�Զ�ִ�еĲ������ID
Private mrsPlugInBar As ADODB.Recordset '�˵���ʽ
Private mlngPromptRow As Long    '��һ�Σ�������ƶ�ͼ������ʾ����ʾ��Ϣ����

'Pass
Private mobjPassMap As Object  'PASS �������ӳ��
Private mblnPass As Boolean  'PASSȨ��


'ģ�����
Private mbln���� As Boolean
Private mblnƤ����֤ As Boolean
Private mbln��ʿǩ�� As Boolean
Private mblnShowExec As Boolean
Private mblnAutoRead As Boolean
Private mblnAutoReadEnabled As Boolean
Private mblnEditState As Boolean    '��ҩ���α༭״̬
Private mblnNotEvaluete As Boolean  'δ����ʱ�������ҽ��������
Private mlngBaby As Long
Private mblnFirstBaby As Boolean    '��һ�ΰ�������ѡ
Private mlngBabyDept As Long      '��һ��ѡ���Ӥ��ѡ��
Private mintBillPrint As Integer   '0-ѡ��ҽ���嵥��ӡ���Ƶ��ݣ���ӡ���һ�η��͵����Ƶ��ݣ���1-ѡ���ͼ�¼��ӡ���Ƶ���
Private mint���뵥��ӡģʽ As Integer  '1-����ʱ��ӡ��2-�¿�ʱ��ӡ
Private mlngPrintType As Long 'ҽ����ӡģʽ
Private mlngPrintPos As Long    'ҽ����ӡʱ��ת�ƺͳ�Ժҽ����ӡ�ڣ�0-����ҽ�����ϣ�1-��ʱҽ�����ϣ�2-���߶���ӡ��
Private mclsMipModule As zl9ComLib.clsMipModule
Attribute mclsMipModule.VB_VarHelpID = -1
Private mbln���� As Boolean '�Ƿ��Ǳ���ҳǩ����֤��ҽ�����л�ʱ��ˢ�²˵�
Private mstr�����Ժ��� As String
Private mstr�Զ������뵥IDs As String 'ID1,����1|ID2,����2������
Private mbln��������ִ�� As Boolean
Private mobjFrmBlood As Object 'ѪҺִ�д���
Private mobjFrmBloodList As Object 'ѪҺ��ϸ����
Private mrsΣ��ֵ As ADODB.Recordset '��ǰ���˵�Σ��ֵ��Ϣ
Private mblnΣ��ֵ As Boolean '��Σ��ֵ��Ȩ��
Private mlngΣ��ֵID As Long '��ǰ�����Σ��ֵ��¼ID
Private mblnȷ�ϻ��� As Boolean  '��ǰ�����¼ȷ��ҽ���Ƿ񵽴�
Private mblnҽ����λ��� As Boolean  'ҽ�����Ĭ�϶�λ�����һ��

'����ҽ����������
Private Enum CMD_FILTER
    ID_����ҽ�� = 1
    ID_����ҽ�� = 2
    ID_Ӥ�� = 3
    ID_���� = 4
    ID_δ���� = 5
    ID_���� = 6
    ID_��� = 7
    ID_��ϸ = 8
    ID_ȫ�� = 9
    ID_��� = 10
    ID_���� = 11
    ID_���� = 12
    ID_ʱ�� = 13
    ID_ʱ���ǩ = 14
    ID_�Ǳ���ҽ�� = 15
    ID_�Ǳ���ҽ�� = 16
    ID_δ����ֹʱ�� = 17
    ID_ҽ����ɫʾ�� = 18
    ID_δ������ = 19
    ID_�ѳ����� = 20
End Enum

Private Enum CMD_EXEC
    ID_��ʾִ�� = 1
    ID_���ִ�� = 2
    ID_ȡ����� = 3
    ID_ִ�м�¼ = 4
    ID_ִ�е��� = 5
    ID_ִ��ɾ�� = 6
    ID_�˶� = 7
    ID_ȡ���˶� = 8
End Enum

Private Type FilterCond
    Ӥ�� As Integer
    ���� As Boolean
    ���� As Boolean
    δ���� As Boolean
    ���� As Integer     '0-ȫ����1����飬2�����飬3������
    δ������ As Boolean
    �ѳ����� As Boolean
    ��ʾģʽ As Integer '0-��࣬1����ϸ
    ҽ����ʾ As Integer '0-����ҽ����1������ҽ��
    ����ģʽ As Integer '0-����������1��������2��������3������
    ��ʼʱ�� As Date
    ����ʱ�� As Date
    �Ǳ���ҽ�� As Boolean
    �Ǳ���ҽ�� As Boolean
    δ����ֹʱ�� As Boolean '�Ƿ���ʾδ��(ִ����ֹʱ��)��ҽ��
    ҽ��ID As Long  '��ǰҽ�������ѡ���е�ҽ��ID
End Type
Private mvarCond As FilterCond
Private mblnHideFilter As Boolean
Private mintPreTime As Integer

'��ŵ�ǰҽ���ɻ����б�
Private Type TYPE_AdviceRoll
    ���ͺ� As Long
    �������� As Integer
    ����ʱ�� As Date
    ������Ա As String
    �������� As String
End Type
Private marrRollList() As TYPE_AdviceRoll
Private mstr����IDs As String '����Ա���������һ���
Private mblnAppend As Boolean '�Ƿ���ʾ������Ϣ
Private mlngFontSize As Long  '�����С

Private Enum COLҽ���嵥
    '�̶���
    COL_F��־ = 0
    COL_F���� = 1
    '������
    COL_ID = 2
    COL_���ID = COL_ID + 1
    COL_��� = COL_ID + 2
    COL_Ӥ��ID = COL_ID + 3
    COL_ҽ��״̬ = COL_ID + 4   'flexcpData�д洢���״̬
    COL_������� = COL_ID + 5
    COL_�������� = COL_ID + 6
    COL_������� = COL_ID + 7
    COL_��־ = COL_ID + 8
    '�ɼ���
    COL_��ʾ = COL_ID + 9 'Pass
    COL_��Ч = COL_ID + 10
    COL_��ʼʱ�� = COL_ID + 11
    COL_�� = COL_ID + 12
    col_ҽ������ = COL_ID + 13
    col_���� = COL_ID + 14
    COL_Ƥ�� = COL_ID + 15
    COL_���� = COL_ID + 16
    COL_���� = COL_ID + 17
    COL_���� = COL_ID + 18
    COL_Ƶ�� = COL_ID + 19
    COL_�÷� = COL_ID + 20
    COL_ҽ������ = COL_ID + 21
    COL_ִ��ʱ�� = COL_ID + 22
    COL_��ֹʱ�� = COL_ID + 23
    COL_ִ�п��� = COL_ID + 24
    COL_ִ������ = COL_ID + 25
    COL_�ϴ�ִ�� = COL_ID + 26
    COL_״̬ = COL_ID + 27
    COL_����ҽ�� = COL_ID + 28
    COL_����ʱ�� = COL_ID + 29
    COL_У�Ի�ʿ = COL_ID + 30
    COL_У��ʱ�� = COL_ID + 31
    COL_ͣ��ҽ�� = COL_ID + 32
    COL_ͣ��ʱ�� = COL_ID + 33
    COL_ͣ����ʿ = COL_ID + 34
    COL_ȷ��ͣ��ʱ�� = COL_ID + 35
    COL_����ҩ�� = COL_ID + 36
    COL_����״̬ = COL_ID + 37
    COL_�걾״̬ = COL_ID + 38
    
    '������
    COL_������ĿID = COL_ID + 39
    COL_�Թܱ��� = COL_������ĿID + 1
    COL_ִ�б�� = COL_������ĿID + 2
    COL_���δ�ӡ = COL_������ĿID + 3
    COL_ǰ��ID = COL_������ĿID + 4
    COL_ǩ���� = COL_������ĿID + 5
    COL_�ļ�ID = COL_������ĿID + 6
    COL_������ = COL_������ĿID + 7 '0-�ޱ��棬1-�б��沢���༭��ʽ��ӡ��2-�б��沢�������ʽ��ӡ��
    COL_����ID = COL_������ĿID + 8
    COL_�շ�ϸĿID = COL_������ĿID + 9
    COL_������λ = COL_������ĿID + 10
    COL_��������ID = COL_������ĿID + 11
    COL_���״̬ = COL_������ĿID + 12
    COL_������� = COL_������ĿID + 13
    COL_��˱�� = COL_������ĿID + 14
    COL_��ΣҩƷ = COL_������ĿID + 15
    COL_�걾��λ = COL_������ĿID + 16   'PASS  ҩƷ����
    COL_��ҩĿ�� = COL_������ĿID + 17
    COL_��鱨��ID = COL_������ĿID + 18
    COL_�������״̬ = COL_������ĿID + 19
    COL_��������� = COL_������ĿID + 20
    COL_RISԤԼID = COL_������ĿID + 21
    COL_RIS����ID = COL_������ĿID + 22
    COL_LIS����ID = COL_������ĿID + 23
    COL_RISԤԼ״̬ = COL_������ĿID + 24
    col_������Ŀ���� = COL_������ĿID + 25
    COL_��鷽�� = COL_������ĿID + 26 '�����Ǳ�Ѫҽ��������Ѫҽ��
    COL_Σ��ֵID = COL_������ĿID + 27 'ҽ���غ�Σ��ֵ����
    COL_�׵��� = COL_������ĿID + 28 'ҩƷ���׵���
End Enum

Private COLPrice As New Collection
Private COLSend As New Collection
Private COLSign As New Collection
Private COLExec As New Collection

Public Sub zlDefCommandBars(ByVal frmParent As Object, ByVal cbsMain As Object, ByVal int���� As Integer, _
                            ByVal blnInsideTools As Boolean, ByRef objSquareCard As Object, Optional ByVal blnModalNew As Boolean)
    Dim lngTmp As Long
    
    mint���� = int����

    mblnInsideTools = blnInsideTools
    Set mfrmParent = frmParent
        mblnModalNew = blnModalNew

    If Not cbsMain Is Nothing Then

        '��һ�ε���ʱ��������(���ܷ���Form_Load�¼��У���ΪGetFormʱ�������¼�ʱ��û�д�mint����)
        If Not mblnFirst Then
            mblnFirst = True

            Set mcbsMain = cbsMain
            Set cbsMain.Icons = zlCommFun.GetPubIcons
            Set gobjSquareCard = objSquareCard

            If mint���� = 0 Then 'ҽ��վ����
                lngTmp = pסԺҽ���´�
            ElseIf mint���� = 1 Then '��ʿվ����
                lngTmp = pסԺҽ������
            ElseIf mint���� = 2 Then 'ҽ��վ����
                lngTmp = pסԺҽ���´�
            End If
        
            '��ҳ�������ʼ��
            If gobjPlugIn Is Nothing Then
                On Error Resume Next
                Set gobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
                err.Clear: On Error GoTo 0
            End If
            If Not gobjPlugIn Is Nothing Then
                On Error Resume Next
                Call gobjPlugIn.Initialize(gcnOracle, glngSys, lngTmp, mint����)
                Call zlPlugInErrH(err, "Initialize")
                err.Clear: On Error GoTo 0
                Call GetPlugInBar(lngTmp, mint����, mrsPlugInBar)
            End If

            'PASS�ӿڳ�ʼ��
            If gobjPass Is Nothing Then
                Set gobjPass = DynamicCreate("zlPassInterface.clsPass", "������ҩ���", True)
                If Not gobjPass Is Nothing Then
                    Call gobjPass.zlPassInit(gcnOracle, glngSys, PM_סԺҽ���嵥)
                    If gobjPass.PassType = 0 Then   'ϵͳ����δ���ú�����ҩ���
                        Set gobjPass = Nothing
                    Else
                        mblnPass = True
                    End If
                End If
            End If
           
        End If
        
        '���ܷ���Form_Load�¼��У���ΪGetFormʱ�������¼�ʱ��û�г�ʼ��������ҩ���)
        Call zlPASSMap
        If mblnPass Then
           'Pass
            Call gobjPass.zlPassAdviceColHidden(mobjPassMap) '��ʾ��
        End If

        If mint���� = 0 Then    'ҽ��վ����
            Call DefCommandsInDoctor(cbsMain)
        ElseIf mint���� = 1 Then    '��ʿվ����
            Call DefCommandsInNurse(cbsMain)
        ElseIf mint���� = 2 Then    'ҽ��վ����
            Call DefCommandsTechnic(cbsMain)
        End If

        If mint���� <> 1 Then   '����ʿվ����ʾ��ҩ��¼(Form_loadʱ���ϻ�û�д���)
            For lngTmp = 0 To tbcAppend.ItemCount - 1
                If tbcAppend(lngTmp).Tag = "��ҩ" Then
                    Call tbcAppend.RemoveItem(lngTmp)
                    Exit For
                End If
            Next
        End If

        '��ҳ����������
        Call DefCommandPlugIn(cbsMain, mrsPlugInBar)
        If mint���� = 1 Then Call SetSendCommandBar '����ǻ�ʿվ���ã�������ӷ��Ͱ�ť
    End If
End Sub

Private Sub DefCommandPlugIn(ByRef cbsMain As Object, ByRef rsBar As ADODB.Recordset)
'���ܣ���Ҳ����˵����롣
'˵�����жϹؼ���  Auto  InTool �����˵���ʽ
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
    '������ť
    rsBar.Filter = "IsInTool=1 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        If Not objMenu Is Nothing Then
            With objMenu.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������)
                        objControl.IconId = rsBar!ͼ��ID
                        objControl.Parameter = rsBar!������
                        objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '������ť�����ֻ��һ����ť��Ҳ����������ť
    rsBar.Filter = "IsInTool=0 and BarType=1"
    If Not rsBar.EOF Then
        rsBar.Sort = "���"
        If Not objMenu Is Nothing Then
            Set objPopup = objMenu.CommandBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����", , False)
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                For i = 1 To rsBar.RecordCount
                    Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
                    objControl.Style = xtpButtonIconAndCaption
                    If Val(rsBar!IsGroup) = 1 Then
                        objControl.BeginGroup = True
                    End If
                    rsBar.MoveNext
                Next
            End With
        End If
    End If
    
    '��������ť
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
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������, lngTmp + 1)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
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
        rsBar.Sort = "���"
        Set objPopup = objBar.Controls.Add(xtpControlPopup, conMenu_Tool_PlugIn, "��չ����", lngTmp + 1, False)
            objPopup.ID = conMenu_Tool_PlugIn
            objPopup.IconId = conMenu_Tool_PlugIn
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        lngTmp = objPopup.Index
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���, lngTmp + 1)
                objControl.IconId = rsBar!ͼ��ID
                objControl.Parameter = rsBar!������
                If Val(rsBar!IsGroup) = 1 Then objControl.BeginGroup = True
                lngTmp = objPopup.Index
                rsBar.MoveNext
            Next
        End With
    End If
    
    '�Զ�ִ�еĹ���
    rsBar.Filter = "IsAuto=1"
    If Not rsBar.EOF Then mlngPlugInID = rsBar!����ID
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
    
    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "ҽ���༭(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "�¿�ҽ��(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "�޸�ҽ��(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "ɾ��ҽ��(&D)"
        End With
        
        intTmp = Val(Mid(gstrInUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",�������:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrInUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",��������:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrInUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��Ѫ����:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrInUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_OperationApply
        intTmp = Val(Mid(gstrInUseApp, 5, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_ConsultationApply
                Get�Զ������뵥 2, mstr�Զ������뵥IDs
        If mstr�Զ������뵥IDs <> "" Then
            For i = 0 To UBound(Split(mstr�Զ������뵥IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(0)
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
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "�´�����"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "�޸�����")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴����")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "ȡ������")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "���ԤԼ")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "ԤԼ(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "����ԤԼ(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "ȡ��ԤԼ(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ҽ������(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ҽ��ֹͣ(&S)")
        If InStr(GetInsidePrivs(pסԺҽ���´�), "�����������") = 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "��������(&G)"): objControl.BeginGroup = True
            objControl.IconId = conMenu_Edit_Send
        Else
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "��������"): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Edit_SendBilling, "סԺ����"    'update�¼��и��ݵ�ǰ�Ƿ����۲����پ�����ʾ������ʻ���סԺ����
                .Add xtpControlButton, conMenu_Edit_SendCharge, "�����շ�"
            End With
        End If
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Untread, "ҽ������(&L)")
        'ҽ������վ�ṩ�鿴������ҩƷ˵����Ĳ˵�
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "�ؼ�ͼ��")
       '2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���������(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "��������(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "������ͼ��ͱ���(&Y)")
                objControl.IconId = 237
        End If

        If gblnѪ��ϵͳ Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_BloodInstant, "��Ѫִ�е�")
            objControl.BeginGroup = True
        End If
    End With
    If Not objMenu Is Nothing Then
        With objMenu.CommandBar.Controls
            If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
            '���������ǰ��,�������
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ���Ƶ���")
            
            '��Ҳ˵�
            Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
        End With
    End If
    
    '����˵�:���������û��,���ڲ鿴�˵�ǰ��
    '-----------------------------------------------------
    '����վ����˵��Զ���ʾ��������Թ���վ��ģ���ͳһ����
    '���⼸�ű�����ҽ������ģ���еģ���Ҫ�ڸ�ģ���е�������
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    End If
    
    If Not objMenu Is Nothing Then
        With objMenu.CommandBar.Controls
            If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
            '���������ǰ��,�������
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ���Ƶ���", 1)
            objPopup.Visible = False '���أ�ֻ�����Ҽ��˵�����
        End With
    End If

    '�鿴�˵�
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '״̬�����
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "������Ϣ(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "�Զ����ع���������(&H)", objControl.Index + 1)
    End With
        
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "����ǩ��(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ҽ��ǩ��(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "��ӡԤԼ��")
                objControl.IconId = 103
        End If
        If gbln����ҩ�����հ������������� Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ҽ��ѡ��(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "���׷�������(&S)"): objControl.BeginGroup = True
    End With

    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Call AddToolBarInDoctor
    
    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�¿�ҽ��
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�ҽ��
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete 'ɾ��ҽ��
        .Add FCONTROL, vbKeyS, conMenu_Edit_Stop 'ֹͣҽ��
        .Add FCONTROL, vbKeyG, conMenu_Edit_SendBilling 'ҽ������
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread 'ҽ������
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend '���ı���
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '��Ƭ����
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '������ͼ��ͱ���
       
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '�Զ����ع���������
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '���������
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '��������
        .Add 0, vbKeyF11, conMenu_Tool_Option 'ҽ��ѡ��
    End With

    '���ò���������
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
    
    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "ҽ���༭(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "�¿�ҽ��(&A)"
            .Add xtpControlButton, conMenu_Edit_Audit, "�������(&T)"
            .Add xtpControlButton, conMenu_Edit_Modify, "�޸�ҽ��(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "ɾ��ҽ��(&D)"
        End With
        
        intTmp = Val(Mid(gstrInUseApp, 1, 1))
        If intTmp = 1 Then strTmp = strTmp & ",�������:" & conMenu_Edit_PacsApply
        intTmp = Val(Mid(gstrInUseApp, 2, 1))
        If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",��������:" & conMenu_Edit_LISApply
        intTmp = Val(Mid(gstrInUseApp, 3, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��Ѫ����:" & conMenu_Edit_BloodApply
        intTmp = Val(Mid(gstrInUseApp, 4, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_OperationApply
        intTmp = Val(Mid(gstrInUseApp, 5, 1))
        If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_ConsultationApply
        Get�Զ������뵥 2, mstr�Զ������뵥IDs
        If mstr�Զ������뵥IDs <> "" Then
            For i = 0 To UBound(Split(mstr�Զ������뵥IDs, "|"))
                strTmp = strTmp & "," & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(0)
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
                        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Apply, "�´�����"): objPopup.BeginGroup = True
                        objPopup.IconId = conMenu_Manage_Request
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    Else
                        Set objControl = objPopup.CommandBar.Controls.Add(xtpControlButton, lngID, strName)
                    End If
                    If UBound(Split(strTmp, ":")) = 2 Then objControl.Parameter = Val(Split(strTmp, ":")(2))
                Next
            End If
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "�޸�����")
            objControl.IconId = 3002
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴����")
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "ȡ������")
        End If
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "���ԤԼ")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "ԤԼ(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "����ԤԼ(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "ȡ��ԤԼ(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "���δ��(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Sort, "����˳��(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ҽ������(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ҽ��ֹͣ(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_StopAudit, "ͣ�����(&W)"): objControl.IconId = conMenu_Edit_Audit
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "ҽ����ͣ(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ҽ������(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ҽ������(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "���δ�ӡ")
        If gblnѪ��ϵͳ Then Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReaction, "��Ѫ��Ӧ"): objControl.BeginGroup = True: objControl.IconId = 4113
        If mblnΣ��ֵ Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_CriticalAdvice, "Σ��ֵҽ��")
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "ҽ������ִ��(&W)"): objControl.BeginGroup = True: objControl.IconId = 3587
        If InStr(GetInsidePrivs(pסԺҽ���´�), "�����������") = 0 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "��������(&G)"): objControl.BeginGroup = True
            objControl.IconId = conMenu_Edit_Send
        Else
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "��������"): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                .Add xtpControlButton, conMenu_Edit_SendBilling, "סԺ����"    'update�¼��и��ݵ�ǰ�Ƿ����۲����پ�����ʾ������ʻ���סԺ����
                .Add xtpControlButton, conMenu_Edit_SendCharge, "�����շ�"
            End With
        End If
        
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Untread, "ҽ������(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "��������(&C)")
        objControl.IconId = conMenu_Edit_ChargeOff
                
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ����(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "�ؼ�ͼ��")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���������(&R)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "��������(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView

        If CreateObjectPacs(mobjPublicPACS) Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewPacs, "������ͼ��ͱ���(&Y)")
                objControl.IconId = 237
        End If

        Set objControl = .Add(xtpControlButton, conMenu_Manage_RecipeAuditView, "�鿴���������")
        objControl.IconId = 3205
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewDrugExplain, "�鿴ҩƷ˵����")
        objControl.IconId = 3205
        If gbln��ϵͳ Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Refcom, "�ܾ��������")
                objControl.IconId = 3205
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ViewRefcom, "�������δͨ����Ϣ")
                objControl.IconId = 3205
        End If
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
        '2012-02-16 by���¶�
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPrint, "������ӡ���鱨��(&J)"): objPopup.BeginGroup = True

        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ���Ƶ���(&1)")
        objPopup.BeginGroup = True
        '��Ҳ˵�
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With
    
    '����˵�:���������û��,���ڲ鿴�˵�ǰ��
    '-----------------------------------------------------
    '����վ����˵��Զ���ʾ��������Թ���վ��ģ���ͳһ����
    '���⼸�ű�����ҽ������ģ���еģ���Ҫ�ڸ�ģ���е�������
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    End If
    With objMenu.CommandBar.Controls
        If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
        '���������ǰ��,�������
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill3, "ҽ����¼��(&3)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill1, "����ҽ����(&2)", 1)
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ���Ƶ���(&1)", 1)
    End With
    
    '�鿴�˵�
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '״̬�����
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "������Ϣ(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "�Զ����ع���������(&H)", objControl.Index + 1)
    End With
        
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "����ǩ��(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ҽ��ǩ��(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "��ӡԤԼ��")
                objControl.IconId = 103
        End If
        If gbln����ҩ�����հ������������� Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ҽ��ѡ��(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "���׷�������(&S)"): objControl.BeginGroup = True
    End With

    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Call AddToolBarInDoctor

    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�¿�ҽ��
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�ҽ��
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete 'ɾ��ҽ��
        .Add FCONTROL, vbKeyS, conMenu_Edit_Stop 'ֹͣҽ��
        .Add FCONTROL, vbKeyG, conMenu_Edit_SendBilling 'ҽ������
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread 'ҽ������
        
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend '���ı���
        .Add FCONTROL, vbKeyW, conMenu_Edit_MarkMap '��Ƭ����
        .Add FCONTROL, vbKeyY, conMenu_Edit_ViewPacs '������ͼ��ͱ���
        
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '�Զ����ع���������
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '���������
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '��������
        .Add 0, vbKeyF11, conMenu_Tool_Option 'ҽ��ѡ��
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
  
    'ҽ���˵�:���ڹ���˵�(���������û��)���ļ��˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ManagePopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_FilePopup)
    Else
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End If
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "ҽ��(&A)", objMenu.Index + 1, False)
    objMenu.ID = conMenu_EditPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewItem, "ҽ���༭(&E)", 1)
            objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Edit_NewItem, "�¿�ҽ��(&A)"
            .Add xtpControlButton, conMenu_Edit_Modify, "�޸�ҽ��(&M)"
            .Add xtpControlButton, conMenu_Edit_Delete, "ɾ��ҽ��(&D)"
        End With
        If HaveRIS Then
            Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_NewRis, "���ԤԼ")
                objPopup.IconId = conMenu_Manage_Request
                objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisSch, "ԤԼ(&A)")
                    objControl.IconId = conMenu_Edit_NewItem
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisModi, "����ԤԼ(&M)")
                    objControl.IconId = conMenu_Edit_Modify
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewRisDel, "ȡ��ԤԼ(&D)")
                    objControl.IconId = conMenu_Edit_Delete
            End With
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Edit_UnUse, "���δ��(&H)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Sort, "����˳��(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "ҽ������(&B)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ҽ��ֹͣ(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "ȷ��ֹͣ(&C)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Pause, "ҽ����ͣ(&P)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ҽ������(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "ҽ��У��(&V)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Price, "�Ƽ۵���(&I)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ClearUp, "ҽ������(&F)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_NoPrint, "���δ�ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ҽ������(&G)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "ҽ������ִ��(&W)"): objControl.BeginGroup = True: objControl.IconId = 3587
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "ҽ�������˶�(&X)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Surplus, "ҩƷ����Ǽ�(&J)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Untread, "ҽ������(&L)"): objPopup.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePrice, "��ҽ������")
            objControl.IconId = conMenu_Edit_Price
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_ChargeOff, "��������(&B)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBack, "���ڷ����ջ�(&N)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Test, "Ƥ�Խ��(&T)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_BatUnPack, "�������(&U)"): objControl.BeginGroup = True: objControl.IconId = 312
        Set objControl = .Add(xtpControlButton, conMenu_Edit_MeetArrive, "ȷ�ϻ���ҽ������(&M)"): objControl.BeginGroup = True: objControl.IconId = 8122
        
        '��ʿվ�ṩ������������ҩƷ˵����Ĳ˵�
        If mblnPass Then
            Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objMenu.CommandBar.Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit)
        End If
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "����(&R)")
        objPopup.BeginGroup = True
        objPopup.IconId = conMenu_Manage_Report
        
        ' 2009-01-15
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���������(&R)"): objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPacsView, "��������(&P)")
            objControl.IconId = conMenu_Manage_ReportLisView
        
        '2012-02-16 by���¶�
        Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportPrint, "������ӡ���鱨��(&J)"): objPopup.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ���Ƶ���(&3)")
        objPopup.BeginGroup = True
        '2017-11-10 ������
        If gblnѪ��ϵͳ Then
            Set objControl = .Add(xtpControlButton, conMenu_Report_BloodInstant, "��Ѫִ�е�")
            objControl.BeginGroup = True
        End If
        '��Ҳ˵�
        Call DefCommandPlugInPopup(objMenu.CommandBar.Controls, mrsPlugInBar)
    End With

    '����˵�:���������û��,���ڲ鿴�˵�ǰ��
    '-----------------------------------------------------
    '����վ����˵��Զ���ʾ��������Թ���վ��ģ���ͳһ����
    '���⼸�ű�����ҽ������ģ���еģ���Ҫ�ڸ�ģ���е�������
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ReportPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ReportPopup, "����(&R)", objMenu.Index, False)
        objMenu.ID = conMenu_ReportPopup '��xtpControlPopup���͵�����ID�����¸�ֵ
    End If
    With objMenu.CommandBar.Controls
        If objMenu.CommandBar.Controls.Count > 0 Then objMenu.CommandBar.Controls(1).BeginGroup = True
        '���������ǰ��,�������
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill3, "ҽ����¼��(&5)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_AdviceBill1, "����ҽ����(&4)", 1): objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Report_ClinicBill, "��ӡ���Ƶ���(&3)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "��ӡִ�е�(&2)", 1)
        Set objControl = .Add(xtpControlButton, conMenu_Report_DrugQuery, "ҩ���շ���ѯ(&1)", 1)
    End With

    '�鿴�˵�
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ViewPopup)
    With objMenu.CommandBar.Controls
        Set objControl = .Find(, conMenu_View_StatusBar) '״̬����
        Set objControl = .Add(xtpControlButton, conMenu_View_Append, "������Ϣ(&A)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_Hide, "�Զ����ع���������(&H)", objControl.Index + 1)
        Set objControl = .Add(xtpControlButton, conMenu_View_AdviceLost, "ҽ��ˢ��ʱ��λ�����(&L)", objControl.Index + 1)
        
        Set objControl = .Find(, conMenu_View_Refresh) 'ˢ����ǰ(���ʱע�ⷴ��)
        Set objControl = .Add(xtpControlButton, conMenu_View_Notify, "ˢ������(&B)", objControl.Index)
        objControl.BeginGroup = True
    End With
    
    '���߲˵�:���������û��,���ڰ����˵�ǰ��
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_ToolPopup)
    If objMenu Is Nothing Then
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Find(, conMenu_HelpPopup)
        Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", objMenu.Index, False)
        objMenu.ID = conMenu_ToolPopup
    End If
    
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Tool_Sign, "����ǩ��(&S)", -1, False): objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ҽ��ǩ��(&I)")
            objControl.IconId = conMenu_Tool_Sign
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "��֤ǩ��(&V)")
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignEarse, "ȡ��ǩ��(&E)"): objControl.BeginGroup = True
        End With
        If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrint, "��ӡԤԼ��")
                objControl.IconId = 103
            Set objControl = .Add(xtpControlButton, conMenu_Tool_RisPrintBat, "������ӡԤԼ��")
        End If
        If gbln����ҩ�����հ������������� Then
            Set objControl = .Add(xtpControlButton, conMenu_Tool_Option, "ҽ��ѡ��(&O)"): objControl.BeginGroup = True
                objControl.IconId = conMenu_File_Parameter
        End If
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Define, "���׷�������(&S)"): objControl.BeginGroup = True
    End With

    '����������:���ļ�������˵������ť֮��ʼ����
    '-----------------------------------------------------
    Call AddToolBarInDoctor
     
    '����Ŀ����
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyA, conMenu_Edit_NewItem '�¿�ҽ��
        .Add FCONTROL, vbKeyM, conMenu_Edit_Modify '�޸�ҽ��
        .Add FSHIFT, vbKeyDelete, conMenu_Edit_Delete 'ɾ��ҽ��
        .Add FCONTROL, vbKeyS, conMenu_Edit_Stop 'ֹͣҽ��
        .Add FCONTROL, vbKeyV, conMenu_Edit_Audit 'ҽ��У��
        .Add FCONTROL, vbKeyI, conMenu_Edit_Price 'ҽ���Ƽ�
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send '����ҽ��
        .Add FCONTROL, vbKeyZ, conMenu_Edit_Untread 'ҽ������(xtpControlSplitButtonPopup��ʽʱ��ݼ���ʾ�����˵���,��������ʾ����)
        .Add FCONTROL, vbKeyT, conMenu_Edit_Test 'Ƥ�Խ��
        .Add FCONTROL, vbKeyR, conMenu_Edit_Compend  '���ı���
         
        .Add FCONTROL, vbKeyH, conMenu_View_Hide '�Զ����ع���������
        .Add FCONTROL, vbKeyL, conMenu_Manage_ReportLisView  '���������
        .Add FCONTROL, vbKeyP, conMenu_Manage_ReportPacsView '��������
        .Add 0, vbKeyF2, conMenu_Edit_SendInfusion '������ҺҩƷҽ�� �˲˵����޴��ڱ仯�У�������SetSendCommandBar�����б���ӻ��߲����
        .Add 0, vbKeyF9, conMenu_Report_AdviceBill1 'ҽ������ӡ
        .Add 0, vbKeyF10, conMenu_View_Notify 'ˢ��ҽ������
        .Add 0, vbKeyF11, conMenu_Tool_Option 'ҽ��ѡ��
    End With

End Sub

Public Sub zlPopupCommandBars(ByVal CommandBar As CommandBar)
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim vRoll As TYPE_AdviceRoll, i As Long
    Dim arrTmp As Variant, strTmp As String
    Dim lngҽ��ID As Long
    Dim rsTmp As ADODB.Recordset
    
    If CommandBar.Parent Is Nothing Then Exit Sub

    Select Case CommandBar.Parent.ID
    Case conMenu_Edit_CriticalAdvice
        If mblnΣ��ֵ And Not mrsΣ��ֵ Is Nothing Then
            mrsΣ��ֵ.Filter = 0
            If Not mrsΣ��ֵ.EOF Then
                Set rsTmp = GetCriticalAdvice(lngҽ��ID)
                With CommandBar.Controls
                    .DeleteAll
                    mrsΣ��ֵ.MoveFirst
                    For i = 1 To mrsΣ��ֵ.RecordCount
                        Set objControl = .Add(xtpControlButton, conMenu_Edit_CriticalAdvice * 100# + i, mrsΣ��ֵ!Σ��ֵ���� & "")
                            objControl.Parameter = mrsΣ��ֵ!ID & "," & lngҽ��ID
                        rsTmp.Filter = "Σ��ֵID=" & mrsΣ��ֵ!ID
                        If Not rsTmp.EOF Then
                            objControl.Checked = True
                        End If
                        mrsΣ��ֵ.MoveNext
                    Next
                    mrsΣ��ֵ.MoveFirst
                End With
            End If
            mrsΣ��ֵ.Filter = 0
        End If
    Case conMenu_Edit_Untread    'ҽ������
        With CommandBar.Controls
            .DeleteAll
            For i = 1 To UBound(marrRollList)
                vRoll = marrRollList(i)
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread * 100# + i, vRoll.��������)
                If i = 1 Then
                    If Not RollFirstEnabled Then objControl.Enabled = False
                Else
                    If i = 2 Then
                        objControl.BeginGroup = True
                    End If
                    objControl.Enabled = False
                End If
                If i = 50 Then Exit For    'ֻ����50����ʾ
            Next
        End With
    Case conMenu_ReportPopup
        Set objControl = CommandBar.FindControl(, conMenu_Report_ClinicBill)
        If Not objControl Is Nothing Then
            objControl.Visible = False
        End If
    Case conMenu_Edit_ChargeOff    '��������
        With CommandBar.Controls
            If .Count = 0 Then
                .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "��������(&L)").BeginGroup = True
                .Add xtpControlButton, conMenu_Edit_ChargeDelAudit, "�������(&U)"
                .Add(xtpControlButton, conMenu_Edit_ChargeOff * 10# + 1, "������ǰѡ��ĵ���(&1)").BeginGroup = True
                .Add xtpControlButton, conMenu_Edit_ChargeOff * 10# + 2, "������ǰҽ���ôη��͵ĵ���(&2)"
                .Add xtpControlButton, conMenu_Edit_ChargeOff * 10# + 3, "�����ôη��͵����е���(&3)"
            End If
        End With
    Case conMenu_Edit_Compend    '����
        With CommandBar.Controls
            If .Count = 0 Then
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 1, "���ı���(������ʽ)"
                .Add xtpControlButton, conMenu_Edit_Compend * 10# + 6, "���ı���(�����ʽ)"
                If gobjExchange Is Nothing Then
                    If mint���� = 1 Then    '��ʿվ
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)"
                    Else
                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 2, "��ӡ����(&P)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)"

                        .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "���Ѳ���(&R)").BeginGroup = True
                        .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "�Զ����(&A)"
                    End If
                End If
            End If
        End With
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99
        'PASSҩ�����
        If mblnPass Then
            Call gobjPass.zlPASSPopupCommandBars(mobjPassMap, CommandBar, conMenu_Edit_MediAudit)
        End If
    End Select
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As CommandBarControl)
    Dim lngParValue As Long, objControl As CommandBarControl
    Dim lng����ס����ҽ�� As Long
    Dim strErr As String
    
    mblnBatch = False
    mblnDirect = False
 
    Select Case Control.ID
    Case conMenu_File_PrintSet '��ӡ����
        Call zlPrintSet
    Case conMenu_File_Preview 'Ԥ��ҽ���嵥
        Call OutputList(2)
    Case conMenu_File_Print '��ӡҽ���嵥
        Call OutputList(1)
    Case conMenu_File_Excel '���ҽ���嵥
        Call OutputList(3)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_View_AdviceLost 'ҽ���Ƿ�λ���
        mblnҽ����λ��� = Not mblnҽ����λ���
        Call zlDatabase.SetPara("ҽ�����Ĭ�϶�λ�����һ��", IIF(mblnҽ����λ���, 1, 0), glngSys, pסԺҽ���´�)
    Case conMenu_View_Append '������Ϣ
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
    Case conMenu_View_Hide '�Զ����ع��˹�����
        mblnHideFilter = Not mblnHideFilter
        cbsSub(2).Visible = Not mblnHideFilter And cbsSub(2).Controls.Count > 0
        cbsSub(3).Visible = Not mblnHideFilter
        fraHide.Visible = mblnHideFilter
        cboTime.Visible = Not mblnHideFilter
        cbsSub.RecalcLayout
    Case conMenu_Edit_NewItem, conMenu_Edit_NewItem * 10# + 1 '�¿�ҽ��
        If Control.Parameter <> "" Then
            mlngΣ��ֵID = Val(Control.Parameter)
            Call GetCriticalData
        Else
            mlngΣ��ֵID = 0
        End If
        Call FuncAdviceAdd
    Case conMenu_Edit_Modify '�޸�ҽ��
        Call FuncAdviceModi
    Case conMenu_Edit_Delete, conMenu_Edit_ApplyDel 'ɾ��ҽ��'ȡ����������
        Call FuncAdviceDel
    Case conMenu_Edit_LISApply, conMenu_Edit_LISApply * 10 + 1   '��������
        Call FuncApplyLIS(0)
    Case conMenu_Edit_ApplyModi '�޸�����
        Call FuncApplyModi
    Case conMenu_Edit_NewRisSch 'RISԤԼ
        Call FuncAdviceRISSch
    Case conMenu_Edit_NewRisDel 'ȡ��ԤԼ
        Call FuncAdviceRISDel
    Case conMenu_Edit_NewRisModi
        Call FuncAdviceRISModi
    Case conMenu_Tool_RisPrint, conMenu_Tool_RisPrintBat
        Call FuncAdviceRISPrintSch(Control.ID)
    Case conMenu_Edit_PacsApply, conMenu_Edit_PacsApply * 10 + 1 '�������
        Call FuncApplyPACS(0, 0)
    Case conMenu_Edit_BloodApply, conMenu_Edit_BloodApply * 10 + 1  '��Ѫ����
        Call FuncApplyBlood(0)
    Case conMenu_Edit_OperationApply, conMenu_Edit_OperationApply * 10 + 1 '��������
        Call FuncApplyOperation(0)
    Case conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101#
        FuncApplyCustom 0, Control.Parameter
    Case conMenu_Edit_ConsultationApply, conMenu_Edit_ConsultationApply * 10 + 1 '��������
        Call FuncApplyConsultation(0)
    Case conMenu_Edit_TraReaction  '��Ѫ��Ӧ
        Call FuncTraReaction(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), pסԺҽ���´�, mblnMoved)
    Case conMenu_Edit_CriticalAdvice * 100# + 1 To conMenu_Edit_CriticalAdvice * 100# + 99
        Call FuncCriticalAdvice(Control.Parameter, Control.Checked)
    Case conMenu_Edit_ApplyView '�鿴����
        Call FuncApplyView
    Case conMenu_Edit_UnUse '���δ��ҽ��
        Call FuncAdviceUnUse
    Case conMenu_Edit_Sort   '����˳��
        Call FuncAdviceSort
    Case conMenu_Edit_Audit '���ҽ��,У��ҽ��
        If mint���� = 1 Then
            Call FuncAdviceVerify
        Else
            Call FuncAdviceAudit
        End If
    Case conMenu_Edit_MediAudit, conMenu_Edit_MediAudit * 10# To conMenu_Edit_MediAudit * 10# + 99  '������ҩ���
        If mblnPass Then
            Call gobjPass.zlPassCommandBarExe(mobjPassMap, Control.ID - conMenu_Edit_MediAudit * 10#)
        End If
    Case conMenu_Edit_Stop 'ֹͣҽ��
        Call FuncAdviceStop
    Case conMenu_Edit_StopAudit 'ͣ�����
        Call FuncAdviceStopAudit
    Case conMenu_Edit_Blankoff '����ҽ��
        Call FuncAdviceRevoke
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6 '���ġ���ӡ����
        Call FuncEPRReport(Control.ID)
    Case conMenu_Edit_Compend * 10# + 4 '���Ƿ��Ѿ����ĸñ���
        Call FuncExecReportRead(Not Control.Checked)
    Case conMenu_Edit_Compend * 10# + 5 '�Զ���ǲ���״̬
        mblnAutoRead = Not mblnAutoRead
        Call zlDatabase.SetPara("�Զ���Ǳ������״̬", IIF(mblnAutoRead, 1, 0), glngSys, pסԺҽ���´�)
    Case conMenu_Edit_MarkMap '��Ƭ����
        RaiseEvent ViewPACSImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    Case conMenu_Edit_MarkKeyMap '�ؼ�ͼ��
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowStaticImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_ViewPacs '������ͼ��ͱ���
        If CreateObjectPacs(mobjPublicPACS) Then
            Call mobjPublicPACS.ShowPatientImage(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
        End If
    Case conMenu_Edit_Price '�Ƽ۵���
        Call FuncAdvicePrice
    Case conMenu_Edit_ReStop 'ȷ��ֹͣ
        Call FuncAdviceConfirm(Control.Parameter = "ҽ������", Control)
    Case conMenu_Edit_Pause 'ҽ����ͣ
        Call FuncAdvicePause
    Case conMenu_Edit_Reuse 'ҽ������
        Call FuncAdviceResume
    Case conMenu_Edit_ClearUp 'ҽ������
        Call FuncAdviceReform
    Case conMenu_Edit_NoPrint '���δ�ӡ
        Call FuncAdviceNoPrint
    Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
        Call FuncViewDrugExplain(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�շ�ϸĿID)), mfrmParent)
    Case conMenu_Edit_Refcom '�ܾ��������
        Call FuncDrugRefcom 'ҩƷ���ܾ�����
    Case conMenu_Edit_ViewRefcom '�������δͨ����Ϣ
        If Not gobjPass Is Nothing And mlng����ID <> 0 And mlng��ҳID <> 0 Then Call gobjPass.ZLPharmReviewResultShow(Me, mlng����ID, mlng��ҳID)
    Case conMenu_Edit_Test 'Ƥ�Խ��
        Call FuncAdviceTest
    Case conMenu_Edit_Send, conMenu_Edit_SendInfusion 'ҽ������
        Call FuncAdviceSend(Control.Parameter = "ҽ������", Control)
    Case conMenu_Edit_SendCharge, conMenu_Edit_SendBilling 'ҽ����ҽ������������
        Call FuncAdviceSend(Control.Parameter = "ҽ������", Control)
    Case conMenu_Edit_BatExecute 'ҽ������ִ��
        '��鲡���Ƿ��������
        If Not CheckPatiIsAduit Then Exit Sub
        frmAdviceBatExecute.ShowMe 1, Me, mlng����ID, mlng����ID, mint����, 0, mlngҽ������ID, mlngӤ������ID, mlngӤ������ID
    Case conMenu_Manage_ThingAudit 'ҽ�������˶�
        '��鲡���Ƿ��������
        If Not CheckPatiIsAduit Then Exit Sub
        frmAdviceBatExecute.ShowMe 1, Me, mlng����ID, mlng����ID, mint����, 1, mlngҽ������ID, mlngӤ������ID, mlngӤ������ID
    Case conMenu_Edit_Surplus 'ҩƷ����Ǽ�
        Call frmDrugSurplus.ShowMe(mfrmParent, mlng����ID)
    Case conMenu_Edit_SendBack '���ڷ����ջ�
        Call FuncAdviceSendBack
    Case conMenu_Edit_Untread, conMenu_Edit_Untread * 100# + 1 'ҽ������(ֻ��˳�����)
        If Control.ID = conMenu_Edit_Untread Then
            '����鿴�����б���δ����
            If Not RollFirstEnabled Then Exit Sub
        End If
        Call FuncAdviceRoll
    Case conMenu_Edit_ChargeOff * 10# + 1 To conMenu_Edit_ChargeOff * 10# + 3 'ֱ�ӳ���
        Call FuncAdviceChargeOff(Control.ID - conMenu_Edit_ChargeOff * 10# - 1)
    Case conMenu_Tool_SignNew 'ҽ��ǩ��
        Call FuncAdviceSign
    Case conMenu_Tool_SignVerify '��֤ǩ��
        Call FuncAdviceSignVerify
    Case conMenu_Tool_SignEarse 'ȡ��ǩ��
        Call FuncAdviceSignErase
    Case conMenu_Report_ClinicBill * 100# + 1 To conMenu_Report_ClinicBill * 100# + 99 '��ӡ���Ƶ���
        Call FuncBillPrint(Control)
    Case conMenu_Edit_AdvicePrice
        RaiseEvent DoByAdvice(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)), conMenu_Edit_AdvicePrice, "")
    Case conMenu_Edit_ChargeDelApply, conMenu_Edit_ChargeDelAudit '�����������
        Call FuncAdviceReCharge(Control.ID)
    Case conMenu_Report_DrugQuery 'ҩ���շ���ѯ
        Call FuncDrugSendQuery
    Case conMenu_Report_Reports '�������ñ���
        Call FuncAdviceReport
    Case conMenu_Report_AdviceBill1 '����ҽ����
        Call frmAdvicePrint.ShowMe(mfrmParent, mlng����ID, mlng��ҳID)
    Case conMenu_Report_AdviceBill3 'ҽ����¼��
        Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1254_3", mfrmParent, "���˿���=" & mlng����ID)
    Case conMenu_Tool_Reference_2 '���ƴ�ʩ�ο�
        Call zlItemRef
    Case conMenu_Edit_BatUnPack '�������
        frmCompoundPack.ShowMe 1, Me, mlng����ID, mlng����ID, mlngҽ������ID, mlngӤ������ID, mlngӤ������ID
    Case conMenu_Tool_Option 'ҽ��ѡ��
        frmInAdviceSetup.Show 1, mfrmParent
    Case conMenu_Tool_Define '���׷�������
        Call FuncToolScheme
    Case conMenu_Manage_ReportLisView  '���鱨�����
        Call FuncViewLisRpt
Case conMenu_Manage_ReportPacsView  '��鱨�����
        Call FuncViewPacsRpt
    Case conMenu_Edit_MeetArrive
        Call Executeȷ�ϻ���(IIF(Control.Caption = "ȡ������ҽ������(&K)", True, False))
    Case conMenu_Manage_RecipeAuditView '�鿴���������
        If InitObjRecipeAudit(pסԺҽ���´�) Then
            Call gobjRecipeAudit.ShowResult(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)), mfrmParent)
        End If
    Case conMenu_Manage_ReportPrint
        Call PrintLisReport(mlng����ID, mfrmParent)
    Case conMenu_Report_BloodInstant
        Call PrintBloodReport(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mfrmParent)
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
        If CreatePlugInOK(pסԺҽ���´�, mint����) Then
            On Error Resume Next
            If PlugExeNew(Control.Parameter) = False Then
                Call gobjPlugIn.ExecuteFunc(glngSys, Decode(mint����, 0, pסԺҽ���´�, 1, pסԺҽ������, 2, pסԺҽ���´�), _
                    Control.Parameter, mlng����ID, mlng��ҳID, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), mlngǰ��ID, mint����)
                Call zlPlugInErrH(err, "ExecuteFunc")
                err.Clear: On Error GoTo 0
            End If
        End If
    End Select
End Sub


Private Function PlugExeNew(ByVal strName As String) As Boolean
'���ܣ����¼�����Ҳ�����ExecuteFunc����
    Dim lngID As Long
    Dim strXML As String
On Error GoTo errH
    If CreatePlugInOK(pסԺҽ���´�, mint����) Then
        With vsAdvice
            lngID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
            strXML = "<ROOT><������Ŀ����>" & .TextMatrix(.Row, col_������Ŀ����) & "</������Ŀ����></ROOT>"
            Call gobjPlugIn.ExecuteFunc(glngSys, Decode(mint����, 0, pסԺҽ���´�, 1, pסԺҽ������, 2, pסԺҽ���´�), strName, mlng����ID, mlng��ҳID, lngID, mlngǰ��ID, mint����, strXML)
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
    ByVal bln���� As Boolean, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal bytӤ�� As Byte, _
    ByVal lng����ID As Long, ByVal lng����id As Long, ByVal lngǰ��ID As Long, ByVal lng�������ID As Long, ByVal int���� As Integer, _
    ParamArray arrPar() As Variant)
'���ܣ��ṩ��������ҽ�������Ľӿ�
    Dim strErr As String
    
    Set mfrmParent = frmParent
    mMainPrivs = strPrivs
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mbytӤ�� = bytӤ��
    mlng����ID = lng����ID
    mlng����ID = lng����id
    mlngǰ��ID = lngǰ��ID
    mlng�������ID = lng�������ID
    mlngҽ������ID = lng�������ID
    mblnSignVisible = True
    If mint���� = 0 Then
        If CheckSign(1, 0, mlng�������ID, mlng����ID, 2, False, gobjESign) = False Then
            mblnSignVisible = False '��ͬ����û������Ҫʹ��ǩ��
        End If
    ElseIf mint���� = 2 Then
        If CheckSign(3, 0, mlng�������ID, mlng����ID, 2, False, gobjESign) = False Then
            mblnSignVisible = False '��ͬ����û������Ҫʹ��ǩ��
        End If
    ElseIf mint���� = 1 Then
        If CheckSign(2, mlngҽ������ID, , , , False, gobjESign) = False Then
            mblnSignVisible = False '��ͬ����û������Ҫʹ��ǩ��
        End If
    End If
    
    mint���� = int����
    mblnBatch = bln����
    mblnDirect = True
    mblnInsideTools = False
    
    mblnMoved = CheckPatiDataMoved(lng����ID, lng��ҳID)
    '����LIS����
    If Control.ID = conMenu_Manage_ReportLisView Or Control.ID = conMenu_Edit_Send Then
       Call InitObjLis(pסԺ��ʿվ)
    End If
    
    Select Case Control.ID
    Case conMenu_Edit_NewItem '�¿�ҽ��
        Call FuncAdviceAdd
    Case conMenu_Edit_Audit '���ҽ��,У��ҽ��
        If mint���� = 1 Then
            Call FuncAdviceVerify
        Else
            Call FuncAdviceAudit
        End If
    Case conMenu_Edit_Price '�Ƽ۵���
        Call FuncAdvicePrice
    Case conMenu_Edit_Send, conMenu_Edit_SendInfusion 'ҽ������
        Call FuncAdviceSend(Not bln����, Control)
    Case conMenu_Edit_Stop 'ֹͣҽ��
        Call FuncAdviceStop
    Case conMenu_Edit_StopAudit 'ͣ�����
        Call FuncAdviceStopAudit
    Case conMenu_Edit_ReStop 'ȷ��ֹͣ
        Call FuncAdviceConfirm(Not bln����, Control)
    
    Case conMenu_Edit_BatExecute 'ҽ������ִ��
        frmAdviceBatExecute.ShowMe 1, frmParent, lng����ID, lng����ID, mint����, 0, mlngҽ������ID, mlngӤ������ID, mlngӤ������ID
    Case conMenu_Manage_ThingAudit 'ҽ�������˶�
        frmAdviceBatExecute.ShowMe 1, frmParent, lng����ID, lng����ID, mint����, 1, mlngҽ������ID, mlngӤ������ID, mlngӤ������ID
        
    Case conMenu_Edit_Blankoff '����ҽ��
        Call FuncAdviceRevoke
        
        
    Case conMenu_Edit_Pause 'ҽ����ͣ
        Call FuncAdvicePause
    Case conMenu_Edit_Reuse 'ҽ������
        Call FuncAdviceResume
        
        
    Case conMenu_Edit_SendBack '���ڷ����ջ�
        Call FuncAdviceSendBack
    Case conMenu_Report_DrugQuery 'ҩ���շ���ѯ
        Call FuncDrugSendQuery
    
    Case conMenu_Manage_ReportLisView  '���鱨�����
        If mlng����ID <> 0 Then
            If Not gobjLIS Is Nothing And Sys.SystemShareWith(2500) Then
                gobjLIS.PatientSampleBrowse mfrmParent, mlng����ID, mMainPrivs, mlng����ID, mlng����ID, 2, mlng��ҳID
            Else
                frmLisView.ShowMe mlng����ID, pסԺҽ���´�, mfrmParent
            End If
        End If
    Case conMenu_Edit_Surplus 'ҩƷ����Ǽ�
        Call frmDrugSurplus.ShowMe(mfrmParent, mlng����ID)
        
    Case conMenu_Report_Reports '�������ñ���
        Call FuncAdviceReport
        
    End Select
End Sub

Private Sub FuncAdviceReport()
'���ܣ����ò������ñ���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    '�ж����ز���������д˲�������ʹ���ϰ�ִ�е���ӡ����
    strSQL = "select 1 from zlParameters a where a.ϵͳ=[1] and a.ģ��=[2] and a.������=[3]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngSys, pסԺҽ������, "��ԭ�ϰ�ִ�е���ӡ����")
    
    
    On Error Resume Next
    If rsTmp.EOF Then
        Call frmAdviceWardReport.ShowMe(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng����ID, mlng����ID)
    Else
        Call frmAdviceReport.ShowMe(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng����ID, mlng����ID, mblnDirect And Not mblnBatch Or mblnInsideTools, mlngҽ������ID, mlngӤ������ID)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSendBack()
'���ܣ����ڷ����ջ�
    Dim blnRoll As Boolean
    
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    '68081��Ժ���������ҽ���������ñ仯
    If mintPState = psԤ�� Or mintPState = ps��Ժ Then
        Call MsgBox("�ò�����" & IIF(mintPState = psԤ��, "Ԥ", "") & "��Ժ�����������ҽ�������ջأ�", vbInformation, gstrSysName)
        Exit Sub
    End If
    On Error Resume Next
    blnRoll = frmAdviceRollSend.ShowMe(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng����ID, mlng����ID, mlng��ҳID, mblnDirect And Not mblnBatch Or mblnInsideTools, False, mlngҽ������ID, mlngӤ������ID)
    
    If blnRoll And mblnDirect = False Then
        RaiseEvent StatusTextUpdate("")
        Call LoadAdvice
    End If
End Sub

Public Sub zlCheckPrivs(ByVal Control As CommandBarControl, ByVal int���� As Integer)
'���ܣ����˵���ť��Ȩ�ޣ���������ɼ���
    mint���� = int����
    Call SetControlVisible(Control)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�ޡ���ǰ���˻�������������ù��ܻ�ɼ��Ϳ�����
'  1.�޲��˵����
'  2.�����ѳ�Ժ�����
'  3.�����ݵ����
    Dim vRoll As TYPE_AdviceRoll
    Dim blnAdvice As Boolean, blnEnabled As Boolean, blnEdit As Boolean, bln��¼ As Boolean
    Dim i As Integer
    
    tbcMain.Enabled = mlng����ID <> 0
    For i = 0 To tbcMain.ItemCount - 1
        tbcMain.Item(i).Enabled = mlng����ID <> 0
    Next
    
    If vsAdvice.Redraw = flexRDNone Then Exit Sub
    'Pass
    '����˴������ƣ��� control.Id ������[conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99]������� ʱ,����ҽ���������ֺͰ�ť�ɼ�״̬�л�ı��Pass
    'Enabled����ֵ�������ڶ������������õ�enabled��ֵ���ᱻ���ǡ�
    If Between(Control.ID, conMenu_Edit_MediAudit * 10#, conMenu_Edit_MediAudit * 10# + 99) Then
        Control.Visible = IIF(Control.Category <> "", InStr(Control.Category, ";�ɼ�;") > 0, True)
        Control.Enabled = IIF(Control.Category <> "", InStr(Control.Category, ";����;") > 0, True)
        Exit Sub
    End If
    
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible Then Exit Sub
    
    'ҽ����������
    '------------------------------------------------------------------------------
    '�ܵ��ж�:�޲��˻��ѻ��ﲡ�˲������κβ���
    If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 998) _
        Or Between(Control.ID, conMenu_Edit_NewItem * 10#, (conMenu_Edit_NewItem + 998) * 10# + 9) Or Control.ID = conMenu_Manage_ThingAudit Then  '���������Ӳ˵�
        
        Control.Enabled = mlng����ID <> 0 And mintPState <> ps���� And mintPState <> ps��ת�� _
            And (InStr(",0,3,", mintִ��״̬) > 0 Or Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Or Control.ID = conMenu_Edit_MarkKeyMap Or Control.ID = conMenu_Edit_Compend _
                Or Between(Control.ID, conMenu_Edit_Compend * 10# + 1, conMenu_Edit_Compend * 10# + 6))
        If Not Control.Enabled Then Exit Sub
    End If
    
    blnAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
    blnEdit = (mintPState = ps��Ժ Or mintPState = ps���� Or mintPState = ps��Ժ)
    bln��¼ = (mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ)
    
    Select Case Control.ID
    Case conMenu_Edit_NewItem, conMenu_Edit_LISApply, conMenu_Edit_PacsApply, conMenu_Edit_BloodApply, conMenu_Edit_OperationApply, conMenu_Edit_ConsultationApply, conMenu_Edit_ApplyCustom * 100# To conMenu_Edit_ApplyCustom * 101# '�¿�ҽ��(����סԺ�����˳�·���Ĳ��˲������¿����޸�)
        Control.Enabled = (blnEdit Or bln��¼)
        
    Case conMenu_Edit_Sort  '����˳��
        Control.Enabled = blnEdit
    Case conMenu_Edit_Modify, conMenu_Edit_Delete '�޸�ҽ��,ɾ��ҽ��
        With vsAdvice
            blnEnabled = blnAdvice
            If blnEnabled Then
                If Control.ID = conMenu_Edit_Modify Then
                    blnEnabled = (blnEdit Or bln��¼)
                End If
            End If
            If blnEnabled Then
                If InStr(",1,2,", .TextMatrix(.Row, COL_ҽ��״̬)) = 0 Then blnEnabled = False
            End If
            If blnEnabled Then
                If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then blnEnabled = False
            End If
            If blnEnabled Then
                '�ٴ���ҽ�����ܻ������
                If mint���� = 2 Then
                    blnEnabled = InStr("," & mstrǰ��IDs & ",", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) & ",") > 0
                Else
                    blnEnabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) = 0
                End If
                
            End If
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Manage_RecipeAuditView
        blnEnabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�������״̬)) <> 0
        Control.Enabled = blnEnabled
    '�������뵥���޸�
    Case conMenu_Edit_ApplyModi
        With vsAdvice
            blnEnabled = blnAdvice
            If blnEnabled Then
                If Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E" Then
                    If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E") Then blnEnabled = False
                ElseIf .TextMatrix(.Row, COL_�������) = "D" And .TextMatrix(.Row, COL_��������) <> "����" Then
                    If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And .TextMatrix(.Row, COL_�������) = "D") Then blnEnabled = False
                ElseIf .TextMatrix(.Row, COL_�������) = "K" Then
                    If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And .TextMatrix(.Row, COL_�������) = "K") Then blnEnabled = False
                ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                    If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And .TextMatrix(.Row, COL_�������) = "F") Then blnEnabled = False
                ElseIf Val(.TextMatrix(.Row, COL_��������)) = 7 And .TextMatrix(.Row, COL_�������) = "Z" Then
                    If Not .TextMatrix(.Row, COL_ҽ��״̬) = "1" Then blnEnabled = False
                Else
                    blnEnabled = Val(.TextMatrix(.Row, COL_�������)) <> 0
                End If
                
            End If
            Control.Enabled = blnEnabled
        End With
    Case conMenu_Edit_NewRis
        blnEnabled = False
        With vsAdvice
            If InStr(",D,F,", .TextMatrix(.Row, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_��������))) > 0 And .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��Ч) = "����" Then
                blnEnabled = True
            End If
        End With
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisSch
        blnEnabled = False
        If gbln����Ӱ����ϢϵͳԤԼ Then
            With vsAdvice
                If (InStr(",D,F,", .TextMatrix(.Row, COL_�������)) > 0 Or InStr(",0,5,", Val(.TextMatrix(.Row, COL_��������))) > 0 And .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��Ч) = "����") And Val(.TextMatrix(.Row, COL_RISԤԼID)) = 0 Then
                    blnEnabled = True
                End If
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_NewRisDel, conMenu_Tool_RisPrint
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RISԤԼID)) <> 0
    Case conMenu_Edit_NewRisModi
        Control.Enabled = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RISԤԼID)) <> 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 8
    Case conMenu_Tool_RisPrintBat
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
        Control.Enabled = blnAdvice And InStr(",5,6,7,", vsAdvice.TextMatrix(vsAdvice.Row, COL_�������)) > 0
    '���뵥ȡ��
    Case conMenu_Edit_ApplyDel
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (.TextMatrix(.Row, COL_ҽ��״̬) = "1" And _
                    (.TextMatrix(.Row, COL_�������) = "D" Or .TextMatrix(.Row, COL_�������) = "F" Or Val(.TextMatrix(.Row, COL_��������)) = 7 And .TextMatrix(.Row, COL_�������) = "Z" Or _
                        Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E" Or _
                        .TextMatrix(.Row, COL_�������) = "K")) Then
                    blnEnabled = Val(.TextMatrix(.Row, COL_�������)) <> 0
                End If
                '��Ѫҽ������˲�����ȡ������Ѫ���������ݣ�
                If blnEnabled = True And .TextMatrix(.Row, COL_�������) = "K" And .TextMatrix(.Row, COL_ҽ��״̬) = "1" Then
                    If Val(.TextMatrix(.Row, COL_��鷽��)) = 1 And Val(.TextMatrix(.Row, COL_���״̬)) = 1 Then blnEnabled = False
                End If
            End With
        End If
        Control.Enabled = blnEnabled
    '�鿴����
    Case conMenu_Edit_ApplyView
        blnEnabled = blnAdvice
        If blnEnabled Then
            With vsAdvice
                If Not (InStr(",F,K,D,", .TextMatrix(.Row, COL_�������)) > 0 Or Val(.TextMatrix(.Row, COL_��������)) = 7 And .TextMatrix(.Row, COL_�������) = "Z") Then blnEnabled = Val(.TextMatrix(.Row, COL_�������)) <> 0
            End With
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_UnUse '���δ��ҽ��
        With vsAdvice
            Control.Checked = Val(.TextMatrix(.Row, COL_ִ�б��)) = -1
            If Val(.TextMatrix(.Row, COL_ִ�б��)) = -1 Then
                Control.Enabled = True
            Else
                blnEnabled = blnAdvice
                If blnEnabled Then
                    blnEnabled = mintPState = ps��Ժ Or mintPState = psԤ�� Or mintPState = ps����
                End If
                If blnEnabled Then 'δУ�ԡ������ϵ�ҽ����������
                    If InStr(",1,2,4,", .TextMatrix(.Row, COL_ҽ��״̬)) > 0 Then blnEnabled = False
                End If
                If blnEnabled Then '�ѷ��͵ĳ�����������
                    If .TextMatrix(.Row, COL_��Ч) = "����" And .TextMatrix(.Row, COL_�ϴ�ִ��) <> "" Then blnEnabled = False
                End If
                If blnEnabled Then '����Ƥ�Խ���Ĳ�������(�൱��ִ����)
                    If .TextMatrix(.Row, COL_Ƥ��) <> "" Then blnEnabled = False
                End If
                Control.Enabled = blnEnabled
            End If
        End With
    Case conMenu_Edit_Stop, conMenu_Edit_Blankoff 'ֹͣҽ��,ҽ������
        If mint���� = 2 Then 'ҽ��ҽ������
            With vsAdvice
                blnEnabled = blnAdvice _
                    And InStr(",1,2,4,8,9,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) = 0 _
                    And (Val(.TextMatrix(.Row, COL_ǩ����)) = 0 Or Not gobjESign Is Nothing) _
                    And InStr("," & mstrǰ��IDs & ",", "," & Val(.TextMatrix(.Row, COL_ǰ��ID)) & ",") > 0 _
                    And .TextMatrix(.Row, COL_����ҽ��) = UserInfo.����
                
                If blnEnabled Then
                    If Control.ID = conMenu_Edit_Stop Then
                        '����(������ҩ�䷽)
                        blnEnabled = .TextMatrix(.Row, COL_��Ч) = "����" And .TextMatrix(.Row, COL_����) = ""
                    ElseIf Control.ID = conMenu_Edit_Blankoff Then
                        'δ���Ͳſ�����
                        blnEnabled = .TextMatrix(.Row, COL_�ϴ�ִ��) = ""
                    End If
                End If
            End With
            Control.Enabled = blnEnabled
        End If
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1 To conMenu_Edit_Compend * 10# + 3, conMenu_Edit_Compend * 10# + 6 '���ġ���ӡ����
        If Not gobjExchange Is Nothing Then
            Control.Enabled = blnAdvice And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) <> 0 And vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬) <> "4"
        Else
            Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_��鱨��ID) <> "" Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS����ID)) <> 0) And vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬) <> "4"
        End If
        If Control.ID = conMenu_Edit_Compend * 10# + 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 1 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        ElseIf Control.ID = conMenu_Edit_Compend * 10# + 6 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 2 Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
        End If
    Case conMenu_Edit_Compend * 10# + 4 '���Ѿ����ĸñ���
        Control.Checked = Val(vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_����״̬)) = 1
        Control.Enabled = blnAdvice And (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)) <> 0 Or vsAdvice.TextMatrix(vsAdvice.Row, COL_��鱨��ID) <> "")
    Case conMenu_Edit_Compend * 10# + 5 '�Զ���ǲ���״̬
        Control.Checked = mblnAutoRead
        Control.Enabled = mblnAutoReadEnabled
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs '��Ƭ����
        blnEnabled = blnAdvice And InStr(",4,5,6,7,8,9,H,M,Z,", vsAdvice.TextMatrix(vsAdvice.Row, COL_�������)) = 0 ' And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID)) <> 0
        If blnEnabled Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) <> 8 Then
                blnEnabled = False
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_Audit, conMenu_Edit_Price 'ҽ��У��(ҽ�����),�Ƽ۵���
        Control.Enabled = (blnEdit Or bln��¼)
    Case conMenu_Edit_Pause, conMenu_Edit_Reuse 'ҽ����ͣ,ҽ������
        Control.Enabled = (mintPState = ps��Ժ Or mintPState = ps����)
    
    Case conMenu_Edit_Untread 'ҽ������(�����ڵ���ʱ�����ÿ���״̬)
        Control.Enabled = UBound(marrRollList) >= 1 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ�б��)) <> -1
        If Control.Enabled And Not RollFirstEnabled Then
            Control.IconId = conMenu_Edit_Untread * 100# + 99 '��ʾ�е������Ի���
        Else
            Control.IconId = conMenu_Edit_Untread
        End If
    Case conMenu_Edit_AdvicePrice
        Control.Enabled = (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) > 4 Or Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 3)
    Case conMenu_Edit_ChargeDelApply, conMenu_Edit_ChargeDelAudit '�����������
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Edit_ChargeOff * 10# + 1 To conMenu_Edit_ChargeOff * 10# + 3 'ֱ�ӳ���
        blnEnabled = False
        If tbcAppend.Selected.Tag = "����" And mblnAppend Then
            If Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("���ͺ�"))) <> 0 Then
                If Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("��¼����"))) = 2 Then
                    blnEnabled = True
                End If
            End If
        End If
        Control.Enabled = blnAdvice And blnEnabled
    Case conMenu_Edit_Test 'Ƥ�Խ��:���ͺ���ܱ�ע
        With vsAdvice
            Control.Enabled = blnAdvice And Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 _
                And .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "1"
        End With
    Case conMenu_Edit_ReStop, conMenu_Edit_ClearUp 'ȷ��ֹͣ,ҽ������
    Case conMenu_Edit_NoPrint '���δ�ӡ
        Control.Enabled = blnAdvice And Control.Visible
        If Control.Enabled Then
            Control.Checked = Val(vsAdvice.ValueMatrix(vsAdvice.Row, COL_���δ�ӡ)) = 1
        End If
    Case conMenu_Edit_Send  '����
        If mint���� <> 1 Then 'ҽ��ҽ����������
            Control.Enabled = (blnEdit Or bln��¼)
        End If
    Case conMenu_Edit_SendBilling
        If InStr(GetInsidePrivs(pסԺҽ���´�), "�����������") > 0 Then
            If mlng�������� = 1 Then    '�������۲��ˣ�ֻ�ܷ���Ϊ������ʵ�
                Control.Caption = "�������"
            Else
                Control.Caption = "סԺ����"
            End If
        End If
    Case conMenu_Edit_TraReaction
        With vsAdvice
            Control.Enabled = (.TextMatrix(.Row, COL_�������) = "K") And Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 And gblnѪ��ϵͳ
        End With
    Case conMenu_Edit_CriticalAdvice
        blnEnabled = False
        If Not mrsΣ��ֵ Is Nothing Then
            If Not mrsΣ��ֵ.EOF Then
                blnEnabled = True
            End If
        End If
        If blnEnabled Then
            blnEnabled = (Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) <> 4 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0)
        End If
        Control.Enabled = blnEnabled
    Case conMenu_Edit_MediAudit 'ҩ�����(��ҩ����ʾ)
        If mblnPass Then
            Call gobjPass.zlPassCommandBarUpdate(mobjPassMap, Control, blnAdvice)
        End If
    Case conMenu_Edit_MeetArrive
        Control.Caption = IIF(mblnȷ�ϻ���, "ȡ������ҽ������(&K)", "ȷ�ϻ���ҽ������(&M)")
        Control.Enabled = True
    End Select
    
    '���˷�Χҽ���´���
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        '�¿�ҽ��
        If Control.Enabled Then Control.Enabled = PatiCanAdvice
    Case conMenu_Edit_Modify, conMenu_Edit_Delete
        '�޸�ҽ��,ɾ��ҽ��
        If Control.Enabled And mint���� = 2 Then
            Control.Enabled = PatiCanAdvice
        ElseIf Control.Enabled Then
            Control.Enabled = PatiCanAdvice
        End If
    Case conMenu_Edit_ClearUp, conMenu_Edit_Untread
        'ҽ������,ҽ������
        If mint���� = 0 Then
            If Control.Enabled Then Control.Enabled = PatiCanAdvice
        End If
    End Select
            
    'ҽ��������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Report_ClinicBill '��ӡ���Ƶ���
        Control.Enabled = Control.CommandBar.Controls.Count > 0
    Case conMenu_Report_Reports, conMenu_Report_DrugQuery, conMenu_Report_MultiBill '�����̶�����
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Report_AdviceBill1 To conMenu_Report_AdviceBill3 '����ҽ����,��ʱҽ����,����ҽ����
        Control.Enabled = mlng����ID <> 0
    End Select
    
    '����ǩ������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_SignNew 'ҽ��ǩ��
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Tool_SignVerify '��֤ǩ��
        blnEnabled = mlng����ID <> 0 And blnAdvice And tbcAppend.Selected.Tag = "ǩ��" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        If blnEnabled Then blnEnabled = vsAppend.RowData(vsAppend.Row) <> 0
        Control.Enabled = blnEnabled
    Case conMenu_Tool_SignEarse 'ȡ��ǩ��
        blnEnabled = mlng����ID <> 0 And blnAdvice And tbcAppend.Selected.Tag = "ǩ��" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0
        If blnEnabled Then blnEnabled = vsAppend.RowData(vsAppend.Row) <> 0 And vsAppend.Cell(flexcpData, vsAppend.Row, 0) <> 3
        Control.Enabled = blnEnabled
    End Select
    
    '��������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = blnAdvice
    Case conMenu_View_Append '������Ϣ
        Control.Checked = tbcAppend.Visible
    Case conMenu_View_AdviceLost 'ҽ���Ƿ�λ���
        Control.Checked = mblnҽ����λ���
    Case conMenu_View_Hide '�Զ����ع��˹�����
        Control.Checked = mblnHideFilter
    Case conMenu_Manage_ReportLisView, conMenu_Manage_ReportPacsView  '���鱨�����
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Manage_ReportPrint  '���鱨��������ӡ
        Control.Enabled = mlng����ID <> 0
    Case conMenu_Report_BloodInstant  'ִ�е���ӡ
        Control.Visible = InStr(GetInsidePrivs(9005, , 2200), ";��Ѫִ�д�ӡ;") <> 0
        Control.Enabled = vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "K" And Control.Visible
    End Select
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strItem As String, blnSendPriv As Boolean

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" And Control.ID <> conMenu_Edit_SendBilling And Control.ID <> conMenu_Edit_Audit And Control.ID <> conMenu_Edit_MeetArrive Then Exit Sub

    blnVisible = True
    
    '���Ȩ���ж�
    '------------------------------------------------------------------------------
    If mint���� = 0 And InStr(UserInfo.����, "ҽ��") = 0 _
        Or mint���� = 1 And InStr(UserInfo.����, "��ʿ") = 0 Then
        If Control.ID = conMenu_EditPopup Then blnVisible = False
        If Control.ID = conMenu_ReportPopup Then blnVisible = False
        If Between(Control.ID, conMenu_Edit_NewItem, conMenu_Edit_NewItem + 999) Then blnVisible = False
    End If
    
    'ҽ����������
    '------------------------------------------------------------------------------
    If mint���� = 0 Or mint���� = 2 Then
        Select Case Control.ID
        Case conMenu_Edit_Untread
            'ҽ������
            If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ���´�;") = 0 Then blnVisible = False
        'Case conMenu_Edit_Send  'û����������Ȩ��ʱ��Ҳ��û�з��������շ�Ȩ�ޣ����Ǵ�����conMenu_Edit_SendBilling
         
        Case conMenu_Edit_SendBilling
            If mlng�������� = 1 Then  '����סԺ����û�е�������Ȩ��
                If InStr(GetInsidePrivs(pסԺҽ���´�), ";�����������;") = 0 Then blnVisible = False
            Else
                If InStr(GetInsidePrivs(pסԺҽ���´�), ";��������;") = 0 Then blnVisible = False
            End If
        Case conMenu_Edit_SendCharge
            If InStr(GetInsidePrivs(pסԺҽ���´�), ";�����������;") = 0 Then blnVisible = False
        Case conMenu_Edit_BatExecute
            'ҽ������ִ��
            If InStr(GetInsidePrivs(pסԺҽ���´�), ";����ִ�еǼ�;") = 0 Then blnVisible = False
        Case conMenu_Edit_NoPrint
            If InStr(GetInsidePrivs(pסԺҽ���´�), ";���δ�ӡ;") = 0 Then
                blnVisible = False
            Else
                blnVisible = True
            End If
            Control.Enabled = blnVisible
        Case conMenu_Edit_TraReaction  '��Ѫ��Ӧ�Ǽ�
            If gblnѪ��ϵͳ Then '��Ѫ��ϵͳĬ����ʾ
                If InStr(GetInsidePrivs(9005, , 2200), ";��Ѫ��Ӧ�Ǽ�;") = 0 Then
                    blnVisible = False
                Else
                    blnVisible = True
                End If
                Control.Enabled = blnVisible
            End If
        End Select

    ElseIf mint���� = 1 Then
        strItem = GetInsidePrivs(pסԺҽ������)
        blnSendPriv = InStr(strItem, ";����ҩ������;") > 0 Or InStr(strItem, ";����ҩ�Ƴ���;") > 0 _
                        Or InStr(strItem, ";������������;") > 0 Or InStr(strItem, ";������������;") > 0
                
        Select Case Control.ID
        Case conMenu_Edit_Untread
            'ҽ������
            If InStr(strItem, ";ҽ��״̬����;") = 0 Then blnVisible = False
        Case conMenu_Edit_Send
            'ҽ������
            If Not blnSendPriv Then blnVisible = False
        Case conMenu_Edit_BatExecute, conMenu_Manage_ThingAudit
            'ҽ������ִ��
            If InStr(strItem, ";����ִ�еǼ�;") = 0 Then blnVisible = False
            If blnVisible And Control.ID = conMenu_Manage_ThingAudit Then
                If Val(gstrҽ���˶�) = 0 Then blnVisible = False
            End If
        Case conMenu_Edit_MeetArrive
            With vsAdvice
                blnVisible = Val(.TextMatrix(.Row, COL_�������)) <> 0 And Val(.TextMatrix(.Row, COL_��������)) = 7 And .TextMatrix(.Row, COL_�������) = "Z" And .TextMatrix(.Row, COL_״̬) = "ֹͣ"
            End With
        Case conMenu_Edit_NoPrint
            If InStr(strItem, ";���δ�ӡ;") = 0 Then
                blnVisible = False
            Else
                blnVisible = True
            End If
            Control.Enabled = blnVisible
        End Select
    End If
    
    Select Case Control.ID
    Case conMenu_Edit_ClearUp
        'ҽ������
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ������;") = 0 Then blnVisible = False
    Case conMenu_Edit_NewItem, conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Sort, conMenu_Edit_LISApply, conMenu_Edit_ApplyDel, conMenu_Edit_ApplyModi, conMenu_Edit_Apply, conMenu_Edit_ApplyDel, conMenu_Edit_ApplyView
        '�¿�ҽ��,�޸�ҽ��,ɾ��ҽ�� ,����˳��,�������롢�޸ġ�ɾ��
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ���´�;") = 0 Then blnVisible = False
        
    Case conMenu_Edit_UnUse 'δ��ҽ��
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";���δ��ҽ��;") = 0 Then blnVisible = False
    Case conMenu_Edit_Pause, conMenu_Edit_Reuse
        'ҽ����ͣ,ҽ������
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ����ͣ;") = 0 Then blnVisible = False
    Case conMenu_Edit_Stop
        'ֹͣҽ��
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ��ֹͣ;") = 0 Then blnVisible = False
    Case conMenu_Edit_Blankoff
        'ҽ������
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ������;") = 0 Then blnVisible = False
    Case conMenu_Edit_Compend, conMenu_Edit_Compend * 10# + 1, conMenu_Edit_Compend * 10# + 6
        '���浯��(����ӡ),���ı���
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";�������;") = 0 Then blnVisible = False
    Case conMenu_Edit_Compend * 10# + 2, conMenu_Edit_Compend * 10# + 3
        '��ӡ����
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";�����ӡ;") = 0 Then blnVisible = False
    Case conMenu_Edit_MarkMap, conMenu_Edit_MarkKeyMap, conMenu_Edit_ViewPacs
        '��Ƭ����
        If GetInsidePrivs(pXWPACS��Ƭ) <> "" And InStr(GetInsidePrivs(pסԺҽ���´�), ";��Ƭ����;") <> 0 Then
            blnVisible = True
        Else
            If Control.ID = conMenu_Edit_MarkMap Or Control.ID = conMenu_Edit_ViewPacs Then
                If InStr(GetInsidePrivs(pסԺҽ���´�), ";��Ƭ����;") = 0 Or GetInsidePrivs(p��Ƭ���߹���) = "" Then
                    blnVisible = False
                End If
            Else
                blnVisible = False
            End If
        End If
    Case conMenu_Edit_Audit
        If mint���� = 1 Then
            'ҽ��У��
            If Val(zlDatabase.GetPara("����ǰ�Զ�У��", glngSys, pסԺҽ������, 0)) = 1 Then
                If InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��У�Դ���;") > 0 And Not blnSendPriv Then
                    '�������۲��˵ĳ�������ʳ��Ӫ���ȣ���ֻУ�Բ�����
                    blnVisible = True
                Else
                    blnVisible = False
                End If
                Control.Enabled = blnVisible
            Else
                If InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��У�Դ���;") = 0 Then blnVisible = False
            End If
        Else
            'ҽ�����:��Ȩ�޻򲻾����ʸ�ʱ
            If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ�����;") = 0 Or Not mblnHaveAuditPriv Then blnVisible = False
        End If
    Case conMenu_Edit_StopAudit
        'ͣ����˺�������˹���Ȩ��
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ�����;") = 0 Or Not mblnHaveAuditPriv Then blnVisible = False
    Case conMenu_Edit_Price
        '�Ƽ۵���
        If InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��У�Դ���;") = 0 Then blnVisible = False
    Case conMenu_Edit_ReStop
        'ȷ��ֹͣ
        If InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��ȷ��ֹͣ;") = 0 Then blnVisible = False
    Case conMenu_Edit_Test
        'Ƥ�Խ��
        If InStr(GetInsidePrivs(pסԺҽ������), ";Ƥ��ҽ�����;") = 0 Then blnVisible = False
    Case conMenu_Edit_SendBack
        '���������ջ�
        If InStr(GetInsidePrivs(pסԺҽ������), ";���ڷ����ջ�;") = 0 Then blnVisible = False
    Case conMenu_Edit_ViewDrugExplain '�鿴ҩƷ˵����
        If gobjDrugExplain Is Nothing Or InStr(GetInsidePrivs(pסԺҽ���´�), ";ҩƷ˵����;") = 0 Then blnVisible = False
    Case conMenu_Edit_ChargeDelApply
        '��������
        '55380
        strItem = GetInsidePrivs(pסԺ���ʲ���)
        If InStr(strItem, ";ҩƷ��������;") = 0 _
            And InStr(strItem, ";������������;") = 0 _
            And InStr(strItem, ";������������;") = 0 Then blnVisible = False
    Case conMenu_Edit_ChargeDelAudit
        '�������
        strItem = GetInsidePrivs(pסԺ���ʲ���)
        If InStr(strItem, ";�������;") = 0 Then blnVisible = False
    Case conMenu_Edit_Surplus
        'ҩƷ����Ǽ�
        strItem = GetInsidePrivs(pסԺҽ������)
        If InStr(strItem, ";ҩƷ����Ǽ�;") = 0 Then blnVisible = False
    Case conMenu_Edit_MediAudit
        '������ҩ���
        strItem = GetInsidePrivs(pסԺҽ���´�)
        If InStr(strItem, "������ҩ���") = 0 Then blnVisible = False
    End Select
    'ҽ��������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Report_DrugQuery 'ҩ���շ���ѯ
        If InStr(GetInsidePrivs(pסԺҽ������), ";ҩ���շ���ѯ;") = 0 Then blnVisible = False
    Case conMenu_Report_AdviceBill1 '����ҽ����,��ʱҽ����
        blnVisible = False
        If InStr(UserInfo.����, "ҽ��") > 0 Then
            If InStr(GetInsidePrivs(pסԺҽ���´�), "����ҽ����") > 0 Or InStr(GetInsidePrivs(pסԺҽ���´�), "��ʱҽ����") > 0 Then
                blnVisible = True
            End If
        End If
        If Not blnVisible Then
            If InStr(UserInfo.����, "��ʿ") > 0 Then
                If InStr(GetInsidePrivs(pסԺҽ������), "����ҽ����") > 0 Or InStr(GetInsidePrivs(pסԺҽ������), "��ʱҽ����") > 0 Then
                    blnVisible = True
                End If
            End If
        End If
    Case conMenu_Report_AdviceBill3 '����ҽ����
        blnVisible = False
        If InStr(UserInfo.����, "ҽ��") > 0 Then
            If InStr(GetInsidePrivs(pסԺҽ���´�), "����ҽ����") > 0 Then
                blnVisible = True
            End If
        End If
        If Not blnVisible Then
            If InStr(UserInfo.����, "��ʿ") > 0 Then
                If InStr(GetInsidePrivs(pסԺҽ������), "����ҽ����") > 0 Then
                    blnVisible = True
                End If
            End If
        End If
    End Select
    
    Control.Category = "���ж�"
    '����ǩ������
    '------------------------------------------------------------------------------
    Select Case Control.ID
    Case conMenu_Tool_Sign, conMenu_Tool_SignNew '����ǩ��,ҽ��ǩ��
        If gobjESign Is Nothing Or (InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ���´�;") = 0 And InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��У�Դ���;") = 0) _
            Or Not mblnHaveAuditPriv Then
            blnVisible = False
        ElseIf mblnSignVisible = False Then
            blnVisible = False '��ͬ����û������Ҫʹ��ǩ��
        End If
        Control.Category = ""  'ǩ����ť��̬�жϿɼ���
    End Select

    Control.Enabled = blnVisible
    Control.Visible = blnVisible
End Sub

Private Function PatiCanAdvice() As Boolean
'���ܣ����Ե�ǰ�����Ƿ�����´�ҽ��
'˵������Ҫ�����´��Ȩ��ʱ,�ټ�鱾�ƺ�ȫԺ���˷�Χ
    Dim strPriv As String, bln�´� As Boolean
    
    strPriv = GetInsidePrivs(pסԺҽ���´�)
    If mlng����ID <> 0 Then
        If mintPState = ps���� Then
            bln�´� = True '�����ﲡ���������
        ElseIf mint���� = 0 Then
            If mstrסԺҽ�� = UserInfo.���� Then
                bln�´� = True '��ǰҽ�����β���
            ElseIf InStr(strPriv, ";ȫԺҽ���´�;") > 0 Then
                bln�´� = True '��ȫԺ����ҽ���´�Ȩ��
            ElseIf InStr(strPriv, ";����ҽ���´�;") > 0 _
                And InStr("," & mstr����IDs & ",", "," & mlng����ID & ",") > 0 Then
                bln�´� = True '�б��Ʋ���ҽ���´�Ȩ��
            End If
        ElseIf mint���� = 1 Then
            If mstr���λ�ʿ = UserInfo.���� Then
                bln�´� = True '��ǰ��ʿ���λ���
            ElseIf InStr(strPriv, ";ȫԺҽ���´�;") > 0 Then
                bln�´� = True '��ȫԺ����ҽ���´�Ȩ��
            ElseIf InStr(strPriv, ";����ҽ���´�;") > 0 _
                And InStr("," & mstr����IDs & ",", "," & mlng����ID & ",") > 0 Then
                bln�´� = True '�б��Ʋ���ҽ���´�Ȩ��
            End If
        Else
            bln�´� = True '���������ݲ�����
        End If
    Else
        bln�´� = True '��������
    End If
    PatiCanAdvice = bln�´�
End Function

Public Sub zlRefresh(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal lng����id As Long, _
    ByVal int״̬ As Integer, Optional ByVal blnMoved As Boolean, Optional ByVal lngǰ��ID As Long, _
    Optional ByVal intִ��״̬ As Integer, Optional ByVal lng�������ID As Long, Optional ByVal lng·��״̬ As Long = -1, _
    Optional ByVal lngҽ������ID As Long, Optional ByRef objMip As Object, Optional ByVal intӤ�� As Integer = -1, Optional ByVal lngǰ�����ID As Long, Optional ByVal lng����ҽ��ID As Long)
'���ܣ�ˢ��סԺҽ������
'������int����=���˵Ĳ�ͬ����
'      lngǰ��ID=����ҽ��վ����ʱ����
'      lng����ID��lng����ID=����5-���ת�Ʋ��ˡ�ʱΪ����ԭ������ԭ����
'      lng�������ID=�����ǰҽ��վ�ǻ��ﲡ�ˣ���Ϊ�������ID�������ҽ��վ���ã���Ϊҽ������ID
'      int״̬=����ҽ��վ����ʱ����,��Ŀ��ִ��״̬
'      lng·��״̬=-1:δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
'      blnMoved=�ò��˵������Ƿ���ת��
'      lngҽ������ID=ҽ��վ�������ID
'      lngǰ�����ID= lngǰ��ID����ҽ����Ӧ��ִ�п���ID����ҽ��վ���������������ʱ��lng�������ID<>lngǰ�����ID  lngǰ�����ID�������봫��
'      lng����ҽ��ID סԺҽ������վ�������ҽ��ʱ��ѡ�еĻ���ҽ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim objControl As CommandBarControl
    Dim lngPre����ID As Long
    Dim lngPre����ID As Long, lngPre����ID As Long
    Dim lngPre�������ID As Long
    Dim strPrivs As String
    
    lngPre����ID = mlng����ID
    lngPre����ID = mlng����ID
    lngPre�������ID = mlng�������ID
    lngPre����ID = mlng����ID
    mlng����ҽ��ID = lng����ҽ��ID
    
    mintPState = int״̬: mblnMoved = blnMoved
    mlng����ID = lng����ID: mlng��ҳID = lng��ҳID
    mlng����ID = lng����ID: mlng����ID = lng����id
    mlngǰ��ID = lngǰ��ID: mintִ��״̬ = intִ��״̬
    mlng�������ID = lng�������ID
    mlngҽ������ID = lngҽ������ID
    mlng·��״̬ = lng·��״̬
    mbytӤ�� = 0
    
    If InitObjPublicExpense Then
        Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng����ID, mlng��ҳID, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
    End If
    
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
    
    If Visible Or mblnInsideTools Then
        mblnSignVisible = True
        If mint���� = 0 Then
            If CheckSign(1, 0, mlng�������ID, mlng����ID, 2, False, gobjESign) = False Then
                mblnSignVisible = False '��ͬ����û������Ҫʹ��ǩ��
            End If
        ElseIf mint���� = 2 Then
            If CheckSign(3, 0, mlng�������ID, mlng����ID, 2, False, gobjESign) = False Then
                mblnSignVisible = False '��ͬ����û������Ҫʹ��ǩ��
            End If
        ElseIf mint���� = 1 Then
            If CheckSign(2, mlngҽ������ID, , , , False, gobjESign) = False Then
                mblnSignVisible = False '��ͬ����û������Ҫʹ��ǩ��
            End If
        End If
    End If
    
    '��ȡһЩ�������Ϣ
    If mlng����ID <> 0 And lngPre����ID <> mlng����ID Then
        On Error GoTo errH
        strSQL = "Select a.��Ժ����, a.סԺҽʦ, a.���λ�ʿ, a.����״̬, a.��������, a.����, a.Ӥ������id, a.Ӥ������id, a.סԺ��, b.����,b.��ǰ����,b.�Ա�,a.ҽ������ʱ�� as ����" & _
            " From ������ҳ A, ������Ϣ B Where a.����id = b.����id And a.����id = [1] And a.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mlng����ID, mlng��ҳID)
        mvInDate = rsTmp!��Ժ����
        mstrסԺҽ�� = NVL(rsTmp!סԺҽʦ)
        mstr���λ�ʿ = NVL(rsTmp!���λ�ʿ)
        mint����״̬ = NVL(rsTmp!����״̬, 0)
        mlng�������� = Val("" & rsTmp!��������)
        mint���� = Val("" & rsTmp!����)
        mlngӤ������ID = NVL(rsTmp!Ӥ������ID, 0)
        mlngӤ������ID = NVL(rsTmp!Ӥ������ID, 0)
        mstr���� = rsTmp!���� & ""
        mstrסԺ�� = rsTmp!סԺ�� & ""
        mstr���� = rsTmp!��ǰ���� & ""
        mstr�Ա� = rsTmp!�Ա� & ""
        mdat���� = NVL(rsTmp!����, CDate("1900-01-01"))
        
        '��ȡӤ����Ϣ
        mstrӤ�� = GetBabyRegList(lng����ID, lng��ҳID)
        If mstrӤ�� <> "" Then
            '��ȡ���ȱʡֵ��-1=����,0=����,1-Ӥ��1
            mvarCond.Ӥ�� = Val(zlDatabase.GetPara("����Ӥ������", glngSys, pסԺҽ���´�, "0"))
            If mvarCond.Ӥ�� > UBound(Split(mstrӤ��, "<Split>")) + 1 Then mvarCond.Ӥ�� = 0
            If mvarCond.Ӥ�� <> -1 Then mbytӤ�� = mvarCond.Ӥ��
        End If
        Call GetCriticalData
        On Error GoTo 0
    ElseIf mlng����ID = 0 Then
        mvInDate = CDate(0)
        mstrסԺҽ�� = ""
        mstr���λ�ʿ = ""
        mint����״̬ = 0
        mlng�������� = 0
        mstrӤ�� = ""
        mlngӤ������ID = 0
        mlngӤ������ID = 0
    End If
    
    If intӤ�� <> -1 And mstrӤ�� <> "" Then
        mbytӤ�� = intӤ��
        mvarCond.Ӥ�� = mbytӤ��
        mlngBaby = mbytӤ��
    End If
    
    If (lngPre�������ID <> mlng�������ID Or lngPre����ID <> mlng����ID) And mlngǰ��ID <> 0 Then
        mstrǰ��IDs = Getҽ������ҽ��IDs(mlng����ID, mlng��ҳID, IIF(0 = lngǰ�����ID, mlng�������ID, lngǰ�����ID), True, mlngǰ��ID)
    ElseIf mlngǰ��ID = 0 Then
        mstrǰ��IDs = ""
    End If
    
    If mstr����IDs = "" Then
        If mint���� = 0 Then
            mstr����IDs = GetUser����IDs(True)
        ElseIf mint���� = 1 Then
            mstr����IDs = GetUser����IDs
        End If
    End If
    'PASS ������ҩ��� ������Ϣ�����䶯
    If lngPre����ID <> mlng����ID Then
        If mblnPass Then
            Call zlPASSPati
            On Error Resume Next
            Call gobjPass.zlPassClearLight(mobjPassMap)
            On Error GoTo 0
        End If
    End If
    
    '�޸ķ��Ͳ˵�
    If mint���� = 1 And gstr��Һ�������� <> "" And lngPre����ID <> mlng����ID Then
        strPrivs = GetInsidePrivs(pסԺҽ������)
        If Not (InStr(";" & strPrivs & ";", ";����ҩ������;") = 0 Or InStr(";" & strPrivs & ";", ";����ҩ�Ƴ���;") = 0) Then
            Call SetSendCommandBar
        End If
    End If
    
    If Not grsTube Is Nothing Then
        If grsTube.State = 1 Then grsTube.Close
        Set grsTube = Nothing
    End If
    
    'ˢ������
    Call RefreshData
    
    'ִ���Զ�������ܣ�����ID=0Ҳ���ã���ʵ����رս���
    If mlngPlugInID <> 0 And lngPre����ID <> mlng����ID Then
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
'���ܣ���Ϣ����
'������objMip zl9ComLib.clsMipModule
    If Not (objMip Is Nothing) Then Set mclsMipModule = objMip
End Sub

Public Sub zlItemRef()
'���ܣ��������Ʋο�
    Dim lng������ĿID As Long, i As Long

    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) <> 0 Then
            If .TextMatrix(.Row, COL_�������) = "E" And (RowIs�䷽��(.Row) Or RowIs������(.Row)) Then
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), True)
            Else
                lng������ĿID = Get������ĿID(Val(.TextMatrix(.Row, COL_ID)), False)
            End If
        End If
    End With
    'ToDo:��ʾ���Ʋο�
    
End Sub

Public Function zlSeekAndViewEPRReport(ByVal lng����ID As Long) As Boolean
'���ܣ���λ�������Ӧ��ҽ�������򿪱���鿴
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, lngRow As Long
    
    strSQL = "Select ҽ��ID From ����ҽ������ Where ����ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng����ID)
    If Not rsTmp.EOF Then
        lngRow = vsAdvice.FindRow(CStr(rsTmp!ҽ��ID), , COL_ID)
        If lngRow <> -1 Then vsAdvice.Row = lngRow
        
        '��Ȩ�����д򿪣������Ƿ�λ��������״̬
        If InStr(GetInsidePrivs(pסԺҽ���´�), ";�������;") > 0 Then
            Select Case CheckEPRReport(rsTmp!ҽ��ID, lng����ID)
            Case 0
                MsgBox "��ҽ���ı���û����д��", vbInformation, gstrSysName
                Exit Function
            Case 2
                If InStr(GetInsidePrivs(pסԺҽ���´�), "����δ��ɱ���") > 0 Then
                    MsgBox "ע�⣺��ҽ���ı��滹û����ʽǩ����", vbInformation, gstrSysName
                Else
                    MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û��Ȩ�޲�����", vbInformation, gstrSysName
                    Exit Function
                End If
            End Select
            
            RaiseEvent ViewEPRReport(lng����ID, False)
        End If
        
        zlSeekAndViewEPRReport = True
    Else
        MsgBox "û���ҵ������Ӧ��ҽ����¼��", vbInformation, gstrSysName
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingAudit() As Boolean
'���ܣ��˶�
    Dim bln��ѪƤ�� As Boolean
    Dim strSQL As String
    Dim str�˶��� As String
    Dim bln��Ѫ As Boolean  '��Ѫҽ������Ѫ;��
    Dim lngRow As Long
    Dim strCheckTime As String
    Dim blnDo As Boolean
    Dim strXML As String
    
    With vsAppend
        bln��Ѫ = (.TextMatrix(.Row, COLSend("�������")) = "K" Or .TextMatrix(.Row, COLSend("�������")) = "E" And .TextMatrix(.Row, COLSend("��������")) = "8") And Mid(gstrҽ���˶�, 1, 1) = "1"
        bln��ѪƤ�� = (bln��Ѫ Or .TextMatrix(.Row, COLSend("�������")) = "E" And .TextMatrix(.Row, COLSend("��������")) = "1" And Mid(gstrҽ���˶�, 2, 1) = "1")
        If Not bln��ѪƤ�� Then
            If Val(gstrҽ���˶�) = 1 Then
                MsgBox "ֻ�ܺ˶�Ƥ��ҽ����", vbInformation, gstrSysName
            ElseIf Val(gstrҽ���˶�) = 10 Then
                MsgBox "ֻ�ܺ˶���Ѫҽ����", vbInformation, gstrSysName
            Else
                MsgBox "ֻ�ܺ˶���Ѫ����Ƥ��ҽ����", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.Row, COLExec("�˶���")) <> "" Then
            MsgBox "��ҽ���Ѿ��˶ԣ������ٴκ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.Row, vsExec.FixedCols) = "" Then
            MsgBox "��ҽ����δ����ִ������Ǽǣ����ܺ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        str�˶��� = zlDatabase.UserIdentifyByUser(Me, "�ں˶�ִ�����ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "ִ������Ǽ�", , True)
        If str�˶��� = "" Then Exit Function
        
        If str�˶��� = vsExec.TextMatrix(vsExec.Row, COLExec("ִ����")) Then
            MsgBox "ִ���˲��ܺ��������ͬ�����ܺ˶ԡ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '��ȡ�˶�ʱ��
        strCheckTime = frmAdviceStopTime.ShowMe(mfrmParent, Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("ҽ��ID"))), mlng����ID, 1, Format(vsExec.TextMatrix(vsExec.Row, COLExec("�Ǽ�ʱ��")), "yyyy-MM-dd HH:mm"))
        
        If Not IsDate(strCheckTime) Then
            Exit Function
        End If
        
    End With
    With vsExec
        On Error GoTo errH
        lngRow = vsExec.Row
        
        '���ú˶�ǰ��ҽӿ�
        If Not gobjPlugIn Is Nothing Then
            On Error Resume Next
            blnDo = gobjPlugIn.AdvcieBeforToReview(glngSys, IIF(mint���� = 0, pסԺҽ��վ, pסԺ��ʿվ), mlng����ID, mlng��ҳID, Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("ҽ��ID"))), Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("���ͺ�"))), str�˶���, strCheckTime, vsExec.TextMatrix(vsExec.Row, COLExec("ִ����")) & "", strXML)
            Call zlPlugInErrH(err, "AdvcieBeforToReview")
            If 0 = err.Number Then '�ӿ�û�г������������жϽӿڵķ���ֵ
                If blnDo Then
                    strSQL = "Zl_����ҽ���˶�_Insert(" & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("ҽ��ID"))) & "," & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("���ͺ�"))) & ",'" & str�˶��� & "'" & _
                    IIF(bln��Ѫ, ",To_Date('" & Format(vsExec.Cell(flexcpData, vsExec.Row, COLExec("ִ��ʱ��")), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", ",Null") & _
                    ",To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'))"
                Else
                    Exit Function
                End If
            End If
            If err.Number <> 0 Then err.Clear: Exit Function
            On Error GoTo 0
        Else
            strSQL = "Zl_����ҽ���˶�_Insert(" & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("ҽ��ID"))) & "," & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("���ͺ�"))) & ",'" & str�˶��� & "'" & _
            IIF(bln��Ѫ, ",To_Date('" & Format(vsExec.Cell(flexcpData, vsExec.Row, COLExec("ִ��ʱ��")), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", ",Null") & _
            ",To_Date('" & strCheckTime & "','YYYY-MM-DD HH24:MI:SS'))"
        End If
        
        If strSQL <> "" Then Call zlDatabase.ExecuteProcedure(strSQL, "ҽ���˶�")
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) 'Ҫ����ִ��״̬'����Ҫ����ִ��״̬
        vsExec.Row = lngRow
        FuncThingAudit = True
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function FuncThingDelAudit() As Boolean
'���ܣ�ȡ���˶�
    Dim bln��ѪƤ�� As Boolean
    Dim strSQL As String
    Dim str�˶��� As String
    Dim bln��Ѫ As Boolean '��Ѫҽ������Ѫ;��
    Dim lngRow As Long
    
    With vsAppend
        bln��Ѫ = (.TextMatrix(.Row, COLSend("�������")) = "K" Or .TextMatrix(.Row, COLSend("�������")) = "E" And .TextMatrix(.Row, COLSend("��������")) = "8") And Mid(gstrҽ���˶�, 1, 1) = "1"
        bln��ѪƤ�� = (bln��Ѫ Or .TextMatrix(.Row, COLSend("�������")) = "E" And .TextMatrix(.Row, COLSend("��������")) = "1" And Mid(gstrҽ���˶�, 2, 1) = "1")
        If Not bln��ѪƤ�� Then
            If Val(gstrҽ���˶�) = 1 Then
                MsgBox "ֻ��ȡ���˶�Ƥ��ҽ����", vbInformation, gstrSysName
            ElseIf Val(gstrҽ���˶�) = 10 Then
                MsgBox "ֻ��ȡ���˶���Ѫҽ����", vbInformation, gstrSysName
            Else
                MsgBox "ֻ��ȡ���˶���Ѫ����Ƥ��ҽ����", vbInformation, gstrSysName
            End If
            Exit Function
        End If
        If vsExec.TextMatrix(vsExec.Row, COLExec("�˶���")) = "" Then
            MsgBox "��ҽ����δ���к˶ԣ�����ȡ����", vbInformation, gstrSysName
            Exit Function
        End If
        

    End With
    With vsExec
        If vsExec.TextMatrix(vsExec.Row, COLExec("�˶���")) <> UserInfo.���� Then
            str�˶��� = zlDatabase.UserIdentifyByUser(Me, "��ȡ���˶�ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "ִ������Ǽ�", , True)
            If str�˶��� = "" Then Exit Function
            If str�˶��� <> vsExec.TextMatrix(vsExec.Row, COLExec("�˶���")) Then
                MsgBox "ֻ��ȡ���Լ��˶Ե�ҽ������ǰҽ���˶�����""" & vsExec.TextMatrix(vsExec.Row, COLExec("�˶���")) & """", vbInformation, gstrSysName
                Exit Function
            End If
        Else
            If MsgBox("��ȷ��Ҫȡ���˶���", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then Exit Function
        End If
        On Error GoTo errH
        lngRow = vsExec.Row
        strSQL = "Zl_����ҽ���˶�_Delete(" & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("ҽ��ID"))) & "," & Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("���ͺ�"))) & _
            IIF(bln��Ѫ, ",To_Date('" & Format(vsExec.Cell(flexcpData, vsExec.Row, COLExec("ִ��ʱ��")), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))", ")")
        Call zlDatabase.ExecuteProcedure(strSQL, "ȡ��ҽ���˶�")
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) 'Ҫ����ִ��״̬'����Ҫ����ִ��״̬
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
    Case "����"
        mvarCond.��ʼʱ�� = CDate(0)
        mvarCond.����ʱ�� = CDate(0)
    Case "����"
        mvarCond.��ʼʱ�� = Format(curDate, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "����"
        mvarCond.��ʼʱ�� = Format(curDate - 1, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate - 1, "yyyy-MM-dd 23:59:59")
    Case "�������"
        mvarCond.��ʼʱ�� = Format(curDate - 2, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "���һ��"
        mvarCond.��ʼʱ�� = Format(curDate - 7, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "�������"
        mvarCond.��ʼʱ�� = Format(curDate - 14, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "���һ��"
        mvarCond.��ʼʱ�� = Format(curDate - 30, "yyyy-MM-dd 00:00:00")
        mvarCond.����ʱ�� = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case "[ָ��..]"
        If Not frmSelectTime.ShowMe(Me, mvarCond.��ʼʱ��, mvarCond.����ʱ��, cboTime, 1) Then
            'ȡ��ʱ�ָ�ԭ����ѡ��
            Call zlControl.CboSetIndex(cboTime.hwnd, mintPreTime)
            If vsAdvice.Enabled Then vsAdvice.SetFocus
            Exit Sub
        Else
            If vsAdvice.Enabled Then vsAdvice.SetFocus
        End If
    End Select
        
    If mvarCond.��ʼʱ�� = CDate(0) Or mvarCond.����ʱ�� = CDate(0) Then
        cboTime.ToolTipText = ""
    Else
        cboTime.ToolTipText = "��Χ��" & Format(mvarCond.��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & " �� " & Format(mvarCond.����ʱ��, "yyyy-MM-dd HH:mm:ss")
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
        Case ID_��ʾִ��
            mblnShowExec = Not mblnShowExec
            Call SetExecShow(True, mblnShowExec)
            Call vsAppend_AfterRowColChange(-1, -1, vsAppend.Row, vsAppend.Col)
        Case ID_���ִ��
            Call FuncExecFinish
        Case ID_ȡ�����
            Call FuncExecCancel
        Case ID_ִ�м�¼
            Call FuncThingNew
        Case ID_ִ�е���
            Call FuncThingModi
        Case ID_ִ��ɾ��
            Call FuncThingDel
        Case ID_�˶�
            Call FuncThingAudit
        Case ID_ȡ���˶�
            Call FuncThingDelAudit
    End Select
End Sub

Private Sub cbsExec_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intִ��״̬ As Integer, blnSelect As Boolean
    Dim str����� As String, intȡ��ִ����� As Integer
    
    If Not tbcAppend.Selected.Tag = "����" Or Not picExec.Visible Then Exit Sub
    
    With vsAppend
        blnSelect = Val(.TextMatrix(.Row, COLSend("ҽ��ID"))) <> 0
        If blnSelect Then '0-δִ��,1-��ִ��,2-�ܾ�ִ��,3-����ִ��
            intִ��״̬ = Val(.Cell(flexcpData, .Row, COLSend("ִ��״̬")))
            str����� = .TextMatrix(.Row, COLSend("ִ����"))
        End If
    End With
    
    Select Case Control.ID
        Case ID_��ʾִ��
            Control.Checked = mblnShowExec
        Case ID_���ִ��
            If InStr(GetInsidePrivs(pסԺҽ������), "ȷ��ִ�����") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = blnSelect And (intִ��״̬ = 0 Or intִ��״̬ = 3)
            End If
        Case ID_ȡ�����
            intȡ��ִ����� = IIF(InStr(GetInsidePrivs(pסԺҽ������), "ȡ��ִ�����") = 0, 0, 1) + IIF(InStr(GetInsidePrivs(pסԺҽ������), "ȡ������ִ�����") = 0, 0, 2)
            
            If intȡ��ִ����� = 0 Then
                Control.Visible = False
            ElseIf intȡ��ִ����� = 1 Then
                Control.Enabled = blnSelect And intִ��״̬ = 1 And str����� = UserInfo.����
            ElseIf intȡ��ִ����� = 2 Then
                Control.Enabled = blnSelect And intִ��״̬ = 1 And str����� <> UserInfo.����
            ElseIf intȡ��ִ����� = 3 Then
                 Control.Enabled = blnSelect And intִ��״̬ = 1
            End If
        Case ID_ִ�м�¼
            If InStr(GetInsidePrivs(pסԺҽ������), "ִ������Ǽ�") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnShowExec And blnSelect And (intִ��״̬ = 0 Or intִ��״̬ = 3)
            End If
        Case ID_ִ�е���, ID_ִ��ɾ��
            If InStr(GetInsidePrivs(pסԺҽ������), "ִ������Ǽ�") = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnShowExec And blnSelect And (intִ��״̬ = 0 Or intִ��״̬ = 3) _
                    And vsExec.TextMatrix(vsExec.Row, vsExec.FixedCols) <> "" And vsExec.Row = vsExec.FixedRows
            End If
        Case ID_�˶�, ID_ȡ���˶�
            If InStr(GetInsidePrivs(pסԺҽ������), "ִ������Ǽ�") = 0 Or Val(gstrҽ���˶�) = 0 Then
                Control.Visible = False
            Else
                Control.Enabled = mblnShowExec And blnSelect And (intִ��״̬ = 0 Or intִ��״̬ = 3)
                If mblnShowExec And (intִ��״̬ = 0 Or intִ��״̬ = 3) Then
                    If vsExec.TextMatrix(vsExec.Row, COLExec("�˶���")) = "" Then
                        If Control.ID = ID_�˶� Then Control.Enabled = True
                        If Control.ID = ID_ȡ���˶� Then Control.Enabled = False
                    Else
                        If Control.ID = ID_�˶� Then Control.Enabled = False
                        If Control.ID = ID_ȡ���˶� Then Control.Enabled = True
                    End If
                End If
            End If
    End Select
End Sub

Private Sub cbsSub_ControlSelected(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Control Is Nothing Then
        Select Case Control.ID
            Case ID_ҽ����ɫʾ��
                vsfAdviceColor.Visible = True
                If vsfAdviceColor.Row <= 0 Then
                    With vsfAdviceColor
                        .ColWidth(0) = 3400
                        .Width = 3400
                        .Height = 3300
                        .TextMatrix(0, 0) = "�¿�"
                        .Cell(flexcpForeColor, 0, 0, 0, 0) = vbBlack
                        .RowHeight(0) = 300

                        .TextMatrix(1, 0) = "У������"
                        .Cell(flexcpForeColor, 1, 0, 1, 0) = &H80&
                        .RowHeight(1) = 300

                        .TextMatrix(2, 0) = "��У��/������/������"
                        .Cell(flexcpForeColor, 2, 0, 2, 0) = &HC00000
                        .RowHeight(2) = 300
                        
                        .TextMatrix(3, 0) = "��ֹͣ/��ȷ��ֹͣ/δ��ҽ��"
                        .Cell(flexcpForeColor, 3, 0, 3, 0) = &H808080
                        .RowHeight(3) = 300
                        
                        .TextMatrix(4, 0) = "����ͣ"
                        .Cell(flexcpForeColor, 4, 0, 4, 0) = &H8000&
                        .RowHeight(4) = 300
                        
                        .TextMatrix(5, 0) = "������"
                        .Cell(flexcpForeColor, 5, 0, 5, 0) = &H808080
                        .Cell(flexcpFontStrikethru, 5, 0, 5, 0) = True
                        .RowHeight(5) = 300

                        .TextMatrix(6, 0) = "ֹͣ����ͣ��ʱ��δ��"
                        .Cell(flexcpForeColor, 6, 0, 6, 0) = &HFF8080
                        .RowHeight(6) = 300

                        .TextMatrix(7, 0) = "���ú�ʱ��δ��"
                        .Cell(flexcpForeColor, 7, 0, 7, 0) = &H4AAD00
                        .RowHeight(7) = 300

                        .TextMatrix(8, 0) = "����ҽ��У�Ժ�ת��ҽ�����ͺ�"
                        .Cell(flexcpForeColor, 8, 0, 8, 0) = vbRed
                        .RowHeight(8) = 300

                        .TextMatrix(9, 0) = "(��ҽ��������)���龫������ҩƷ"
                        .Cell(flexcpFontBold, 9, 0, 9, 0) = True
                        .RowHeight(9) = 300
                        
                        .TextMatrix(10, 0) = "�����ѷ��͵�(�������ܷ��͵�����)"
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
    Case ID_Ӥ��
        strTmp = IIF(mvarCond.����ģʽ = 3, "����", "ҽ��")
        arrBaby = Split(mstrӤ��, "<Split>")
        With CommandBar.Controls
            .DeleteAll
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100#, "����" & strTmp)
            Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + 1, "����" & strTmp): objControl.BeginGroup = True
            For i = 0 To UBound(arrBaby)
                Set objControl = .Add(xtpControlButton, ID_Ӥ�� * 100# + i + 2, "Ӥ�� " & i + 1 & IIF(arrBaby(i) <> "", "��" & arrBaby(i), ""))
                If i = 0 Then objControl.BeginGroup = True
            Next
        End With
    Case Else
        Call zlPopupCommandBars(CommandBar)
    End Select
End Sub

Private Sub DkpBlood_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If picBlood.Tag <> "�ɼ�" Then Exit Sub
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
        '��Ѫִ��
        picBlood.Top = picBlood.Top + Y
        picBlood.Height = picBlood.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub fraHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    timHide.Enabled = True
End Sub

Private Sub mfrmCompoundMedicine_SetEditState(ByVal blnEditState As Boolean)
'���ܣ����ݵ�ǰ�Ƿ��޸�״̬�������Ƿ��ת�ƽ���
    RaiseEvent SetEditState(blnEditState)
    mblnEditState = blnEditState
    vsAdvice.Enabled = Not blnEditState
End Sub

Private Sub mfrmCompoundMedicine_StatusTextUpdate(ByVal Text As String)
    RaiseEvent StatusTextUpdate(Text)
End Sub

Private Sub mfrmEdit_EditDiagnose(ParentForm As Object, ByVal ����ID As Long, ByVal ��ҳID As Long, ByVal ����ID As Long, ByVal str���� As String, Succeed As Boolean)
    RaiseEvent EditDiagnose(ParentForm, ����ID, ��ҳID, ����ID, str����, Succeed)
End Sub

Private Sub mfrmEdit_FormUnload(Cancel As Integer)
    If mlngΣ��ֵID <> 0 Then
        Call GetCriticalData
    End If
    mlngΣ��ֵID = 0
    If Not Cancel Then
        If mfrmEdit.mblnOK Then Call LoadAdvice(True)
        Set mfrmEdit = Nothing
        
        If Me.Visible Then
            Call BringWindowToTop(Me.hwnd)
        End If
        
         '����·���嵥��ˢ�£�ҽ���¿��������ɾ���������ȣ�
        If mlng·��״̬ = 1 And Not gobjPath Is Nothing Then
            If GetInsidePrivs(p�ٴ�·��Ӧ��) <> "" Then
                Call gobjPath.zlRefresh(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, mintPState, mblnMoved, True)
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
    
    '���뵥�ݴ�ӡ֮��Ĵ���
    If mstrBillPrint <> "" Then
        If Split(mstrBillPrint, ",")(0) = ReportNum Then
            strSQL = "Zl_���Ƶ��ݴ�ӡ_Insert('" & Split(mstrBillPrint, ",")(1) & "'," & Val(Split(mstrBillPrint, ",")(2)) & ",1,'" & UserInfo.���� & "')"
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

    If picAppend.Tag = "��ִ��" Then Exit Sub
    
    vsAppend.Left = 0
    vsAppend.Top = 0
    vsAppend.Width = picAppend.Width
        
     '��Ѫִ�еǼǺ�ҽ��ִ�еǼ�ʼ��ֻ����ʾһ�����򶼲���ʾ
    If picBlood.Tag = "�ɼ�" Then
        vsAppend.Height = picAppend.Height - picBlood.Height - IIF(DkpBlood.Tag = "�ɼ�", fraExecUD.Height, 0)
        vsAppend.TopRow = vsAppend.Row
                
        fraExecUD.Left = 0
        fraExecUD.Width = picAppend.Width
        fraExecUD.Top = vsAppend.Top + vsAppend.Height
        
        '��Ѫִ�����
        With picBlood
            .Left = 0
            If DkpBlood.Tag = "�ɼ�" Then
                .Top = fraExecUD.Top + fraExecUD.Height
            Else
                .Top = vsAppend.Top + vsAppend.Height
            End If
            .Width = picAppend.Width
        End With
    Else
        vsAppend.Height = picAppend.Height - IIF(picExec.Tag = "�ɼ�", picExec.Height, 0) - IIF(vsExec.Tag = "�ɼ�", fraExecUD.Height + vsExec.Height, 0)
    
        vsAppend.TopRow = vsAppend.Row
                
        fraExecUD.Left = 0
        fraExecUD.Width = picAppend.Width
        fraExecUD.Top = vsAppend.Top + vsAppend.Height
        
        picExec.Left = 0
        picExec.Width = picAppend.Width
        If vsExec.Tag = "�ɼ�" Then
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
    
    '��ѡ����
    With vsAdvice
        fraColSel.Left = .Left + (.ColWidth(COL_F��־) + .ColWidth(COL_F����) - fraColSel.Width) / 2 + 30
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
    blnTmp = mbln����
    mbln���� = False
    Select Case Item.Tag
    Case "����������"
        mvarCond.����ģʽ = 0
    Case "����"
        mvarCond.����ģʽ = 1
    Case "����"
        mvarCond.����ģʽ = 2
    Case "����"
        mvarCond.����ģʽ = 3
        mbln���� = True
    End Select
    
    If Item.Tag <> "" And mlng����ID <> 0 Then
        If blnTmp <> mbln���� Then
            Call AddToolBarInDoctor
            Call DefInSidePlugInBar(mrsPlugInBar)
            cbsSub.RecalcLayout
        End If
        Call RefreshData
    End If
End Sub

Private Sub timBRefresh_Timer()
    '��Ѫ����Ѫִ�д�����д��ִ�����ݺ�ҽ����Ӧ���ݵ�ˢ��
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
                Case 1, 2 '��¼ִ�л����ִ�У�ɾ��ִ��
                    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) 'Ҫ����ִ��״̬'����Ҫ����ִ��״̬
                Case 3, 4 'ִ�����,ȡ�����
                    Call LoadAdvice
                Case 5, 6 'ִ�к˶�,ȡ���˶�
                    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) 'Ҫ����ִ��״̬'����Ҫ����ִ��״̬
            End Select
            mobjFrmBlood.ExecFresh = False
            mobjFrmBlood.AdviceExecState = 0
        End If
    End If
End Sub

Private Sub timHide_Timer()
'���ܣ�������˹��������Զ���ʾ������
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
    Dim bln���� As Boolean
    
    If Control.ID <> 0 Then
        If cbsSub.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If

    Select Case Control.ID
        Case ID_����ҽ��, ID_����ҽ��
            mvarCond.ҽ����ʾ = IIF(mvarCond.ҽ����ʾ = 1, 0, 1)
        Case ID_ȫ��
            mvarCond.���� = 0
        Case ID_���
            mvarCond.���� = 1
        Case ID_����
            mvarCond.���� = 2
        Case ID_����
            mvarCond.���� = 3
        Case ID_δ������
            If mvarCond.δ������ Then
                If mvarCond.�ѳ����� Then
                    mvarCond.δ������ = Not mvarCond.δ������
                End If
            Else
                mvarCond.δ������ = Not mvarCond.δ������
            End If
        Case ID_�ѳ�����
            If mvarCond.�ѳ����� Then
                If mvarCond.δ������ Then
                    mvarCond.�ѳ����� = Not mvarCond.�ѳ�����
                End If
            Else
                mvarCond.�ѳ����� = Not mvarCond.�ѳ�����
            End If
        Case ID_Ӥ�� * 100# '����ҽ��
            If mvarCond.Ӥ�� = -1 Then Exit Sub
            mvarCond.Ӥ�� = -1
            mbytӤ�� = 0
            Call zlDatabase.SetPara("����Ӥ������", mvarCond.Ӥ��, glngSys, pסԺҽ���´�)
        Case ID_Ӥ�� * 100# + 1 To ID_Ӥ�� * 100# + 6 '���ˡ�Ӥ��ҽ��
            If mvarCond.Ӥ�� = Control.ID - ID_Ӥ�� * 100# - 1 Then Exit Sub
            mvarCond.Ӥ�� = Control.ID - ID_Ӥ�� * 100# - 1
            mbytӤ�� = mvarCond.Ӥ��
            Call zlDatabase.SetPara("����Ӥ������", mvarCond.Ӥ��, glngSys, pסԺҽ���´�)
        Case ID_����
            mvarCond.���� = Not mvarCond.����
        Case ID_δ����ֹʱ��
            mvarCond.δ����ֹʱ�� = Not mvarCond.δ����ֹʱ��
        Case ID_δ����
            mvarCond.δ���� = Not mvarCond.δ����
        Case ID_����
            mvarCond.���� = Not mvarCond.����
        Case ID_�Ǳ���ҽ��
            If mvarCond.�Ǳ���ҽ�� Then
                If mvarCond.�Ǳ���ҽ�� Then
                    mvarCond.�Ǳ���ҽ�� = Not mvarCond.�Ǳ���ҽ��
                End If
            Else
                mvarCond.�Ǳ���ҽ�� = Not mvarCond.�Ǳ���ҽ��
            End If
        Case ID_�Ǳ���ҽ��
            If mvarCond.�Ǳ���ҽ�� Then
                If mvarCond.�Ǳ���ҽ�� Then
                    mvarCond.�Ǳ���ҽ�� = Not mvarCond.�Ǳ���ҽ��
                End If
            Else
                mvarCond.�Ǳ���ҽ�� = Not mvarCond.�Ǳ���ҽ��
            End If
        Case ID_���
            mvarCond.��ʾģʽ = 0
        Case ID_��ϸ
            mvarCond.��ʾģʽ = 1
        Case Else
            Call zlExecuteCommandBars(Control)
            blnUnrefresh = True
    End Select
    
    bln���� = InStr("," & ID_δ������ & "," & "," & ID_�ѳ����� & "," & "," & ID_ȫ�� & "," & ID_��� & "," & ID_���� & "," & ID_���� & ",", "," & Control.ID & ",") > 0
    
    If Not blnUnrefresh Then cbsSub.RecalcLayout
    
    If Not bln���� And blnUnrefresh = False Then
        Call RefreshData
    ElseIf bln���� Then
        Call Refresh����
    End If
End Sub

Private Sub cbsSub_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not Between(Control.ID, conMenu_Edit_Untread * 100# + 1, conMenu_Edit_Untread * 100# + 99) Then Control.Enabled = mlng����ID <> 0 And mblnEditState = False
    If Not Control.Enabled Then Exit Sub
    Select Case Control.ID
        Case ID_ʱ��, ID_ʱ���ǩ
            If mvarCond.����ģʽ <> 3 And mvarCond.ҽ����ʾ = 0 Then
                Control.Visible = False
            Else
                Control.Visible = True
            End If
        Case ID_δ����ֹʱ��
            Control.Checked = mvarCond.δ����ֹʱ��
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            If mvarCond.ҽ����ʾ = 0 And (mvarCond.����ģʽ = 0 Or mvarCond.����ģʽ = 1) Then
                Control.Visible = mvarCond.����ģʽ <> 3
            Else
                Control.Visible = False
            End If
        Case ID_����ҽ��
            Control.Checked = mvarCond.ҽ����ʾ = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ <> 3
        Case ID_����ҽ��
            Control.Checked = mvarCond.ҽ����ʾ = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ <> 3
        Case ID_ȫ��
            Control.Checked = mvarCond.���� = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_���
            Control.Checked = mvarCond.���� = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_����
            Control.Checked = mvarCond.���� = 2
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_����
            Control.Checked = mvarCond.���� = 3
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_δ������
            Control.Checked = mvarCond.δ������
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_�ѳ�����
            Control.Checked = mvarCond.�ѳ�����
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.����ģʽ = 3
        Case ID_Ӥ�� 'Ӥ��ҽ������
            If mstrӤ�� <> "" Then
                Control.Visible = False
                If mvarCond.Ӥ�� = -1 Then
                    Control.Caption = IIF(mvarCond.����ģʽ = 3, "���б���", "����ҽ��")
                ElseIf mvarCond.Ӥ�� = 0 Then
                    Control.Caption = IIF(mvarCond.����ģʽ = 3, "���˱���", "����ҽ��")
                Else
                    Control.Caption = "Ӥ�� " & mvarCond.Ӥ��
                End If
                Control.Visible = True
            Else
                If mvarCond.Ӥ�� <> -1 Or Control.Visible Then
                    mvarCond.Ӥ�� = -1
                    mbytӤ�� = 0
                    Control.Visible = False
                    Call zlDatabase.SetPara("����Ӥ������", mvarCond.Ӥ��, glngSys, pסԺҽ���´�)
                End If
            End If
        Case ID_Ӥ�� * 100# '����ҽ��
            Control.Checked = mvarCond.Ӥ�� = -1
        Case ID_Ӥ�� * 100# + 1 To ID_Ӥ�� * 100# + 6 '���ˡ�Ӥ��ҽ��
            Control.Checked = mvarCond.Ӥ�� = Control.ID - ID_Ӥ�� * 100# - 1
        Case ID_����
            Control.Checked = mvarCond.����
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
        Case ID_δ����
            If mint���� <> 1 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Checked = mvarCond.δ����
                Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            End If
        Case ID_����
            If mint���� <> 2 Then
                Control.Visible = False
            Else
                Control.Visible = True
                Control.Checked = mvarCond.����
                Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            End If
        Case ID_�Ǳ���ҽ��
            Control.Checked = mvarCond.�Ǳ���ҽ��
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.����ģʽ <> 3
        Case ID_�Ǳ���ҽ��
            Control.Checked = mvarCond.�Ǳ���ҽ��
            Control.IconId = IIF(Not Control.Checked, 90001, 90002)
            Control.Visible = mvarCond.����ģʽ <> 3
        Case ID_���
            Control.Checked = mvarCond.��ʾģʽ = 0
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ <> 3
        Case ID_��ϸ
            Control.Checked = mvarCond.��ʾģʽ = 1
            Control.IconId = IIF(Control.Checked, 90004, 90003)
            Control.Visible = mvarCond.����ģʽ <> 3
        Case Else
            Call zlUpdateCommandBars(Control)
    End Select
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If vsColumn.Visible Then
        vsColumn.SetFocus '��ѡ����
    Else
        If Me.Visible And vsAdvice.Enabled Then vsAdvice.SetFocus
    End If
    RaiseEvent Activate
End Sub

Private Sub Form_Deactivate()
    vsColumn.Visible = False '��ѡ����
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call ActiveHotKey(KeyCode, Shift)
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
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
    'ˢ��ҽ��������ϸ
    Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    Set mfrmBilling = Nothing
End Sub

Private Function CheckWindow() As Boolean
'���ܣ����ҽ���༭�����Ƿ��Ѿ���
    If Not mfrmEdit Is Nothing Then
        '��ǰ���ڴ���
        MsgBox "ҽ���༭�����Ѿ��򿪣�������ɵ�ǰ��������ִ�С�", vbInformation, gstrSysName
        '��λ����ǰ�Ĵ���
        If mfrmEdit.WindowState = vbMinimized Then mfrmEdit.WindowState = vbNormal
        If mfrmEdit.Visible Then mfrmEdit.SetFocus
        Exit Function
    Else
        '�������ڴ���
        If Not CheckAdviceWindow("סԺҽ���༭") Then Exit Function
    End If
 
    '���������뵥�����Ƿ��Ѿ���
    If Not mfrmEac Is Nothing Then
        '��ǰ���ڴ���
        MsgBox "�������봰���Ѿ��򿪣�������ɵ�ǰ��������ִ�С�", vbInformation, gstrSysName
        '��λ����ǰ�Ĵ���
        If mfrmEac.WindowState = vbMinimized Then mfrmEac.WindowState = vbNormal
        If mfrmEac.Visible Then mfrmEac.SetFocus
        Exit Function
    End If
 
    CheckWindow = True
End Function

Private Sub FuncBillPrint(objControl As CommandBarControl)
'���ܣ���ӡ���Ƶ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strNO As String, lng��¼���� As Long
    Dim strParameter As String
    Dim lng���ID As Long
    Dim strErr As String
    Dim blnDo As Boolean
    Dim strBillName As String '���Ƶ��ݵ�����  �����ļ��б�.����
    
    If objControl.Parameter = "" Then Exit Sub
    strParameter = objControl.Parameter
    If InStr(objControl.Parameter, "|") > 0 Then strParameter = Split(objControl.Parameter, "|")(0): strNO = Split(objControl.Parameter, "|")(1)
    
    strBillName = objControl.Caption
    strBillName = Replace("<Tab>" & strBillName, "<Tab>��ӡ:", "")
    If InStr(strBillName, "(&") > 0 Then
        strBillName = Mid(strBillName, 1, InStr(strBillName, "(&") - 1)
    End If
    
    '��Ժ���˲������ӡ
    If mintPState = ps��Ժ Then
        MsgBox "�ò����Ѿ���Ժ,���ܴ�ӡ:" & strBillName & "��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsAdvice
        '��ӡ������ʾ
        On Error GoTo errH
        lng���ID = Decode(Val(.TextMatrix(.Row, COL_���ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_���ID)))
        If .TextMatrix(.Row, COL_�������) = "E" And Val(.TextMatrix(.Row, COL_��������)) = 6 Then
            If Not gobjLIS Is Nothing Then    '��ӡ�������뵥��
                 blnDo = gobjLIS.CheckAcceptance(CStr(lng���ID), strErr)
                 If Not blnDo Then
                    MsgBox "�ñ걾�Ѿ�������ƺ��գ����ܴ�ӡ:" & strBillName & "��", vbInformation, gstrSysName
                    Exit Sub
                 End If
            End If
        End If
        If mintBillPrint = 0 Then
            If strNO <> "" Then
                strSQL = "Select A.NO,A.��¼���� from ����ҽ������ A,����ҽ����¼ B Where a.ҽ��ID=b.id And a.NO=[2] And (b.ID=[3] Or b.���ID=[3])"
            Else
                strSQL = "Select NO,��¼���� from ����ҽ������ Where ҽ��ID=[1] order By ����ʱ�� Desc"
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, COL_ID)), strNO, lng���ID)
            If rsTmp.RecordCount > 0 Then
                strNO = rsTmp!NO & ""
                lng��¼���� = Val(rsTmp!��¼���� & "")
            End If
        Else
            strNO = vsAppend.TextMatrix(vsAppend.Row, COLSend("���ݺ�"))
            lng��¼���� = Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("��¼����")))
        End If
        strSQL = "Select ��ӡ��,��ӡʱ�� From ���Ƶ��ݴ�ӡ Where NO=[1] And ��¼����=[2] And ��ӡ����=1 Order by ��ӡʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strNO, lng��¼����)
        If Not rsTmp.EOF Then
            If MsgBox("��[" & strBillName & "]�Ѿ���ӡ�� " & rsTmp.RecordCount & " �Σ����һ����""" & _
                rsTmp!��ӡ�� & """��""" & Format(rsTmp!��ӡʱ��, "yyyy-MM-dd HH:mm") & """��ӡ��" & vbCrLf & vbCrLf & "Ҫ������ӡ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
         '��Ѫҽ����ӡ���뵥������غ������м��
        If InStr(1, ",ZL1_INSIDE_1254_17_1,ZL1_INSIDE_1254_17_2,", "," & strParameter & ",") <> 0 Then
            If BloodApplyPrintCheck(Val(.TextMatrix(.Row, COL_ID)), 2, IIF(strParameter = "ZL1_INSIDE_1254_17_1", 1, 2), 1) = False Then Exit Sub
        End If
        On Error GoTo 0
        
        '���ô�ӡ
        If mobjReport.ReportPrintSet(gcnOracle, glngSys, strParameter, mfrmParent) Then
            mstrBillPrint = strParameter & "," & strNO & "," & lng��¼����
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strParameter, mfrmParent, "NO=" & strNO, "����=" & lng��¼����, "ҽ��ID=" & lng���ID, 2)
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
        MsgBox "���˵ı���סԺ�����Ѿ�ת���������ݿ⣬�����������" & vbCrLf & _
            "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
        CheckDataMoved = True
    End If
End Function

Private Sub FuncAdviceAdd()
'���ܣ�����ҽ��
    Dim datTurn As Date, intӤ�� As Integer

    On Error GoTo errH
    
    If Not CheckWindow Then Exit Sub
    If CheckAdviceAddModi(0, 0, datTurn) = False Then Exit Sub
    
    If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
        If CheckPatiTurnLimit(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, datTurn, mintPState) = False Then Exit Sub
    End If

    If Not FuncPathAdd() Then Exit Sub
    '-1��ʾ���˺�Ӥ��
    If mvarCond.Ӥ�� >= 0 Then intӤ�� = mvarCond.Ӥ��
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mint����, mMainPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, _
            intӤ��, , mblnModalNew, mlng�������ID, , , mintPState, mlng����ID, mlng����ID, datTurn, mlngҽ������ID, mstrǰ��IDs, mclsMipModule, mlngΣ��ֵID, mlng����ҽ��ID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceConfirm(ByVal blnOnePati As Boolean, ByVal Control As XtremeCommandBars.ICommandBarControl)
'���ܣ�ȷ��ֹͣҽ��
    Dim lngҽ��ID As Long
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 2, mlng����ID, mlng��ҳID, mlng����ID, lngҽ��ID, mint���� = 1, , , , , mblnDirect And Not mblnBatch Or mblnInsideTools Or blnOnePati, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdviceAudit()
'���ܣ����ҽ��
    Dim datTurn As Date
        
    If Not CheckWindow Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    If Not mblnHaveAuditPriv Then
        MsgBox "�㲻�������ҽ�����ʸ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If CheckDataMoved Then Exit Sub
    
    '���ʱ������ʱ�ޣ���Ϊ�¿����޸�ʱ����û��ʱ�ޣ������ʱ���ˣ��ᵼ�²�����ҽ���޷���������
'    If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
'        If CheckPatiTurnLimit(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, datTurn, mintPState) = False Then Exit Sub
'    End If
    
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mint����, mMainPrivs, mlng����ID, mlng��ҳID, _
            mlngǰ��ID, , , , mlng�������ID, True, , mintPState, mlng����ID, mlng����ID, datTurn, mlngҽ������ID, mstrǰ��IDs, mclsMipModule)
End Sub

Private Sub FuncAdviceDel()
'ɾ����ɾ����ǰҽ��
'˵������������ɾ��,�Լ�����,�������,��ҩ�䷽,������ɾ��,һ����ҩֻɾ����ǰҩƷ
    Dim strSQL As String, lngҽ��ID As Long
    Dim blnGroup As Boolean, i As Long, blnBat As Boolean, blnTrans As Boolean
    Dim lngRow As Long, arrSQL As Variant, lng������� As Long
    Dim strDelIDs As String, arrDelID() As String
    Dim strDelDrugIDs As String              '��¼ɾ����ҩƷҽ��,���ڴ��������ҩ���
    Dim lngBabyEdit As Long, int��Ч As Integer
    Dim strMsg As String
    Dim lng��ID As Long
    Dim blnRISԤԼ As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strWhere As String
    Dim bln��Ѫ As Boolean, strErr As String
    Dim bln����ҽ�� As Boolean
    
    With vsAdvice
        '����Ƿ����ɾ��
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ������ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub

        '��鲡���Ƿ��������
        If Not CheckPatiIsAduit Then Exit Sub
        lngBabyEdit = CheckBabyEdit(Val(.TextMatrix(.Row, COL_Ӥ��ID)))
        If lngBabyEdit = 1 Then
            MsgBox "��ǰ���˲��ڱ����ң�������ɾ������ҽ����", vbInformation, gstrSysName
            Exit Sub
        ElseIf lngBabyEdit = 2 Then
            MsgBox "��ǰ���˵�Ӥ�����ڱ����ң�������ɾ��Ӥ��ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        'ҽ����ҽ���´��ҽ�����ܻ�ɾ��ҽ��վֻ��ɾ�������´��ҽ��
        If mint���� = 2 Then
            If InStr("," & mstrǰ��IDs & ",", "," & .TextMatrix(.Row, COL_ǰ��ID) & ",") = 0 Then
                MsgBox "��ҽ����Ϊ��ǰҽ�������´����ɾ����ҽ����", vbInformation, gstrSysName
                Exit Sub
            ElseIf Val(.TextMatrix(.Row, COL_ǰ��ID)) = 0 Then
                MsgBox "��ҽ������ҽ�������´����ɾ����ҽ����", vbInformation, gstrSysName
                Exit Sub
            End If
        ElseIf Val(.TextMatrix(.Row, COL_ǰ��ID)) <> 0 Then
            MsgBox "��ҽ��Ϊҽ�������´����ɾ����ҽ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        'ת�Ʋ���
        If CheckOtherDeptPatiOpt = False Then Exit Sub

        If InStr(",1,2,", .TextMatrix(.Row, COL_ҽ��״̬)) = 0 Then
            MsgBox "��ǰѡ���ҽ���Ѿ���У�ԣ�����ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ǩ����ҽ������ɾ��
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            MsgBox "��ǰѡ���ҽ���Ѿ�ǩ��������ɾ��������ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If

        If mint���� = 1 Then
            '��ʿ�����Ѿ�����˵�ҽ�����������޸�ɾ����
            If .TextMatrix(.Row, COL_����ҽ��) Like "*/*" Then
                MsgBox "��ǰѡ���ҽ���Ѿ���ҽ����ˣ�����ɾ����", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            '��ִҵ�ʸ��ҽ��ֻ��ɾ���޸�δ��˵�ҽ����
            If Not mblnHaveAuditPriv Then
                If HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_����ҽ��))) Then
                    MsgBox "��û���ʸ�ɾ����ǰѡ���ҽ�������ߵ�ǰѡ���ҽ���Ѿ�����ˣ�����ɾ����", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        '�Ѿ�ִ�еǼǵ�ҽ����Ŀ����ɾ��
        If mlng·��״̬ = 1 Then
            If CheckPathAdviceIsExe(lngҽ��ID) Then
                MsgBox "��ҽ����Ӧ����Ŀ�Ѿ�ִ�С�" & vbCrLf & "��ȡ��ִ�еǼǺ��ٽ���ɾ��������", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Sub
            End If
        End If
        
        '����Ѫ��ϵͳ��Ѫҽ��ɾ�����ƣ�����Ѫ����˽׶ε��¿�ҽ������ɾ
        bln��Ѫ = gblnѪ��ϵͳ And .TextMatrix(.Row, COL_�������) = "K"
        If gblnѪ��ϵͳ And .TextMatrix(.Row, COL_�������) = "K" And InStr("5,2", Val(.TextMatrix(.Row, COL_���״̬))) > 0 Then
            MsgBox "����Ѫҽ���ѱ�Ѫ�����" & IIF(Val(.TextMatrix(.Row, COL_���״̬)) = 5, "������Ѫ", "�����������Ѫ") & "������ɾ��������ɾ��������Ѫ����ϵ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        bln����ҽ�� = Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 2
        
        'PASS
        If InStr(",5,6,", "," & .TextMatrix(.Row, COL_�������) & ",") > 0 Then
            strDelDrugIDs = "����ҩ��" & lngҽ��ID & "|" & .TextMatrix(.Row, COL_���ID)
        ElseIf .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "4" Then
            strDelDrugIDs = "����ҩ��" & .Cell(flexcpData, .Row, COL_���ID) & "|" & .TextMatrix(.Row, COL_ID)
        End If
        
        arrSQL = Array()

        If InStr(",5,6,", .TextMatrix(.Row, COL_�������)) > 0 Then
            If .Row - 1 >= .FixedRows Then
                If Val(.TextMatrix(.Row - 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If Not blnGroup And .Row + 1 <= .Rows - 1 Then
                If Val(.TextMatrix(.Row + 1, COL_���ID)) = Val(.TextMatrix(.Row, COL_���ID)) Then blnGroup = True
            End If
            If blnGroup Then
                lng��ID = Val(.TextMatrix(.Row, COL_���ID))
                If MsgBox("ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """������ҩƷһ����ҩ,ȷʵҪɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        ElseIf .TextMatrix(.Row, COL_�������) <> "" Then
            If .TextMatrix(.Row, COL_�������) = "K" Then
                If MsgBox("ȷʵҪȡ����Ѫ����""" & .TextMatrix(.Row, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            Else
                If MsgBox("Ҫ��""" & .TextMatrix(.Row, col_ҽ������) & """ͬʱ�����������Ŀһ��ȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    blnBat = True
                End If
            End If
        Else
            If MsgBox("ȷʵҪɾ��ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        If .TextMatrix(.Row, COL_�������) = "D" Then
            If HaveRIS And gbln����Ӱ����ϢϵͳԤԼ Then
                blnRISԤԼ = True
            End If
        End If
        
        Call CreatePlugInOK(pסԺҽ���´�, mint����)
        If blnBat Then
            lng������� = Val(.TextMatrix(.Row, COL_�������))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, COL_ҽ��״̬) = "1" And Val(.TextMatrix(i, COL_�������)) = lng������� Then
                    '����ɾ��ǰ��ҽӿ�
                    On Error Resume Next
                    If Not gobjPlugIn Is Nothing Then
                        If gobjPlugIn.AdviceDeletBefor(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(.TextMatrix(i, COL_ID)), mint����) = False Then
                            If err.Number = 0 Then Exit Sub
                        End If
                        Call zlPlugInErrH(err, "AdviceDeletBefor")
                    End If
                    If Not CheckDelAdivceOfPathItem(Val(.TextMatrix(i, COL_ID))) Then Exit Sub

                    If err.Number <> 0 Then err.Clear
                    On Error GoTo 0
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & .TextMatrix(i, COL_ID) & ",1)"
                    strDelIDs = strDelIDs & "," & .TextMatrix(i, COL_ID)
                End If
            Next
        Else
            '����ɾ��ǰ��ҽӿ�
            On Error Resume Next
            If Not gobjPlugIn Is Nothing Then
                If gobjPlugIn.AdviceDeletBefor(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, lngҽ��ID, mint����) = False Then
                    If err.Number = 0 Then Exit Sub
                End If
                Call zlPlugInErrH(err, "AdviceDeletBefor")
            End If
            If Not CheckDelAdivceOfPathItem(lngҽ��ID) Then Exit Sub    '�������ɵ�·��ҽ�����
            If err.Number <> 0 Then err.Clear
            On Error GoTo 0
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & lngҽ��ID & ",1)"
            strDelIDs = strDelIDs & "," & lngҽ��ID
        End If
        
        'ҽ����ӡ�ж�
        strDelIDs = Mid(strDelIDs, 2)
        If blnGroup Then
            For i = .Row To .FixedRows - 1 Step -1
                If .TextMatrix(i, COL_��Ч) <> "" Then
                    int��Ч = IIF(.TextMatrix(i, COL_��Ч) = "����", 0, 1)
                    Exit For
                End If
            Next
        Else
            int��Ч = IIF(.TextMatrix(.Row, COL_��Ч) = "����", 0, 1)
        End If
        strSQL = Get���˴�ӡ��¼DelSQL(4, mlng����ID, mlng��ҳID, Val(.TextMatrix(.Row, COL_Ӥ��ID)), int��Ч, , strDelIDs, Val(.TextMatrix(.Row, COL_Ӥ��ID)) <> 0, strMsg)
        
        If strMsg <> "" Then
            If MsgBox("��ɾ����ҽ����֮���ҽ���Ѿ���ӡ��������ش�" & vbCrLf & strMsg & vbCrLf & "�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        If strSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
    End With
    If blnRISԤԼ Then
        Set rsTmp = GetDataRISԤԼ(strDelIDs)
        On Error Resume Next
        For i = 1 To rsTmp.RecordCount
            If 0 <> gobjRis.HISSchedulingEx(Val(rsTmp!ID & ""), Val(rsTmp!ԤԼid & "")) Then 'ҽ��ɾ��
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ�������Ӱ����Ϣϵͳ�ӿ�(HISSchedulingEx)ȡ��ϢԤԼδ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            rsTmp.MoveNext
        Next
        err.Clear: On Error GoTo 0
    End If
    
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    If bln��Ѫ = True Then
        If InitObjBlood(True) = True Then
            If gobjPublicBlood.AdviceOperation(pסԺҽ���´�, lngҽ��ID, 2, False, strErr) = False Then
                gcnOracle.RollbackTrans: blnTrans = False
                MsgBox "Ѫ�⹫����������ʧ�ܣ���ϸ��Ϣ��" & strErr, vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            gcnOracle.RollbackTrans: blnTrans = False
            MsgBox "Ѫ�⹫����������ʧ�ܣ����飡", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    For i = 0 To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0

    '�����¿���Ϣ
    '�����ɾ����ҽ��Ҫͬ��һ���¿���Ϣ
    If gblnKSSStrict Or gbln�����ּ����� Or gbln��Ѫ�ּ����� Or gblnѪ��ϵͳ Then
        strWhere = strWhere & " And (Nvl(A.���״̬,0) Not in(1,3,7" & IIF(gblnѪ��ϵͳ = True, "", ",4,5") & ") or a.ҽ����Ч=0 and a.���״̬=1 and a.������־=1 and (instr(',5,6,',A.�������)>0 or A.�������='E' and B.��������='2'))"
    End If
    strSQL = "select 1 from ����ҽ����¼ a,������ĿĿ¼ b where a.������Ŀid=b.id(+) and A.ҽ��״̬=1 and a.����id=[1] and a.��ҳid=[2]" & strWhere & _
            " And Exists ( Select M.���� From ��Ա�� M,ִҵ��� N" & _
            " Where M.����=Decode(A.��˱��,1,Substr(A.����ҽ��,1,Instr(A.����ҽ��,'/')-1),Substr(A.����ҽ��,Instr(A.����ҽ��,'/')+1))" & _
            " And M.ִҵ���=N.���� And N.���� IN('ִҵҽʦ','ִҵ����ҽʦ')) And Nvl(A.ִ�б��,0)<>-1 And A.������Դ<>3 and Rownum<2"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.EOF Then '����������Ϣ����Ϊ����
         strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & mlng����ID & "," & mlng��ҳID & ",'ZLHIS_CIS_001',3,'" & UserInfo.���� & "'," & mlng����ID & ")"
         Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    '���������ҽ����Ϣ����
    If bln����ҽ�� Then
        strSQL = "select 1 from ����ҽ����¼ a where A.ҽ��״̬=2 and a.����id=[1] and a.��ҳid=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp.EOF Then '����������Ϣ����Ϊ����
             strSQL = "Zl_ҵ����Ϣ�嵥_Read(" & mlng����ID & "," & mlng��ҳID & ",'ZLHIS_CIS_035',2,'" & UserInfo.���� & "'," & mlng����ID & ")"
             Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If

    With vsAdvice
        '������ֱ��ɾ��
        .Redraw = False

        'ɾ��һ����ҩ��һ��ʱ����ʾ����
        If blnGroup And .Row + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(.Row, COL_���ID)) = Val(.TextMatrix(.Row + 1, COL_���ID)) Then
                If .TextMatrix(.Row, COL_��ʼʱ��) <> "" And .TextMatrix(.Row + 1, COL_��ʼʱ��) = "" Then
                    .TextMatrix(.Row + 1, COL_��Ч) = .TextMatrix(.Row, COL_��Ч)
                    .TextMatrix(.Row + 1, COL_��ʼʱ��) = .TextMatrix(.Row, COL_��ʼʱ��)
                    .TextMatrix(.Row + 1, COL_Ƶ��) = .TextMatrix(.Row, COL_Ƶ��)
                    .TextMatrix(.Row + 1, COL_�÷�) = .TextMatrix(.Row, COL_�÷�)
                End If
            End If
        End If

        lngRow = .Row
        If blnBat Then
            For i = .Rows - 1 To 1 Step -1
                If .TextMatrix(i, COL_ҽ��״̬) = "1" And Val(.TextMatrix(i, COL_�������)) = lng������� Then
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
        If lng��ID <> 0 Then
            i = .FindRow(CStr(lng��ID), , COL_���ID)
            If i <> -1 Then
                 .TextMatrix(i, COL_��) = ""
                Call SetTagһ����ҩ(i)
            End If
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = True
        Call vsAdvice_AfterRowColChange(-1, -1, .Row, .Col)    '��ɫ���������
        
        
        '�Զ�ˢ��ҽ����������
        RaiseEvent RequestRefresh(True)

        '����ɾ������ҽӿ�
        On Error Resume Next
        arrDelID = Split(strDelIDs, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                If Not gobjPlugIn Is Nothing Then
                    Call gobjPlugIn.AdviceDeleted(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(arrDelID(i)), mint����)
                    Call zlPlugInErrH(err, "AdviceDeleted")
                End If
            End If
        Next
        If err.Number <> 0 Then err.Clear
        On Error GoTo errH
    End With
    If mlng·��״̬ = 1 And Not gobjPath Is Nothing Then
        If GetInsidePrivs(p�ٴ�·��Ӧ��) <> "" Then
            Call gobjPath.zlRefresh(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, mintPState, mblnMoved, True)
        End If
    End If
    'PASSҽ��ɾ�����Զ�������鹦��
    If mblnPass And mint���� = 0 Then
        Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 4, strDelDrugIDs)
    End If
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckOtherDeptPatiOpt() As Boolean
'���ܣ����ת�Ʋ��˵ĵ�ǰҽ���Ƿ��������
    
     'ת�Ʋ���
    If mintPState = ps���ת�� Then
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������ID)) <> mlng����ID Then
            MsgBox "������������������´�Ĳ���ҽ����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CheckOtherDeptPatiOpt = True
End Function

Private Sub FuncAdviceModi()
'���ܣ��޸ĵ�ǰҽ��
    Dim lngҽ��ID As Long
    Dim datTurn As Date
    
    If Not CheckWindow Then Exit Sub
    With vsAdvice
        If CheckAdviceAddModi(1, lngҽ��ID, datTurn) = False Then Exit Sub
        Set mfrmEdit = frmInAdviceEdit
        Call frmInAdviceEdit.ShowMe(mfrmParent, mint����, mMainPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, Val(.TextMatrix(.Row, COL_Ӥ��ID)), lngҽ��ID, , mlng�������ID, , , mintPState, _
                 mlng����ID, mlng����ID, datTurn, mlngҽ������ID, mstrǰ��IDs, mclsMipModule, , mlng����ҽ��ID)
    End With
End Sub

Private Sub FuncAdviceSort()
'���ܣ�����ҽ��˳��
    Set mfrmEdit = frmInAdviceEdit
    Call frmInAdviceEdit.ShowMe(mfrmParent, mint����, mMainPrivs, mlng����ID, mlng��ҳID, mlngǰ��ID, _
            , , , mlng�������ID, False, 3, mintPState, mlng����ID, mlng����ID, , mlngҽ������ID)

End Sub

Private Sub FuncAdviceUnUse()
'���ܣ����δ��ҽ��
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String, strSQL As String
    Dim lngҽ��ID As Long
    Dim i As Long, strTab As String
    Dim strErr As String
    Dim bln��� As Boolean
    Dim blnFallback As Boolean
    Dim bln��Ѫ As Boolean
    Dim strԭ�� As String
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If CheckOtherDeptPatiOpt = False Then Exit Sub
    
    On Error GoTo errH
    
    With vsAdvice
        lngҽ��ID = IIF(Val(.TextMatrix(.Row, COL_���ID)) <> 0, Val(.TextMatrix(.Row, COL_���ID)), Val(.TextMatrix(.Row, COL_ID)))
        bln��Ѫ = gblnѪ��ϵͳ And .TextMatrix(.Row, COL_�������) = "K"
        If Val(.TextMatrix(.Row, COL_ִ�б��)) = -1 Then
            strMsg = strMsg & "ȷʵҪ����ǰ" & IIF(RowInһ����ҩ(.Row, 0, 0), "һ����ҩ��", "") & "ҽ��ȡ�����Ϊδ����" & _
                IIF(.TextMatrix(.Row, COL_��Ч) = "����", vbCrLf & vbCrLf & "�������Ҫ���·���ҽ���Բ������ú�ִ�С�", "")
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            bln��� = False
        ElseIf Val(.TextMatrix(.Row, COL_ִ�б��)) <> -1 Then
            bln��Ѫ = gblnѪ��ϵͳ And .TextMatrix(.Row, COL_�������) = "K"
            If bln��Ѫ Then
                '��д���δ�õ�ԭ��
                Call zlCommFun.ShowMsgBox("��Ѫҽ�����δ��", "��¼��ԭ��", "ȷ��(&O),ȡ��?(&C)", Me, , , , "2", , , "ԭ��", 200, strԭ��)
                If strԭ�� = "" Then
                    MsgBox "����ʧ�ܣ�δ¼����Ѫԭ��", vbInformation, gstrSysName
                    Exit Sub
                ElseIf Len(strԭ��) > 200 Then
                    MsgBox "����ʧ�ܣ�δ¼����Ѫԭ�򳬳���ֻ��¼��200���ַ�����100�����֡�", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            If .TextMatrix(.Row, COL_��Ч) = "����" And Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 Then
                
                If GetAdviceFeeKind(lngҽ��ID) = 2 Then  'סԺҽ��վ�������ɷ��͵�����
                    strTab = "סԺ���ü�¼"
                Else
                    strTab = "������ü�¼"
                End If
            
                '�������ҽ�������ͱ�ʾ��ִ�У��������ǣ�3-ת��;5-��Ժ;6-תԺ,11-����
                If .TextMatrix(.Row, COL_�������) = "Z" And InStr(",3,5,6,11,", Val(.TextMatrix(.Row, COL_��������))) > 0 Then
                    MsgBox "������ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """�ѷ���ִ�У����ܱ��Ϊδ�á�", vbInformation, gstrSysName
                    Exit Sub
                End If
                
                '����ѷ���ҽ����Ӧ��ִ��״̬���շ�״̬
                strSQL = "Select B.ִ��״̬ as ҽ��ִ��,C.ִ��״̬ as ����ִ��,C.��¼����,C.��¼״̬" & _
                    " From ����ҽ����¼ A,����ҽ������ B," & strTab & " C" & _
                    " Where A.ID=[1] And A.ID=B.ҽ��ID" & _
                    " And B.NO=C.NO(+) And B.��¼����=C.��¼����(+) And B.ҽ��ID=C.ҽ�����(+)" & _
                    " And (B.ִ��״̬ IN(1,3) Or C.ִ��״̬ IN(1,2) Or (C.��¼����=1 And C.��¼״̬=1))" & _
                    " Union ALL " & _
                    " Select B.ִ��״̬ as ҽ��ִ��,C.ִ��״̬ as ����ִ��,C.��¼����,C.��¼״̬" & _
                    " From ����ҽ����¼ A,����ҽ������ B," & strTab & " C" & _
                    " Where A.���ID=[1] And A.ID=B.ҽ��ID" & _
                    " And B.NO=C.NO(+) And B.��¼����=C.��¼����(+) And B.ҽ��ID=C.ҽ�����(+)" & _
                    " And (B.ִ��״̬ IN(1,3) Or C.ִ��״̬ IN(1,2) Or (C.��¼����=1 And C.��¼״̬=1))"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lngҽ��ID)
                If Not rsTmp.EOF Then
                    If InStr(",1,3,", NVL(rsTmp!ҽ��ִ��, 0)) > 0 Then
                        strMsg = "ҽ���Ѿ�ִ�л�����ִ�С�"
                    ElseIf InStr(",1,2,", NVL(rsTmp!����ִ��, 0)) > 0 Then
                        strMsg = "ҽ�������ķ����Ѿ�ִ�л򲿷�ִ�С�"
                    ElseIf NVL(rsTmp!��¼����, 0) = 1 And NVL(rsTmp!��¼״̬, 0) = 1 Then
                        strMsg = "ҽ�������ķ����Ѿ��������շѡ�"
                    End If
                    MsgBox "��ǰ" & IIF(RowInһ����ҩ(.Row, 0, 0), "һ����ҩ��", "") & "ҽ�����ܱ��Ϊδ�ã�" & strMsg, vbInformation, gstrSysName
                    Exit Sub
                End If
                
                strMsg = "��ҽ���Ѿ����ͣ����Ϊδ�ý�ȡ��������ķ��úͷ���״̬��" & vbCrLf & vbCrLf
            End If
            strMsg = strMsg & "ȷʵҪ����ǰ" & IIF(RowInһ����ҩ(.Row, 0, 0), "һ����ҩ��", "") & "ҽ�����Ϊδ����"
            If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            bln��� = True
        End If
    End With
    
    If bln��Ѫ Then
        If InitObjBlood(True) Then
            If gobjPublicBlood.AdviceTermination(pסԺҽ��վ, lngҽ��ID, bln���, False, strErr, blnFallback) = False Then
                MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    End If
    
    strSQL = "Zl_����ҽ����¼_δ��(" & lngҽ��ID
    If bln��� Then
        strSQL = strSQL & ",-1"
    Else
        strSQL = strSQL & ",0"
    End If
    If bln��Ѫ Then
        strSQL = strSQL & ",1," & IIF(blnFallback, "null,", "1,") & IIF(strԭ�� = "", "null,", "'" & strԭ�� & "',") & "'" & UserInfo.���� & "')"
    Else
        strSQL = strSQL & ")"
    End If
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "6" Then
        'ɾ��������������ϡ��еļ�¼
        Call InitObjLis(pסԺҽ��վ)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(CStr(lngҽ��ID), strErr) = False Then
                MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
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
'���ܣ���ͣҽ��
    Dim lngҽ��ID As Long, blnOnePati As Boolean
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Or mblnBatch Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If mblnDirect And mblnBatch Then
        blnOnePati = False
    ElseIf mblnDirect And Not mblnBatch Or mblnInsideTools Then
        blnOnePati = True
    Else
        If mint���� = 1 Then
            blnOnePati = Val(zlDatabase.GetPara("����ҽ����ͣ", glngSys, pסԺҽ������)) = 0
        Else
            blnOnePati = True
        End If
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 5, mlng����ID, mlng��ҳID, mlng����ID, lngҽ��ID, mint���� = 1, , , , , blnOnePati, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdvicePrice()
'���ܣ��������˵�ҽ���Ƽ���Ŀ
    Dim lngҽ��ID As Long
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 4, mlng����ID, mlng��ҳID, mlng����ID, lngҽ��ID, mint���� = 1, , , , , True, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
End Sub

Private Sub FuncAdviceReform()
'���ܣ�����ҽ��
    Dim strSQL As String
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If MsgBox("Ҫ�����ò��˵�ҽ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    strSQL = "ZL_����ҽ����¼_����(" & mlng����ID & "," & mlng��ҳID & ",'" & UserInfo.���� & "')"
    
    On Error GoTo errH
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    On Error GoTo 0

        '��ȡ����ʱ��
    mdat���� = GetRsRedoDate(mlng����ID, mlng��ҳID)

    If mblnDirect = False Then Call LoadAdvice
    
    If Val(zlDatabase.GetPara("�Զ�����ҽ����ӡ", glngSys, pסԺҽ������)) = 1 Then
        Call frmAdvicePrint.ShowMe(Me, mlng����ID, mlng��ҳID)
    Else
        MsgBox "����ҽ��������ϡ�", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceNoPrint()
'���ܣ����δ�ӡ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strClear As String
    Dim lngҽ��ID As Long, lngҳ�� As Long
    Dim blnTran As Boolean, i As Long
    Dim datPrint As Date
    Dim int��Ч As Integer
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If CheckOtherDeptPatiOpt = False Then Exit Sub
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        
        int��Ч = IIF(.TextMatrix(.Row, COL_��Ч) = "����", 0, 1)
        
        On Error GoTo errH
        
        If Val(.TextMatrix(.Row, COL_���δ�ӡ)) = 0 Then
            strSQL = "Select Min(ҳ��) as ҳ��,Min(��ӡʱ��) as ��ӡʱ��,Min(LPad(ҳ��,4,'0')||LPad(�к�,3,'0')) As λ��" & _
                    " From ����ҽ����ӡ Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = [1] Or ���id = [1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncAdviceNoPrint", lngҽ��ID)
            If Not rsTmp.EOF Then
                lngҳ�� = NVL(rsTmp!ҳ��, 0)
                datPrint = NVL(rsTmp!��ӡʱ��, CDate("1900-01-01"))
            End If
            
            If lngҳ�� > 0 Then
                '�����������֮ǰ��ӡ�ģ������������δ�ӡ
                If datPrint < mdat���� And datPrint <> CDate("1900-01-01") And int��Ч = 0 Then
                    MsgBox "��ҽ�������������֮ǰ�Ѿ���ӡ���ˣ����������δ�ӡ��", vbInformation, gstrSysName
                    Exit Sub
                Else
                    lngҳ�� = GetAdvicePrintPage(mlng����ID, mlng��ҳID, Val(.TextMatrix(.Row, COL_Ӥ��ID)), int��Ч, lngҳ��)
                    If datPrint <> CDate("1900-01-01") Then
                        If MsgBox("��ҽ���Ѿ���ӡ���ˣ����������δ�ӡ��" & vbCrLf & _
                            vbCrLf & "�����ȷʵ�����δ�ӡ�������" & .TextMatrix(.Row, COL_��Ч) & "ҽ�����ڵ� " & lngҳ�� & " ҳ��ʼ�Ĵ�ӡ���ݣ���Щҳ��Ҫ���´�ӡ��" & _
                            vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        strClear = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & Val(.TextMatrix(.Row, COL_Ӥ��ID)) & "," & int��Ч & "," & lngҳ�� & ")"
                    Else
                        strClear = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & Val(.TextMatrix(.Row, COL_Ӥ��ID)) & "," & int��Ч & "," & lngҳ�� & "," & rsTmp!λ�� & ")"
                    End If
                End If
            End If
        Else
            If int��Ч = 0 And mdat���� <> CDate("1900-01-01") Then
                '���ȡ�����κ�Ӧ������ǰ��ӡ����������ȡ������
                strSQL = "Select Count(*) From ����ҽ��״̬ A,����ҽ����¼ B" & _
                    " Where A.����ʱ��+0<[2] And A.�������� Not In(1,2) And A.ҽ��ID=B.ID And (B.ID=[1] Or B.���ID=[1]) and b.ҽ����Ч=0"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncAdviceNoPrint", lngҽ��ID, mdat����)
                If Not rsTmp.EOF Then
                    If NVL(rsTmp.Fields(0).value, 0) > 0 Then
                        MsgBox "��ҽ���������������֮ǰ���δ�ӡ�ģ�������ȡ�����δ�ӡ��", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Else
                'ȡ������ʱ���������ҽ���Ƿ��Ѵ�ӡ����
                strSQL = "Select Min(ҳ��) as ҳ��,Min(��ӡʱ��) as ��ӡʱ��,Min(LPad(ҳ��,4,'0')||LPad(�к�,3,'0')) As λ�� From ����ҽ����ӡ" & _
                    " Where ҽ��id In (" & _
                        " Select ID From ����ҽ����¼ A" & _
                        " Where a.����id = [2] And a.��ҳid = [3] And Nvl(a.Ӥ��, 0) = [4] And a.ҽ����Ч = [5] " & _
                        " And ��� > (Select Max(���) From ����ҽ����¼ Where ID = [1] Or ���id = [1]))"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncAdviceNoPrint", lngҽ��ID, mlng����ID, mlng��ҳID, Val(.TextMatrix(.Row, COL_Ӥ��ID)), int��Ч)
                If Not rsTmp.EOF Then
                    lngҳ�� = NVL(rsTmp!ҳ��, 0)
                    datPrint = NVL(rsTmp!��ӡʱ��, CDate("1900-01-01"))
                End If
                If lngҳ�� > 0 Then
                    lngҳ�� = GetAdvicePrintPage(mlng����ID, mlng��ҳID, Val(.TextMatrix(.Row, COL_Ӥ��ID)), int��Ч, lngҳ��)
                    If datPrint <> CDate("1900-01-01") Then
                        If MsgBox("��ҽ��֮���ҽ���Ѿ���ӡ���ˣ�������ȡ�����δ�ӡ��" & vbCrLf & _
                            vbCrLf & "�����ȷʵ��ȡ�����δ�ӡ�������" & .TextMatrix(.Row, COL_��Ч) & "ҽ�����ڵ� " & lngҳ�� & " ҳ��ʼ�Ĵ�ӡ���ݣ���Щҳ��Ҫ���´�ӡ��" & _
                            vbCrLf & "Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                        strClear = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & Val(.TextMatrix(.Row, COL_Ӥ��ID)) & "," & int��Ч & "," & lngҳ�� & ")"
                    Else
                        strClear = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & Val(.TextMatrix(.Row, COL_Ӥ��ID)) & "," & int��Ч & "," & lngҳ�� & "," & rsTmp!λ�� & ")"
                    End If
                End If
            End If
        End If
    End With
    
    'ִ��
    gcnOracle.BeginTrans: blnTran = True
    If strClear <> "" Then
        zlDatabase.ExecuteProcedure strClear, Me.Name
    End If
    strSQL = "Zl_����ҽ����¼_���δ�ӡ(" & lngҽ��ID & "," & IIF(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���δ�ӡ)) = 0, 1, 0) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Name
    gcnOracle.CommitTrans: blnTran = False
    
    '�������ݿ�ˢ��
    With vsAdvice
        .TextMatrix(.Row, COL_���δ�ӡ) = IIF(Val(.TextMatrix(.Row, COL_���δ�ӡ)) = 0, 1, 0)
        Call SetAdviceIcon(.Row)
        For i = .Row - 1 To .FixedRows Step -1
            If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) = _
                IIF(Val(.TextMatrix(.Row, COL_���ID)) <> 0, Val(.TextMatrix(.Row, COL_���ID)), Val(.TextMatrix(.Row, COL_ID))) Then
                .TextMatrix(i, COL_���δ�ӡ) = .TextMatrix(.Row, COL_���δ�ӡ)
                Call SetAdviceIcon(i)
            Else
                Exit For
            End If
        Next
        For i = .Row + 1 To .Rows - 1
            If IIF(Val(.TextMatrix(i, COL_���ID)) <> 0, Val(.TextMatrix(i, COL_���ID)), Val(.TextMatrix(i, COL_ID))) = _
                IIF(Val(.TextMatrix(.Row, COL_���ID)) <> 0, Val(.TextMatrix(.Row, COL_���ID)), Val(.TextMatrix(.Row, COL_ID))) Then
                .TextMatrix(i, COL_���δ�ӡ) = .TextMatrix(.Row, COL_���δ�ӡ)
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
'���ܣ�����ҽ��
    Dim lngҽ��ID As Long, blnOnePati As Boolean
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Or mblnBatch Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    
    If mblnDirect And mblnBatch Then
        blnOnePati = False
    ElseIf mblnDirect And Not mblnBatch Or mblnInsideTools Then
        blnOnePati = True
    Else
        If mint���� = 1 Then
            blnOnePati = Val(zlDatabase.GetPara("����ҽ����ͣ", glngSys, pסԺҽ������)) = 0
        Else
            blnOnePati = True
        End If
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 6, mlng����ID, mlng��ҳID, mlng����ID, lngҽ��ID, mint���� = 1, , , , , blnOnePati, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then Call LoadAdvice
    End If
End Sub

Private Sub FuncAdviceRevoke()
'���ܣ�ҽ������
    Dim lngҽ��ID As Long
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
        
    If mlngǰ��ID = 0 Or mblnDirect Then
        '����ҽ��վ,��ʿվ
        
        'ת�Ʋ���
        If CheckOtherDeptPatiOpt = False Then Exit Sub
        
        If mblnDirect Then
            lngҽ��ID = 0
        Else
            lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        End If

        If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 0, mlng����ID, mlng��ҳID, mlng����ID, lngҽ��ID, mint���� = 1, , , , , True, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
            If mblnDirect = False Then
                Call LoadAdvice(True)
            End If
             'PASSҽ�����Ϻ��Զ�������鹦��
            If mblnPass And mint���� = 0 Then
                Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 3)
            End If
        End If
        
    Else
        '����ҽ��վ
        If FuncAdviceRevokeTech Then
            Call LoadAdvice(True)
        End If
    End If
       
End Sub

Private Function FuncAdviceRevokeTech() As Boolean
'ɾ������ǰҽ������(һ��ҽ������)
    Dim strSQL As String, lngҽ��ID As Long
    
    Dim strҽ��ID As String, intRule As Integer
    Dim lngǩ��id As Long, lng֤��ID As Long
    Dim strSource As String, strSign As String
    Dim strTimeStamp As String, blnTran As Boolean, strTimeStampCode As String
    Dim datCur As Date, i As Integer
    Dim arrSQL As Variant
    Dim strMsg As String, rsTmp As ADODB.Recordset
    Dim strAdvice��Ѫ As String
    Dim strErr As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ���������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
                
        If InStr(",1,2,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
            MsgBox "��ǰѡ���סԺҽ����δУ�ԣ���ֱ��ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
            MsgBox "��ǰѡ���סԺҽ���Ѿ����ϻ�ֹͣ��", vbInformation, gstrSysName
            Exit Function
        End If
        If .TextMatrix(.Row, COL_�ϴ�ִ��) <> "" Then
            MsgBox "��ǰѡ���סԺҽ���Ѿ����ͣ����������ϡ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '92129:ҽ���ѱ���Ѫ�ƽ������ܽ�������
        If .TextMatrix(.Row, COL_�������) = "K" And gblnѪ��ϵͳ And InStr(1, ",2,5,6,", "," & Val(.TextMatrix(.Row, COL_���״̬)) & ",") <> 0 Then
            On Error GoTo errH
            strSQL = "Select Nvl(ִ�з���,0) as ִ�з��� from ����ҽ����¼ A, ������ĿĿ¼ B  where A.���ID  = [1] and A.������ĿID = B.ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ������Ŀ��ִ�з���", lngҽ��ID)
            strSQL = ""
            If rsTmp.RecordCount > 0 Then
                If Val(rsTmp!ִ�з���) = 0 Then
                    MsgBox "�������ϵ���Ѫҽ��" & IIF(Val(.TextMatrix(.Row, COL_���״̬)) = 2, "�Ѿ������Ѫ", "����������Ѫ�׶�") & "������ֱ������ҽ������Ҫ����������Ѫ����ϵ��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            On Error GoTo 0
            If gblnѪ��ϵͳ Then strAdvice��Ѫ = lngҽ��ID
        End If
        
        '����ǩ��������ʾ
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ��������ϡ�", vbInformation, gstrSysName
                Else
                    MsgBox "������ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ���������ϡ�", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            If gobjESign.CertificateStoped(UserInfo.����) = False Then strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
        End If
        
        If RowInһ����ҩ(.Row, 0, 0) Then
            If MsgBox("����һ����ҩ��ҽ������һ�����ϣ�ȷʵҪ������" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("ȷʵҪ����ҽ��""" & .TextMatrix(.Row, col_ҽ������) & """��" & strSign, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        arrSQL = Array()
        
        'ҽ����ӡ���
        strSQL = Get���˴�ӡ��¼DelSQL(3, mlng����ID, mlng��ҳID, Val(.TextMatrix(.Row, COL_Ӥ��ID)), IIF(.TextMatrix(.Row, COL_��Ч) = "����", 0, 1), lngҽ��ID, , Val(.TextMatrix(.Row, COL_Ӥ��ID)) <> 0, strMsg)
        If strMsg <> "" Then
            MsgBox "�����ϵ�ҽ���а����Ѿ���ӡ��ҽ�������ش�", vbInformation, gstrSysName
        End If
        If strSQL <> "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        End If
        datCur = zlDatabase.Currentdate
        strSQL = "ZL_����ҽ����¼_����(" & lngҽ��ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "',Null,To_Date('" & Format(datCur, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����Σ��ֵҽ��_Update(3,null," & lngҽ��ID & ")"    'ɾ��Σ��ֵ��Ӧ��ϵ
        
        '����ʱ�ĵ���ǩ��
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '��ȡǩ��ҽ��Դ��
            strҽ��ID = lngҽ��ID '��ID,����Ϊ��ϸID
            intRule = ReadAdviceSignSource(4, mlng����ID, mlng��ҳID, strҽ��ID, 0, mblnMoved, strSource)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫ���ϵ���ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
                strSign = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",4," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��ID & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
                
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
    If strAdvice��Ѫ <> "" Then
        If InitObjBlood(True) Then
            If gobjPublicBlood.AdviceOperation(pסԺҽ��վ, lngҽ��ID, 4, mblnMoved, strErr) = False Then
                gcnOracle.RollbackTrans: blnTran = False
                Screen.MousePointer = 0
                MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    gcnOracle.CommitTrans: blnTran = False
    On Error GoTo 0
    If Not (mclsMipModule Is Nothing) Then
        If mclsMipModule.IsConnect Then
            With vsAdvice
                Call ZLHIS_CIS_003(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mlng����ID, "", mlng����ID, "", , mstr����, _
                    Val(.TextMatrix(.Row, COL_ID)), .TextMatrix(.Row, COL_��Ч), .TextMatrix(.Row, COL_�������), .TextMatrix(.Row, COL_��������), "", 0, UserInfo.����, datCur)
            End With
        End If
    End If
    '�������Ϻ���ҽӿ�
    On Error Resume Next
    If CreatePlugInOK(pסԺҽ���´�, mint����) Then
        Call gobjPlugIn.AdviceRevoked(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, lngҽ��ID, mint����)
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
'���ܣ����ó���
'������Index=�����ӹ�������(0,1,2)
    Dim lng���ͺ� As String, lngҽ��ID As Long
    Dim strNO As String, bln���� As Boolean
    Dim strCommon As String, intAtom As Integer
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)) <> 0 Then
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID))
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If lngҽ��ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
        
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    lng���ͺ� = Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("���ͺ�")))
    If lng���ͺ� = 0 Then Exit Sub
    
    strNO = vsAppend.TextMatrix(vsAppend.Row, COLSend("���ݺ�"))
    If strNO = "" Then Exit Sub
    
    If Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("��¼����"))) <> 2 Then Exit Sub
    
    '��ǰ�����Ƿ񻮼۵�,ֻ��ȷ��ȱʡ��ҳ���Ƿ񻮼�
    bln���� = vsAppend.TextMatrix(vsAppend.Row, COLSend("�Ʒ�״̬")) = "���ʻ���"
        
    '���÷��ò�������
    On Error Resume Next
    If gobjInExse Is Nothing Then
        Set gobjInExse = CreateObject("zl9InExse.clsInExse")
        If gobjInExse Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    
    Set mfrmBilling = Nothing
    
    If Index = 0 Then
        Set mfrmBilling = gobjInExse.CallByNurse( _
            mfrmParent, gcnOracle, gstrDBUser, glngSys, mlng����ID, mlng����ID, 0, lngҽ��ID, strNO, bln����)
    ElseIf Index = 1 Then
        Set mfrmBilling = gobjInExse.CallByNurse( _
            mfrmParent, gcnOracle, gstrDBUser, glngSys, mlng����ID, mlng����ID, lng���ͺ�, lngҽ��ID, "", bln����)
    ElseIf Index = 2 Then
        Set mfrmBilling = gobjInExse.CallByNurse( _
            mfrmParent, gcnOracle, gstrDBUser, glngSys, mlng����ID, mlng����ID, lng���ͺ�, 0, "", bln����)
    End If
    Call GlobalDeleteAtom(intAtom)
    
    If mfrmBilling Is Nothing Then
        'ˢ��ҽ��������ϸ
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col)
    End If
    
    RaiseEvent StatusTextUpdate("")
End Sub

Private Function GetUploadAdvice(ByVal lng���ͺ� As Long, ByVal lngҽ��ID As Long, Optional ByVal blnBat As Boolean) As Recordset
'���ܣ���ȡ����ҽ���ļ��˵��ݼ�¼��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    'ȡҪ���˵ļ���NO
    '�����˲�����support��������ѽ��ʵļ��ʵ��� �� ��������ѽ��˵ĵ��ݻ��ˣ���������ֻȡ���ݺŵģ����Բ�������¼����=12�ļ���
    If blnBat Then
        strSQL = "Select Distinct A.NO From ����ҽ������ A,����ҽ����¼ B, סԺ���ü�¼ C" & _
            " Where A.ҽ��ID=B.ID And c.No = a.No And c.ҽ����� = a.ҽ��id And c.��¼���� = 2 And c.��¼״̬ = 1 And A.��¼����=2 And A.���ͺ�=[1] "
    Else
        strSQL = "Select Distinct A.NO From ����ҽ������ A,����ҽ����¼ B, סԺ���ü�¼ C" & _
            " Where A.ҽ��ID=B.ID And c.No = a.No And c.ҽ����� = a.ҽ��id And c.��¼���� = 2 And c.��¼״̬ = 1 And A.��¼����=2 And A.���ͺ�=[1] And (B.ID=[2] Or B.���ID=[2])"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPublic", lng���ͺ�, lngҽ��ID)
    
    Set GetUploadAdvice = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub FuncAdviceRoll()
'���ܣ�ҽ������
'������Index=���������ڲ˵��ϵ�����
    Dim strSQL As String, lngFlag As Long
    Dim lngҽ��ID As Long, blnBat As Boolean
    Dim lngǩ��id As Long, strSign As String
    Dim vRoll As TYPE_AdviceRoll, str���� As String, blnDo As Boolean, blnTran As Boolean
    Dim lngStarPage As Long, lngӤ����� As Long, strDelPrintTag As String
    Dim strSignIDs As String, arrSignSQL As Variant
    Dim i As Long, arrSQL As Variant
    Dim strAdvices As String, strErr As String
    Dim blnIsMany   As Boolean
    Dim lngBabyEdit As Long
    Dim strAdviceIDs As String
    Dim strAllmsg As String, strMsg As String
    Dim rsUpload As Recordset
    Dim rsTmp As ADODB.Recordset
    Dim bln���ֻ��� As Boolean
    Dim lngҽ��IDToRis As Long
    Dim strLISIDs As String
    Dim varSend As Variant
    Dim lngTmp As Long
    Dim strAdvices��Ѫ As String
    Dim var��Ѫ As Variant
    Dim colSQL As Collection
    
    '(��ID)ȡһ��ҽ�������IDΪ�յ�ҽ��ID(��ҩ;��,��ҩ�÷�,��Ҫ����,�����Ŀ,������ҽ��)
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID)) <> 0 Then
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���ID))
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If lngҽ��ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub

    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    lngBabyEdit = CheckBabyEdit(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_Ӥ��ID)))
    If lngBabyEdit = 1 Then
        MsgBox "��ǰ���˲��ڱ����ң���������˲���ҽ����", vbInformation, gstrSysName
        Exit Sub
    ElseIf lngBabyEdit = 2 Then
        MsgBox "��ǰ���˵�Ӥ�����ڱ����ң����������Ӥ��ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    'ת�Ʋ���
    If CheckOtherDeptPatiOpt = False Then Exit Sub

    '������Ϣ
    If UBound(marrRollList) < 1 Then Exit Sub
    vRoll = marrRollList(1)

    'Ȩ�޼��
    If mint���� = 1 Then
        '��ʿ����
        If InStr(GetInsidePrivs(pסԺҽ������), "�������˲���") = 0 And vRoll.������Ա <> UserInfo.���� Then
            MsgBox "��û��Ȩ�޻��������˶�ҽ���Ĳ�����" & vbCrLf & vbCrLf & vRoll.�������� & vbTab, vbInformation, gstrSysName
            Exit Sub
        End If
        '��ʿ���ܻ���ҽ������
        str���� = Get��Ա����(vRoll.������Ա)
        If InStr("," & str���� & ",", ",ҽ��,") > 0 And InStr("," & str���� & ",", ",��ʿ,") = 0 Then
            MsgBox "�㲻�ܻ���ҽ����ҽ���Ĳ�����" & vbCrLf & vbCrLf & vRoll.�������� & vbTab, vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        'ҽ�����ˣ�ֻ�ܻ������ѵĲ���,�Ե���ǩ��ͬʱҲ�ж����Ƿ���˱��˵�ǩ��
        If vRoll.������Ա <> UserInfo.���� Then
            MsgBox "�㲻�ܻ��������˶�ҽ���Ĳ�����" & vbCrLf & vbCrLf & vRoll.�������� & vbTab, vbInformation, gstrSysName
            Exit Sub
        End If
    End If

    '���˷���ʱ�����з���ת��������
    If vRoll.�������� = 0 Then
        If zlDatabase.DateMoved(vRoll.����ʱ��) Then
            If MovedBySend(lngҽ��ID, vRoll.���ͺ�, 2) Then
                MsgBox "��ҽ���ķ����Ѿ�ȫ���򲿷�ת���������ݿ⣬�����������" & vbCrLf & _
                       "��������ϵͳ����Ա��ϵ������Ӧ���ݳ�ѡ���ء�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '���ܻ��˷��Ͳ���
        If Not RollFirstEnabled Then Exit Sub
    End If

    'Ƥ�Խ��������ֱ�ӻ���
    If vRoll.�������� = 10 Then
        MsgBox "Ƥ�Խ������������ֱ�ӻ��ˡ�", vbInformation, gstrSysName
        Exit Sub
    End If

    '����ǩ����飺��������ʱ
    '------------------------------------------------------------------
    If vRoll.�������� = 0 Then bln���ֻ��� = InStr(GetInsidePrivs(pסԺҽ������), ";���ֻ���ҽ��;") > 0
    If mint���� = 1 Then
        '��ʿ����
        If (vRoll.�������� = 4 Or vRoll.�������� = 8) And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 Then
            lngǩ��id = GetAdviceSign(lngҽ��ID, vRoll.��������, vRoll.������Ա, vRoll.����ʱ��)
            If lngǩ��id <> 0 Then
                MsgBox "��ҽ��" & Decode(vRoll.��������, 4, "����", 8, "ֹͣ") & "ʱ����ҽ��ǩ�����㲻��ִ�л��ˡ�", vbInformation, gstrSysName
                Exit Sub
            End If

        ElseIf vRoll.�������� = 9 Then
            lngǩ��id = GetAdviceSign(lngҽ��ID, vRoll.��������, vRoll.������Ա, vRoll.����ʱ��)
        End If

        If MsgBox("ȷʵҪ�������²�����" & vbCrLf & vbCrLf & _
                  vRoll.�������� & vbTab, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

        If vRoll.�������� = 5 Then
            '������������һ�����
            blnBat = True
        Else
            If InStr(";" & GetInsidePrivs(pסԺҽ������) & ";", ";ҽ����������;") > 0 Then
                If RollBatchNurse(lngҽ��ID, vRoll.��������, vRoll.���ͺ�, vRoll.����ʱ��, vRoll.�������� = 4 Or vRoll.�������� = 8 Or vRoll.�������� = 9, lngǩ��id, blnIsMany) Then
                    If MsgBox("��������ҽ���͵�ǰҽ��һ��ͬʱ" & _
                              Decode(vRoll.��������, 0, "����", 4, "����", 5, "����", 6, "��ͣ", 7, "����", 8, "ֹͣ", 9, "ȷ��ֹͣ", 10, "��дƤ�Խ��") & "��Ҫͬʱ������Щҽ����" & IIF(blnIsMany, vbCrLf & "ѡ��ֻ����ͬʱǩ���ĵ�ǰ���˵�����ҽ����", ""), _
                              vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        blnBat = True
                    Else
                        If vRoll.�������� = 0 And Not bln���ֻ��� Then
                            MsgBox "��û�С����ֻ���ҽ������Ȩ�ޣ�һ���͵�ҽ��ֻ��һ����ˡ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    If blnIsMany Then
                        If MsgBox("��������ҽ���͵�ǰҽ��һ��ǩ��������һ����ˣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbNo Then
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Else
        'ҽ�����ˣ�����ҽ���Ƿ�������ǩ����������ʾ
        blnBat = RollBatchDoctor(lngҽ��ID, vRoll.��������, vRoll.���ͺ�, vRoll.����ʱ��, lngǩ��id)    '����һ�������ҽ���Ƿ���ǩ��
        If vRoll.�������� = 5 Then
            '������������һ�����
            blnBat = True
        End If

        strSQL = Decode(vRoll.��������, 0, "����", 4, "����", 5, "����", 6, "��ͣ", 7, "����", 8, "ֹͣ", 9, "ȷ��ֹͣ", 10, "��дƤ�Խ��", 13, "ͣ������")
        If MsgBox("ȷʵҪ�������²�����" & vbCrLf & vbCrLf & vRoll.�������� & vbTab & _
                  IIF(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 And lngǩ��id <> 0, _
                      vbCrLf & vbCrLf & "��ʾ����ҽ��" & strSQL & "ʱ��ǩ������ͬʱ��������һ��" & strSQL & "��ǩ��������ҽ����", ""), vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub

        '����������ʾ
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 And lngǩ��id <> 0 Then
            '��ǰ������ҽ��һ�������ǩ��,�̶�һ�����(blnBat=True)
        Else
            If blnBat And vRoll.�������� <> 5 Then    '������������һ�����
                If MsgBox("��������ҽ���͵�ǰҽ��һ��ͬʱ" & strSQL & "��Ҫͬʱ������Щҽ����", _
                          vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    blnBat = False
                    If vRoll.�������� = 0 And Not bln���ֻ��� Then
                        MsgBox "��û�С����ֻ���ҽ������Ȩ�ޣ�һ���͵�ҽ��ֻ��һ����ˡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    '��ҽ�����õĽ���������м��
    If vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Then
        If Not CheckAdviceBalanceRoll(vRoll.���ͺ�, lngҽ��ID, blnBat) Then Exit Sub
    End If

    '��ʿ���ˣ���ҩƷҽ�����˵���������������
    If mint���� = 1 And vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Then
        If Not (Not blnBat And InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_�������)) = 0) Then
            If Not blnBat And InStr(",5,6,", vsAdvice.TextMatrix(vsAdvice.Row, COL_�������)) > 0 Then
                strSQL = CheckAdviceDrugSurplus(vRoll.���ͺ�, lngҽ��ID)
            Else
                strSQL = CheckAdviceDrugSurplus(vRoll.���ͺ�)
            End If
            If strSQL <> "" Then
                If MsgBox(strSQL, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
        End If
    End If

    If vRoll.�������� = 8 Then '��������ֱ�ӻ����Զ�ֹͣ
        If Not blnBat Then
            If RowIs�䷽��(vsAdvice.Row) Then
                lngFlag = 1    '��ҩ�䷽ʼ�ձ���ִ����ֹʱ��
            End If
        End If
    End If

    '���漰������ǩ���Ĳ�������ȡ��ǩ��
    '-------------------------------------------------------
    If blnBat Then
        If lngǩ��id = 0 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 Then
            lngǩ��id = GetAdviceSign(lngҽ��ID, vRoll.��������, vRoll.������Ա, vRoll.����ʱ��)
        End If
        If vRoll.�������� = 9 Then
            strSignIDs = GetAdviceSigns(lngҽ��ID, vRoll.��������, vRoll.������Ա, vRoll.����ʱ��)
        End If
    Else
        lngǩ��id = 0
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǩ����)) = 1 Then
            lngǩ��id = GetAdviceSign(lngҽ��ID, vRoll.��������, vRoll.������Ա, vRoll.����ʱ��)
        End If
    End If
    '����SQL
    arrSignSQL = Array()
    arrSQL = Array()
    If vRoll.�������� = 9 And blnBat Then
        If strSignIDs <> "" Then
            For i = 0 To UBound(Split(strSignIDs, ","))
                ReDim Preserve arrSignSQL(UBound(arrSignSQL) + 1)
                arrSignSQL(UBound(arrSignSQL)) = "zl_ҽ��ǩ����¼_Delete(" & Split(strSignIDs, ",")(i) & ")"
            Next
        End If
    Else
        If lngǩ��id <> 0 Then
            strSign = "zl_ҽ��ǩ����¼_Delete(" & lngǩ��id & ")"
        End If
    End If
    '����ܷ����ǩ��
    If strSign <> "" Or UBound(arrSignSQL) > -1 Then
        If gobjESign Is Nothing Then
            If gintCA = 0 Then
                MsgBox "ϵͳû�����õ���ǩ����֤���ģ����˲������ܼ�����", vbInformation, gstrSysName
            Else
                MsgBox "����ǩ������δ����ȷ��װ�����˲������ܼ�����", vbInformation, gstrSysName
            End If
            Exit Sub
        Else
            If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '����ǻ��˷������ѼƷ�,��δ���������ϴ�����(1.�����ǲ�������,2.Ҳ�ɲ���,Ԥ��ʱ�Զ��ϴ�)
    If blnBat Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_����ҽ����¼_��������(" & lngҽ��ID & "," & vRoll.�������� & "," & _
                                 "To_Date('" & Format(vRoll.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                 vRoll.���ͺ� & "," & lngFlag & ")"

        '����ǻ��� ȷ��ֹͣ���� ��Ҫ����Ϣ���д���
        If vRoll.�������� = 9 Then
            strAdviceIDs = GetRollAdviceIDs(lngҽ��ID, 2, vRoll.��������, vRoll.����ʱ��)
        End If
        
    Else
        If vRoll.�������� = 9 Then  '����ȷ��ֹͣ����
            lngStarPage = CheckAdvicePrinted(lngҽ��ID, lngӤ�����)
            If lngStarPage > 0 Then
                'zl_����ҽ����¼_���ˣ����л��鲻�ܻ�������֮ǰ�Ĳ���
                If MsgBox("��ҽ����ͣ��ʱ���Ѿ���ӡ�����������ӡ��¼֮����ܻ��ˡ�" & vbNewLine & "�Ƿ������" & lngStarPage & _
                          "ҳ��֮���ҽ����ӡ��¼?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    strDelPrintTag = "Zl_����ҽ����ӡ_Delete(" & mlng����ID & "," & mlng��ҳID & "," & lngӤ����� & ",0," & lngStarPage & ")"
                End If
            End If
        End If
        '�����ǩ���ģ�����˵�ʱ��ͬһǩ��IDһ�����
        If vRoll.�������� = 9 And lngǩ��id <> 0 Then
            strAdvices = GetAdvicesSameSign(lngǩ��id)
            If strAdvices = "" Then Exit Sub
            For i = 0 To UBound(Split(strAdvices, ","))
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_����ҽ����¼_����(" & Split(strAdvices, ",")(i) & "," & lngFlag & ",Null,Null,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
            Next
        Else
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_����ҽ����¼_����(" & lngҽ��ID & "," & lngFlag & ",Null,Null,'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        End If
        
    End If
    
    '·�����ˣ��ڻ���ֹͣ������ҽ��ʱ
    If mlng·��״̬ = 1 And Not gobjPath Is Nothing And (vRoll.�������� = 4 Or vRoll.�������� = 8) Then
        If blnBat Then
            strAdviceIDs = GetRollAdviceIDs(lngҽ��ID, 2, vRoll.��������, vRoll.����ʱ��)
        Else
            strAdviceIDs = GetRollAdviceIDs(lngҽ��ID, 1)
        End If
    End If
    
    varSend = Array()
    If vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Then
        If blnBat Then
            Call GetAdvicesSameSend(vRoll.���ͺ�, strLISIDs, strAdvices, "C")
        Else
            strAdvices = lngҽ��ID
            
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������)) = 6 Then
                strLISIDs = lngҽ��ID
            End If
        End If
        
        varSend = Split(strAdvices, ",")
        
        '��ȡ���˵ļ��ʵ���
        Set rsUpload = GetUploadAdvice(vRoll.���ͺ�, lngҽ��ID, blnBat)
        
        '---RIS��Ŀ�ж�
        If blnBat Then
            blnDo = HaveItemToRis(vRoll.���ͺ�, lngҽ��IDToRis)
            If blnDo Then
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ��������˰������������ϵ���Ŀ���͵�Ӱ����Ϣϵͳ�У����ܽ����������ˣ��밴����ҽ�����л��ˣ�", vbInformation, gstrSysName
                Exit Sub
            End If
            blnDo = False
        Else
            If InStr(",D,F,", vsAdvice.TextMatrix(vsAdvice.Row, COL_�������)) > 0 Or _
                vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" And InStr(",0,5,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������))) > 0 And vsAdvice.TextMatrix(vsAdvice.Row, COL_��Ч) = "����" Then
                
                lngҽ��IDToRis = lngҽ��ID
                
            End If
        End If
        
        If HaveRIS(True) And lngҽ��IDToRis <> 0 Then
            On Error Resume Next
            If gobjRis.HISRollAdvice(lngҽ��IDToRis) <> 1 Then 'RISҽ�����˲���
                MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(HISRollAdvice)δ���óɹ�������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            End If
            err.Clear: On Error GoTo 0
        End If
        '---
        
        '���ҽ�����������ҩƷ�����ѷ�ҩ���Զ�����ҩ�ѵ���������
        If blnBat Then
            '��������
            strSQL = "select a.id from ����ҽ����¼ a,����ҽ������ b where a.id=b.ҽ��id and b.���ͺ�=[1] and a.�������='D' and a.���id is null"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, vRoll.���ͺ�)
            For i = 1 To rsTmp.RecordCount
                lngTmp = MakeBillCharge(Val(rsTmp!ID & ""))
                If lngTmp = 1 Then
                    Exit Sub
                End If
                rsTmp.MoveNext
            Next
        Else
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "D" Then
                lngTmp = MakeBillCharge(lngҽ��ID)
                If lngTmp = 1 Then
                    Exit Sub
                End If
            End If
        End If
        
        'ҽ�����˷���ǰ������ҽӿ�
        Call CreatePlugInOK(pסԺҽ���´�)
        
        If Not gobjPlugIn Is Nothing Then 'ҽ�����˷���ǰ��ҽӿ�
            If UBound(varSend) > -1 Then
                On Error Resume Next
                For i = 0 To UBound(varSend)
                    If Val(varSend(i)) <> 0 Then
                        strMsg = ""
                        blnDo = gobjPlugIn.AdviceRollSendBefore(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(varSend(i)), mint����, strMsg)
                        Call zlPlugInErrH(err, "AdviceRollSendBefore")
                        If 0 = err.Number Then '�ӿ�û�г������������жϽӿڵķ���ֵ
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
    
    If gblnѪ��ϵͳ Then
        If vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Or vRoll.�������� = 4 Then '���˷��ͻ��߻������ϲ���
            strAdvices = ""
            
            If blnBat Then
                If vRoll.�������� = 4 Then
                    strAdvices��Ѫ = GetRollAdviceIDs(lngҽ��ID, 2, vRoll.��������, vRoll.����ʱ��, True)
                    strAdvices = strAdvices��Ѫ
                Else
                    Call GetAdvicesSameSend(vRoll.���ͺ�, strAdvices��Ѫ, strAdvices, "K")
                End If
            Else
                If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "K" Then
                    strAdvices��Ѫ = lngҽ��ID
                    strAdvices = lngҽ��ID
                End If
            End If
            If strAdvices��Ѫ <> "" Then
                var��Ѫ = Split(strAdvices��Ѫ, ",")
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
                If .TextMatrix(i, COL_�������) = "Z" And .TextMatrix(i, COL_��������) = "5" And .TextMatrix(i, COL_״̬) = "ֹͣ" Then
                    strSQL = "Select b.Id, b.����id, b.����id From ����ҽ����¼ A, ���˱䶯��¼ B" & _
                        " Where a.����id = b.����id And a.��ҳid = b.��ҳid And b.��ʼԭ�� = 10 And a.��ʼִ��ʱ�� = b.��ʼʱ�� And a.Id = [1]"
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                 End If
            End With
        End If
    End If
    
    '�ٴ�·���������
    If Not gobjPath Is Nothing And mlng·��״̬ = 1 And (vRoll.�������� = 8 Or vRoll.�������� = 4) And strAdviceIDs <> "" Then
        Call gobjPath.zlAddOutPathItem(strAdviceIDs, mlng����ID, mlng��ҳID, vRoll.��������, colSQL)
        If GetInsidePrivs(p�ٴ�·��Ӧ��) <> "" Then
            Call gobjPath.zlRefresh(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, mintPState, mblnMoved, True)
        End If
    End If
    
    'ִ��SQL
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTran = True
    If vRoll.�������� = 9 Then
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
    
    '�ٴ�·��
    If Not colSQL Is Nothing Then
        For i = 1 To colSQL.Count
            Call zlDatabase.ExecuteProcedure(colSQL("C" & i), Me.Name)
        Next
    End If
    
    If strAdvices��Ѫ <> "" Then
        If InitObjBlood(True) Then
            For i = 0 To UBound(var��Ѫ)
                If gobjPublicBlood.AdviceOperation(pסԺҽ��վ, Val((var��Ѫ(i))), IIF(vRoll.�������� = 0, 6, 7), mblnMoved, strErr) = False Then
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    MsgBox "Ѫ��ϵͳ�ӿڵ���ʧ�ܣ�" & strErr, vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        End If
    End If
    
    'ҽ�������ϴ�
    strAllmsg = ""
    If mint���� <> 0 And vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Then
        If gclsInsure.GetCapability(supportҽ���ϴ�, mlng����ID, mint����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, mint����) Then
            Do While Not rsUpload.EOF
                strMsg = "" '��Ϊ����һ��NO�ڿ϶�Ϊһ�����˵�,��������˲������Բ���
                'strAdvance�д��롰�ܵ�����|��ǰ���������Ա�ҽ���ӿڴ���
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 2, strMsg, , mint����, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    'δ�ύǰ�ϴ�ʧ����ع�����ֹ����
                    gcnOracle.RollbackTrans: blnTran = False
                    Screen.MousePointer = 0
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName 'ÿ����ʾ
                    Else
                        MsgBox "�����ϴ�ʧ�ܣ����˲���������ֹ��", vbExclamation, gstrSysName
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
                Call ZLHIS_CIS_024(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mlng����ID, "", _
                    .TextMatrix(i, COL_ID), .TextMatrix(i, COL_�������), .TextMatrix(i, COL_��������))
                '���˲����䶯��¼��ҽ��ʱ�����������˳�Ժҽ���ķ��͡�
                If .TextMatrix(i, COL_�������) = "Z" And .TextMatrix(i, COL_��������) = "5" And .TextMatrix(i, COL_״̬) = "ֹͣ" Then
                    Call ZLHIS_PATIENT_006(mclsMipModule, mlng����ID, mlng��ҳID, mstr����, mstr�Ա�, mstrסԺ��, rsTmp!ID, "Ԥ��Ժ", NVL(rsTmp!����ID, 0), NVL(rsTmp!����ID, 0), NVL(rsTmp!����ID, 0), NVL(rsTmp!����ID, 0), "")
                End If
            End With
        End If
    End If
    '����ȷ��ֹͣ��������Ϣ����
    If vRoll.�������� = 9 Then
        If blnBat Then
            If strAdviceIDs <> "" Then
                strSQL = "select a.����ID,a.��ҳID,nvl(a.������־,0) as ����,max(id) as ҽ��ID from ����ҽ����¼ a " & _
                    " where a.id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
                    " group by a.����ID,a.��ҳID,a.������־"
                    
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAdviceIDs)
                strMsg = ""
                rsTmp.Filter = "����=1"
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        If InStr("," & strMsg & ",", "," & rsTmp!����ID & "," & rsTmp!��ҳID & ",") = 0 Then
                            strMsg = strMsg & "," & rsTmp!����ID & "," & rsTmp!��ҳID
                            Call SetCISMsg(rsTmp!����ID, rsTmp!��ҳID, rsTmp!ҽ��ID, 1)
                        End If
                        rsTmp.MoveNext
                    Next
                End If
                rsTmp.Filter = "����<>1"
                If Not rsTmp.EOF Then
                    For i = 1 To rsTmp.RecordCount
                        If InStr("," & strMsg & ",", "," & rsTmp!����ID & "," & rsTmp!��ҳID & ",") = 0 Then
                            strMsg = strMsg & "," & rsTmp!����ID & "," & rsTmp!��ҳID
                            Call SetCISMsg(rsTmp!����ID, rsTmp!��ҳID, rsTmp!ҽ��ID, 0)
                        End If
                        rsTmp.MoveNext
                    Next
                End If
                strMsg = ""
            End If
        Else
            With vsAdvice
                strSQL = "select 1 From ҵ����Ϣ�嵥 A Where a.����id=[1] And a.����id=[2] And a.���ͱ��� ='ZLHIS_CIS_002' And a.���ȳ̶�=[3] And a.�Ƿ�����=0 And Rownum<2"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, IIF(Val(.TextMatrix(.Row, COL_��־)) = 1, 2, 1))
                If rsTmp.EOF Then
                    strSQL = "Zl_ҵ����Ϣ�嵥_Insert(" & mlng����ID & "," & mlng��ҳID & "," & mlng����ID & "," & mlng����ID & "," & IIF(mlng�������� = 1, 1, 2) & ",'����ֹͣҽ����','0010','ZLHIS_CIS_002'," & _
                        Val(.TextMatrix(.Row, COL_ID)) & "," & IIF(Val(.TextMatrix(.Row, COL_��־)) = 1, 2, 1) & ",0,null," & mlng����ID & ")"
                    Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
                End If
            End With
        End If
    End If

    '����LIS�������뵥
    If strLISIDs <> "" Then
        Call InitObjLis(pסԺҽ��վ)
        If Not gobjLIS Is Nothing Then
            If gobjLIS.DelLisApplicationForm(strLISIDs, strErr) = False Then
                MsgBox "ɾ����������ʧ�ܣ�" & strErr, vbInformation, gstrSysName
            End If
        End If
    End If

    If gobjExchange Is Nothing Then
        On Error Resume Next
        Set gobjExchange = CreateObject("zlExchange.clsExchange")
        If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
        err.Clear: On Error GoTo 0
    End If
    '�������ݽ���ƽ̨����LIS,PACSȡ�����뵥
    If Not gobjExchange Is Nothing And vRoll.�������� = 0 Then
        With vsAdvice
            If .TextMatrix(.Row, COL_�������) = "D" Then
                blnDo = True
            ElseIf .TextMatrix(.Row, COL_�������) = "E" Then
                blnDo = RowIs������(.Row)
            End If
            If blnDo Then
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_�������) = "D", 2, 1), "����ID::" & mlng����ID & "||��ҳID::" & mlng��ҳID & "||ҽ��ID::" & lngҽ��ID & "||��������::0||��������::" & IIF(blnBat, "1", "0"))
            End If
        End With
    End If
    
    If Not gobjPlugIn Is Nothing Then 'ҽ�����˷��ͺ���ҽӿ�
        If UBound(varSend) > -1 Then
            On Error Resume Next
            For i = 0 To UBound(varSend)
                If Val(varSend(i)) <> 0 Then
                    strMsg = ""
                    blnDo = gobjPlugIn.AdviceRollSend(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(varSend(i)), mint����, strMsg)
                    Call zlPlugInErrH(err, "AdviceRollSend")
                    If 0 = err.Number Then '�ӿ�û�г������������жϽӿڵķ���ֵ
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
    
    'ˢ������
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "Z" _
       And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������))) > 0 _
       And vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Then
        '���˳�Ժҽ��ˢ��������
        RaiseEvent RequestRefresh(False)
    Else
        RaiseEvent StatusTextUpdate("")
        Call LoadAdvice
    End If
    
    'ҽ�������ϴ�
    If mint���� <> 0 And vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Then
        If gclsInsure.GetCapability(supportҽ���ϴ�, mlng����ID, mint����) And gclsInsure.GetCapability(support������ɺ��ϴ�, mlng����ID, mint����) Then
            Do While Not rsUpload.EOF
                strMsg = ""
                Screen.MousePointer = 0
                If Not gclsInsure.TranChargeDetail(2, rsUpload!NO, 2, 2, strMsg, , mint����, rsUpload.RecordCount & "|" & rsUpload.AbsolutePosition) Then
                    '�ύ���ϴ�ʧ��,����ʾ
                    If strMsg <> "" Then
                        MsgBox strMsg, vbInformation, gstrSysName
                    Else
                        MsgBox "���ʵ�""" & rsUpload!NO & """�ϴ�ʧ�ܣ�HIS���������ύ����ȷ���������ˡ�", vbExclamation, gstrSysName
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
    
    'PASSҽ�����˺��Զ�������鹦��
    If mblnPass And mint���� = 0 Then
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
        'HIS����ع��ٵ���RIS���� lngRISҽ��ID
        If HaveRIS(True) And lngҽ��IDToRis <> 0 Then
            strSQL = "Select a.����id, a.��ҳid, a.�Һŵ�, a.��������id, a.ִ�п���id, a.������ĿID,a.������� As ���, b.���ͺ�, a.Id As ҽ��id, Decode(a.�Һŵ�, Null, 2, 1) As ������Դ" & vbNewLine & _
                "From ����ҽ����¼ A, ����ҽ������ B Where a.Id = b.ҽ��id And a.Id =[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��IDToRis)
            If Not rsTmp.EOF Then
                Call gobjRis.HISSendAdvice(rsTmp, 2, Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""), "", Val(rsTmp!���ͺ� & ""))
            End If
        End If
    End If
End Sub

Private Function CheckAdvicePrinted(ByVal lngҽ��ID As Long, ByRef lngӤ����� As Long) As Long
'���ܣ���鵱ǰҽ����ֹͣʱ���Ƿ��Ѵ�ӡ
'���أ���ʼҳ�ţ�lngӤ�����=���ڴ��ݸ������ӡ��¼�Ĺ���
    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select Nvl(Min(ҳ��), 0) ҳ��, Nvl(Min(b.Ӥ��), 0) Ӥ�����" & vbNewLine & _
            "From ����ҽ����ӡ A, ����ҽ����¼ B" & vbNewLine & _
            "Where a.��ӡ��� = 1 And a.ҽ��id = b.Id And (b.Id = [1] Or b.���id = [1])"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    CheckAdvicePrinted = Val(rsTmp!ҳ��)
    lngӤ����� = Val(rsTmp!Ӥ�����)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RollBatchDoctor(ByVal lngҽ��ID As Long, ByVal int���� As Integer, ByVal lng���ͺ� As Long, ByVal datʱ�� As Date, lngǩ��id As Long) As Boolean
'���ܣ����ָ��ҽ����ǰ�����Ƿ�������ҽ��һ������ִ�е�,���ж��Ƿ������������
'������lngҽ��ID=���IDΪ�յ�ҽ����ID(һ��ҽ����ID)
'      int����=ҽ����������
'      datʱ��=ҽ��������ʱ��
'���أ��Ƿ��п���һ����˵�����ҽ��
'      lngǩ��ID=��ЩҪ���˵�ҽ���Ƿ���ǩ��(����,ֹͣ),�����򷵻�ǩ��ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    lngǩ��id = 0
    If int���� = 0 Then
        strSQL = "Select ҽ��ID From ����ҽ������ A Where ���ͺ�=[2]" & _
            " And Not Exists(Select ID From ����ҽ����¼ B Where B.ID=A.ҽ��ID And (ID=[1] Or ���ID=[1]))"
    Else
        strSQL = "Select ��������,����ʱ��,������Ա From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[3] And ����ʱ��=[4]"
        strSQL = "Select ҽ��ID,Nvl(ǩ��ID,0) as ǩ��ID From ����ҽ��״̬ A Where (��������,����ʱ��,������Ա)=(" & strSQL & ")" & _
            " And Not Exists(Select ID From ����ҽ����¼ B Where B.ID=A.ҽ��ID And (ID=[1] Or ���ID=[1] Or (A.��������=8 And ҽ����Ч=1)))"
    End If

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lngҽ��ID, lng���ͺ�, int����, datʱ��)
    If Not rsTmp.EOF Then
        If int���� = 0 Then
'            '����ͨ�����������ѳ�Ժ��Ԥ��Ժ���˵�ҽ������
'            strSQL = "Select C.����ID,C.��ҳID From ����ҽ������ A,����ҽ����¼ B,������ҳ C" & _
'                " Where A.ҽ��ID=B.ID And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
'                " And (C.��Ժ���� is Not NULL Or C.״̬=3) And A.���ͺ�=[1] And Rownum=1"
'            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng���ͺ�)
'            If Not rsTmp.EOF Then Exit Function
        ElseIf int���� <> 0 Then
            rsTmp.Filter = "ǩ��ID<>0"
            If Not rsTmp.EOF Then lngǩ��id = rsTmp!ǩ��ID
        End If
        RollBatchDoctor = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RollBatchNurse(ByVal lngҽ��ID As Long, ByVal int���� As Integer, _
    Optional ByVal lng���ͺ� As Long, Optional ByVal datʱ�� As Date, Optional ByVal blnCheckSign As Boolean, _
    Optional ByVal lngǩ��id As Long, Optional ByRef blnIsMany As Boolean) As Boolean
'���ܣ���ʿ���ˣ����ָ��ҽ����ǰ�����Ƿ�������ҽ��һ������ִ�е�,���ж��Ƿ������������
'������lngҽ��ID=���IDΪ�յ�ҽ����ID(һ��ҽ����ID)
'      int����=0-����,n-ҽ����������
'      lng���ͺ�=���˷���ʱ�ķ��ͺ�
'      datʱ��=ҽ��������ʱ��
'      blnCheckSign=�Ƿ������ǩ����ֻ��ȫ��δǩ���Ĳ�����һ����������(ȷ��ֹͣ����)
'      lngǩ��ID=��ǰҽ����ǩ��ID,������ȷ��ֹͣǩ����������ǩ������ʱʹ��
'      blnIsMany=������ȷ��ֹͣǩ����������ǩ������ʱ���ڷ����Ƿ�����ͬǩ��ID�Ķ���ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    blnIsMany = False
    
    If int���� = 0 Then
        strSQL = "Select ҽ��ID From ����ҽ������ A Where ���ͺ�=[2]" & _
            " And Not Exists(Select ID From ����ҽ����¼ B Where B.ID=A.ҽ��ID And (ID=[1] Or ���ID=[1]))"
    Else
        '�ſ���������(����Ϊֹͣ)
        strSQL = "Select ��������,����ʱ��,������Ա From ����ҽ��״̬ Where ҽ��ID=[1] And ��������=[3] And ����ʱ��=[4]"
        strSQL = "Select a.ҽ��ID,c.���ID,Nvl(a.ǩ��ID,0) as ǩ��ID From ����ҽ��״̬ A,����ҽ����¼ C Where a.ҽ��ID=c.ID And (a.��������,a.����ʱ��,a.������Ա)=(" & strSQL & ")" & _
            " And Not Exists(Select ID From ����ҽ����¼ B Where B.ID=A.ҽ��ID And (b.ID=[1] Or b.���ID=[1] Or (A.��������=8 And b.ҽ����Ч=1)))"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lngҽ��ID, lng���ͺ�, int����, datʱ��)
    If Not rsTmp.EOF Then
        If int���� = 0 Then
            '����ͨ�����������ѳ�Ժ��Ԥ��Ժ���˵�ҽ������
            strSQL = "Select C.����ID,C.��ҳID From ����ҽ������ A,����ҽ����¼ B,������ҳ C" & _
                " Where A.ҽ��ID=B.ID And B.����ID=C.����ID And B.��ҳID=C.��ҳID" & _
                " And (C.��Ժ���� is Not NULL Or C.״̬=3) And A.���ͺ�=[1] And Rownum=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lng���ͺ�)
            If Not rsTmp.EOF Then Exit Function
        ElseIf int���� <> 0 And blnCheckSign Then
            If int���� = 9 Then
                '���ͬһ��ǩ��ID���ж���ҽ�����򷵻�true������ʾ�û��������Ƕ���ҽ��
                If lngǩ��id <> 0 Then
                    rsTmp.Filter = "ǩ��ID=" & lngǩ��id
                    blnIsMany = rsTmp.RecordCount > 0
                    
                    '�Ƿ���ڶ��ǩ��ID�������������ˣ�������ֻ���˵�ǰ���ˣ��ٸ���blnIsMany�����Ƿ���˶���ҽ����
                    rsTmp.Filter = "ǩ��ID<>" & lngǩ��id
                    If rsTmp.EOF Then Exit Function
                End If
            Else
                '����ҽ��ǩ��������ʿһ����������(���ϻ�ֹͣ)
                rsTmp.Filter = "ǩ��ID<>0"
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
'���ܣ�ҽ������
    Dim blnRefresh As Boolean, blnOK As Boolean, lngTmp As Long
    
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    If mint���� = 1 Then
        '��ʿվ����
        On Error Resume Next
        If mintPState = psԤ�� Or mintPState = ps��Ժ Then
            Call MsgBox("�ò�����" & IIF(mintPState = psԤ��, "Ԥ", "") & "��Ժ�����������ҽ�����ͣ�", vbInformation, gstrSysName)
            Exit Sub
        End If
        If Control.ID = conMenu_Edit_SendInfusion Then
            If frmAdviceSendInfusion.ShowMe(mfrmParent, mlng����ID, mlng����ID, mlng��ҳID, mMainPrivs, blnRefresh, mblnDirect And Not mblnBatch Or mblnInsideTools Or blnOnePati, mlngҽ������ID, mlngӤ������ID) Then
                blnOK = True
            End If
        ElseIf Control.ID = conMenu_Edit_Send Then
            If frmAdviceSendALL.ShowMe(mfrmParent, mlng����ID, mlng����ID, mlng��ҳID, mMainPrivs, blnRefresh, mblnDirect And Not mblnBatch Or mblnInsideTools Or blnOnePati, mlngҽ������ID, mlngӤ������ID, mclsMipModule) Then
                blnOK = True
            End If
        End If
    Else
        If Control.ID = conMenu_Edit_SendCharge Then
            lngTmp = 1  '�����շ�
        Else
            If mlng�������� = 1 Then
                If InStr(GetInsidePrivs(pסԺҽ���´�), ";�����������;") > 0 Then
                    lngTmp = 2  '�������
                Else
                    lngTmp = 1  '�����շ�
                End If
            Else
                lngTmp = 0  'סԺ����
                If mintPState = psԤ�� Or mintPState = ps��Ժ Then
                    Call MsgBox("�ò�����" & IIF(mintPState = psԤ��, "Ԥ", "") & "��Ժ�����������ҽ�����ͣ�", vbInformation, gstrSysName)
                    Exit Sub
                End If
            End If
        End If
        'ҽ��(��)����
        If frmInAdviceSend.ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mstrǰ��IDs, mlng����ID, mlng����ID, mlng�������ID, blnRefresh, lngTmp, mlng��������, mlngҽ������ID, mclsMipModule) Then
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
'���ܣ�ҩ���շ���ѯ
    Call frmDrugSendQuery.ShowQuery(IIF(mblnDirect, mfrmParent, Me), mMainPrivs, mlng����ID, mlng����ID, mblnDirect And Not mblnBatch Or mblnInsideTools)
End Sub

Private Sub FuncAdviceStop()
'���ܣ�ֹͣҽ��
    Dim blnRefresh As Boolean, lngҽ��ID As Long

    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub

    If mlngǰ��ID = 0 Or mblnDirect Then
        '����ҽ��վ,��ʿվ

        If mblnDirect Then
            lngҽ��ID = 0
        Else
            lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
        End If

        If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 1, mlng����ID, mlng��ҳID, mlng����ID, _
                                   lngҽ��ID, mint���� = 1, blnRefresh, , , , True, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
            If mblnDirect = False Then
                If blnRefresh Then
                    '���¶�ȡ������ˢ�»���ȼ�������
                    RaiseEvent RequestRefresh(False)
                Else
                    Call LoadAdvice(True)
                End If
            End If
             'PASSҽ��ֹͣҽ���Զ�������鹦��
            If mblnPass And mint���� = 0 Then
                Call gobjPass.zlPassAdviceSave(mobjPassMap, , , 1)
            End If
        End If
    Else
        '����ҽ��վ
        If FuncAdviceStopTech Then
            Call LoadAdvice(True)
        End If
    End If
End Sub

Private Sub FuncAdviceStopAudit()
'���ܣ�ͣ�����
    Dim blnRefresh As Boolean, lngҽ��ID As Long
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If

    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 7, mlng����ID, mlng��ҳID, mlng����ID, _
        lngҽ��ID, mint���� = 1, blnRefresh, , , , True, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
        If mblnDirect = False Then
            If blnRefresh Then
                '���¶�ȡ������ˢ�»���ȼ�������
                RaiseEvent RequestRefresh(False)
            Else
                Call LoadAdvice(True)
            End If
        End If
    End If
End Sub

Private Function FuncAdviceStopTech() As Boolean
'ɾ������ǰҽ��ֹͣ(������סԺ����)
    Dim strSQL As String, lngҽ��ID As Long
    Dim strStopTime As String
    
    Dim strҽ��ID As String, intRule As Integer
    Dim lngǩ��id As Long, lng֤��ID As Long
    Dim strSource As String, strSign As String, strTimeStamp As String, strTimeStampCode As String
    Dim colStopTime As New Collection, blnTran As Boolean
    
    With vsAdvice
        '����Ƿ��������
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
        If lngҽ��ID = 0 Then
            MsgBox "�ò���û��ҽ������ֹͣ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        '���
        If .TextMatrix(.Row, COL_��Ч) <> "����" Then
            MsgBox "��ǰѡ���ҽ������סԺ����ҽ����", vbInformation, gstrSysName
            Exit Function
        End If
        If .TextMatrix(.Row, COL_����) <> "" Then
            MsgBox "��ҩ�䷽�ڷ��ͺ���Զ�ֹͣ��", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",1,2,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
            MsgBox "��ǰѡ���סԺҽ����δУ�ԣ���ֱ��ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
        If InStr(",4,8,9,", Val(.TextMatrix(.Row, COL_ҽ��״̬))) > 0 Then
            MsgBox "��ǰѡ���סԺҽ���Ѿ����ϻ�ֹͣ��", vbInformation, gstrSysName
            Exit Function
        End If
        
        '����ǩ��������ʾ
        If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
            If gobjESign Is Nothing Then
                If gintCA = 0 Then
                    MsgBox "ֹͣ��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ������ϵͳû������ǩ����֤���ģ�����ֹͣ��", vbInformation, gstrSysName
                Else
                    MsgBox "ֹͣ��ǩ��ҽ��ʱ��Ҫ�ٴ�ǩ����������ǩ������δ����ȷ��װ������ֹͣ��", vbInformation, gstrSysName
                End If
                Exit Function
            End If
            strSign = vbCrLf & vbCrLf & "��ʾ����ҽ���Ѿ�ǩ��������ʱ����Ҫ�ٴ�ǩ����"
        End If
        
        'ͣ��ʱȱʡ��ҽ����ֹʱ��
        strStopTime = frmAdviceStopTime.ShowMe(Me, lngҽ��ID, mlng����ID)
        If strStopTime = "" Then Exit Function
        
        strSQL = "ZL_����ҽ����¼_ֹͣ(" & lngҽ��ID & ",To_Date('" & strStopTime & "','YYYY-MM-DD HH24:MI'),'" & UserInfo.���� & "')"
        
        'ֹͣʱ�ĵ���ǩ��
        '----------------------------------------------------------------------------------
        If strSign <> "" Then
            '��ȡǩ��ҽ��Դ��
            strҽ��ID = lngҽ��ID '��ID,����Ϊ��ϸID
            colStopTime.Add Format(strStopTime, "yyyy-MM-dd HH:mm:00"), "_" & lngҽ��ID
            intRule = ReadAdviceSignSource(8, mlng����ID, mlng��ҳID, strҽ��ID, 0, mblnMoved, strSource, , colStopTime)
            If intRule = 0 Then Exit Function
            If strSource = "" Then
                MsgBox "���ܶ�ȡ��Ҫֹͣ����ǩ��ҽ��Դ�����ݡ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            strSign = gobjESign.Signature(strSource, gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
            If strSign <> "" Then
                If strTimeStamp <> "" Then
                    strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
                Else
                    strTimeStamp = "NULL"
                End If
                lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
                strSign = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",8," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & strҽ��ID & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
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
            Call ZLHIS_CIS_002(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mlng����ID, , mlng����ID, "", , mstr����, _
                lngҽ��ID, 0, 0, vsAdvice.TextMatrix(vsAdvice.Row, COL_�������), vsAdvice.TextMatrix(vsAdvice.Row, COL_��������), UserInfo.����, strTimeStamp, vsAdvice.TextMatrix(vsAdvice.Row, COL_��־))
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
'���ܣ���дƤ�Խ��
    Dim strSQL As String, str��� As String
    Dim int��� As Integer, strLabel As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTran As Boolean
    Dim dateInput As Date
    Dim strSelect As String, i As Long
    Dim strSelectInput As String
    Dim strTextInput As String
    
    Dim cnNew As ADODB.Connection
    Dim strOwner As String
    
    If mlng����ID = 0 Then Exit Sub
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
        If Not CheckAdviceIsAduit Then Exit Sub
    
    If CheckOtherDeptPatiOpt = False Then Exit Sub
    
    If Not (vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" And vsAdvice.TextMatrix(vsAdvice.Row, COL_��������) = "1") Then
        MsgBox "��ǰҽ�����ݲ��ǹ���������Ŀ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If InStr(",1,2,", vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) > 0 Then
        MsgBox "�ù�������ҽ����δͨ��У�ԣ�����У�ԡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 3 Then
        MsgBox "�ù�������ҽ����δ���ͣ�������д�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) = 4 Then
        MsgBox "�ù�������ҽ���Ѿ����ϣ�������д�����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) <> "" Then
        If MsgBox("�ù�������ҽ���Ѿ���д�˽����Ҫ������д��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    On Error GoTo errH
    
    '���������֤
    If mblnƤ����֤ Then
        Set cnNew = New ADODB.Connection
        If zlDatabase.UserIdentify(Me, "����дƤ�Խ��ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "Ƥ��ҽ�����", cnNew) = "" Then Exit Sub
    End If
    
    strSQL = "Select Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    '����
    For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(0), ","))
        strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(0), ",")(i) & "|0"
    Next
    '����
    For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(1), ","))
        strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(1), ",")(i) & "|0|2"
    Next
    strSelect = Mid(strSelect, 2)
    
    str��� = zlCommFun.ShowMsgBox("Ƥ�Խ��", vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������) & "��^^����ݹ���������ѡ����Ӧ�İ�ť������", _
            "ȷ��(&O),?ȡ��(&C)", Me, vbQuestion, "Ƥ��ʱ��", dateInput, "yyyy-MM-dd HH:mm", "Ƥ�Խ��(&P):" & strSelect, strSelectInput, _
            "������Ӧ(&F)", 50, strTextInput, , True)
    If str��� = "" Then Exit Sub
    If strSelectInput = "" Then Exit Sub
    
    
    
    If Format(IIF(mvarCond.��ʾģʽ = 0, vsAdvice.Cell(flexcpData, vsAdvice.Row, COL_��ʼʱ��), vsAdvice.TextMatrix(vsAdvice.Row, COL_��ʼʱ��)), "yyyy-MM-dd HH:mm") > dateInput Then
        MsgBox "Ƥ��ʱ�䲻����ҽ����Чʱ����ǰ��������¼�롣", vbInformation, gstrSysName
        Exit Sub
    End If
    If mbln��ʿǩ�� Then
        If Not (Check����ǩ��) Then Exit Sub
    End If
    Call GetTestLabel(rsTmp!�걾��λ, strSelectInput, strLabel, int���)
    strSQL = "ZL_����ҽ����¼_Ƥ��(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",'" & strLabel & "'," & int��� & _
    ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
    
    
    If mblnƤ����֤ And Not cnNew Is Nothing Then
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
    
    vsAdvice.TextMatrix(vsAdvice.Row, COL_Ƥ��) = strLabel
    If mvarCond.��ʾģʽ = 0 Then
        '����Ǽ��ģʽ��������ҩƷ���Ƥ�Խ����
        If InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(+)") > 0 Or InStr(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(-)") > 0 Then
            vsAdvice.TextMatrix(vsAdvice.Row, col_����) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(+)", strLabel)
            vsAdvice.TextMatrix(vsAdvice.Row, col_����) = Replace(vsAdvice.TextMatrix(vsAdvice.Row, col_����), "(-)", strLabel)
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, col_����) = vsAdvice.TextMatrix(vsAdvice.Row, col_����) & strLabel
        End If
    End If
    If int��� = 1 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_Ƥ��) = vbRed
    ElseIf int��� = 0 Then
        vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_Ƥ��) = vbBlue
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
'���ܣ�ҽ��У��
    Dim blnRefresh As Boolean, lngҽ��ID As Long, blnOnePati As Boolean
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    
    If mblnDirect Then
        lngҽ��ID = 0
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    End If
    If mblnDirect And mblnBatch Then
        blnOnePati = False
    ElseIf mblnDirect And Not mblnBatch Or mblnInsideTools Then
        blnOnePati = True
    Else
        If mint���� = 1 Then
            blnOnePati = Val(zlDatabase.GetPara("����ҽ��У��", glngSys, pסԺҽ������)) = 0
        Else
            blnOnePati = True
        End If
    End If
    
    If frmAdviceOperate.ShowMe(mfrmParent, mMainPrivs, 3, mlng����ID, mlng��ҳID, mlng����ID, _
        lngҽ��ID, mint���� = 1, blnRefresh, , , , blnOnePati, , , , , mlngҽ������ID, , mclsMipModule, , IIF(mlngFontSize = 12, 1, 0)) Then
 
        If mblnInsideTools Then
            If blnRefresh Then Call LoadAdvice
        ElseIf mblnDirect = False Then
            If blnRefresh Then
                '���¶�ȡ������ˢ�»���ȼ�������
                RaiseEvent RequestRefresh(False)
            Else
                Call LoadAdvice
            End If
        End If
    End If
    
    
End Sub

Private Sub FuncAdviceSign()
'���ܣ���ҽ�����е���ǩ��
    Dim strSQL As String, strIDs As String, i As Long
    Dim strSource As String, strSign As String
    Dim lngǩ��id As Long, lng֤��ID As Long
    Dim intRule As Integer, strTimeStamp As String, strTimeStampCode As String
    Dim ColIDs As Collection, ColSource As Collection
    
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.����) Then
        MsgBox "����ǩ��֤���ѱ�ͣ�ã�����ϵ��Ϣ�ơ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��ȡǩ��ҽ��Դ��
    intRule = ReadAdviceSignSource(1, mlng����ID, mlng��ҳID, strIDs, 0, mblnMoved, strSource, mstrǰ��IDs, , ColIDs, ColSource)
    If intRule = 0 Then Exit Sub
    If strSource = "" Then
        MsgBox "�ò���Ŀǰû�п���ǩ����ҽ����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    For i = 1 To ColIDs.Count
        strSign = gobjESign.Signature(ColSource(i), gstrDBUser, lng֤��ID, strTimeStamp, Nothing, strTimeStampCode)
        If strSign <> "" Then
            If strTimeStamp <> "" Then
                strTimeStamp = "To_Date('" & strTimeStamp & "','YYYY-MM-DD HH24:MI:SS')"
            Else
                strTimeStamp = "NULL"
            End If
            lngǩ��id = zlDatabase.GetNextID("ҽ��ǩ����¼")
            strSQL = "zl_ҽ��ǩ����¼_Insert(" & lngǩ��id & ",1," & intRule & ",'" & Replace(strSign, "'", "''") & "'," & lng֤��ID & ",'" & ColIDs(i) & "'," & strTimeStamp & ",'" & UserInfo.���� & "','" & strTimeStampCode & "')"
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            On Error GoTo 0
        End If
    Next
    If strSign <> "" Then
        Call LoadAdvice 'ˢ�½���
        MsgBox "����ɵ���ǩ����", vbInformation, gstrSysName
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncAdviceSignVerify()
'���ܣ�У��ҽ���ĵ���ǩ��(�ɶ���ת�Ƶ�����)
    Dim strSource As String
    
    If mlng����ID = 0 Then Exit Sub
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.����) Then
        MsgBox "����ǩ��֤���ѱ�ͣ�ã�����ϵ��Ϣ�ơ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "ǩ��" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '��ȡǩ��ҽ��Դ��
        If ReadAdviceSignSource(.Cell(flexcpData, .Row, 0), 0, 0, "", .RowData(.Row), mblnMoved, strSource) = 0 Then Exit Sub
        
        '��֤ǩ��
        Call gobjESign.VerifySignature(strSource, .RowData(.Row), 1)
    End With
End Sub


Private Function Check����ǩ��() As Boolean
    '�ж��Ƿ���������ǩ��
    Check����ǩ�� = True
    If gintCA > 0 And CheckSign(2, mlngҽ������ID, , , , False, gobjESign) Then
        If gobjESign Is Nothing Then
            On Error Resume Next
            Set gobjESign = CreateObject("zl9ESign.clsESign")
            err.Clear: On Error GoTo 0
            If Not gobjESign Is Nothing Then
                Call gobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If gobjESign Is Nothing Then
            MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
            Check����ǩ�� = False
            Exit Function
        Else
            If Not gobjESign.CheckCertificate(UserInfo.�û���) Then
                Check����ǩ�� = False
                Exit Function
            End If
        End If
    End If
End Function

Private Sub FuncAdviceSignErase()
'���ܣ�ȡ��ҽ���ĵ���ǩ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
        
    If mlng����ID = 0 Then Exit Sub
    If CheckDataMoved Then Exit Sub
    
    If gobjESign Is Nothing Then Exit Sub
    If gobjESign.CertificateStoped(UserInfo.����) Then
        MsgBox "����ǩ��֤���ѱ�ͣ�ã�����ϵ��Ϣ�ơ�", vbInformation, gstrSysName
        Exit Sub
    End If
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 Or tbcAppend.Selected.Tag <> "ǩ��" Then Exit Sub
    
    With vsAppend
        If .RowData(.Row) = 0 Then
            MsgBox "��ǰѡ���ҽ��û��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '���Ϻ�ֹͣҽ����ǩ������ȡ��
        If InStr(",4,8,", .Cell(flexcpData, .Row, 0)) > 0 Then
            MsgBox "����ֱ��ȡ�����ϻ�ֹͣҽ����ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        '�¿�ǩ�����������¿���У������״̬
        If .Cell(flexcpData, .Row, 0) = 1 Then
            If InStr(",1,2,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬))) = 0 Then
                MsgBox "����ҽ���Ѿ�����У�ԣ���ǩ������ȡ����", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        'ֻ��ȡ������ǩ����
        If .TextMatrix(.Row, 2) <> UserInfo.���� Then
            MsgBox "��ǩ���˲����㱾�ˣ�����ȡ��ǩ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("ȷʵҪȡ�����ǩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        If Not gobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        
        strSQL = "zl_ҽ��ǩ����¼_Delete(" & .RowData(.Row) & ")"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
    End With
    
    Call LoadAdvice 'ˢ�½���
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncToolScheme()
'���ܣ����ó��׷���ά��
    On Error Resume Next
    If gobjCISBase Is Nothing Then
        Set gobjCISBase = CreateObject("zl9CISBase.clsCISBase")
        If gobjCISBase Is Nothing Then
            MsgBox "���ƻ�������û����ȷ��װ�������޷�ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    err.Clear: On Error GoTo 0
    
    Call gobjCISBase.CallClinicScheme(mfrmParent, gcnOracle, glngSys, gstrDBUser, IIF(mint���� = 2, 3, IIF(mlng�������� = 1, 1, 2)))
End Sub

Private Sub FuncEPRReport(ByVal lngMenu As Long)
'���ܣ����ġ���ӡ��Ԥ������
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strBill As String, strTmp As String
    Dim strNO As String, int���� As Long, i As Long
    Dim lngҽ��ID As Long, lngReportID As Long, blnPrint As Boolean, bln��ӡ As Boolean
    Dim bln������ As Boolean, bln�䷽�� As Boolean, arrRPTPar(19) As String, strFlagString As String
    Dim strSQLEPR As String, rsTmpEPR As ADODB.Recordset
    Dim str��鱨��ID As String
    Dim lngViewMode As Long ' 1-������ʽ��6-�����ʽ
    Dim objRichEPR As New zlRichEPR.cRichEPR
    Dim blnLis�ӿ� As Boolean
    
        If mblnMoved Then
        MsgBox "��ǰ���˱���������ת������ͳһ�����Ӳ�������ģ���н��в鿴��", vbInformation, gstrSysName
        Exit Sub
    End If
    '�������ݽ���ƽ̨����LIS,PACS���ı���
    If lngMenu = conMenu_Edit_Compend * 10# + 1 Or lngMenu = conMenu_Edit_Compend * 10# + 6 Or lngMenu = conMenu_Edit_Compend Then
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Then
            lngViewMode = 1
        ElseIf lngMenu = conMenu_Edit_Compend * 10# + 6 Then
            lngViewMode = 6
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 1 Then
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
                '�����д���ǲɼ��������������ΪE��������ֻ�жϼ����
                Call gobjExchange.SendMsg(IIF(.TextMatrix(.Row, COL_�������) = "D", 4, 3), "ҽ��ID::" & .TextMatrix(.Row, COL_ID) & "||����Ա����::" & UserInfo.���� & "||����Աȱʡ����::" & UserInfo.������)
            End With
            Exit Sub
        End If
    End If
    
    lngReportID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_����ID))
    lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    str��鱨��ID = vsAdvice.TextMatrix(vsAdvice.Row, COL_��鱨��ID)
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_LIS����ID)) <> 0 Then
        Call FuncLisRptFileView(mfrmParent, lngҽ��ID)   '������LIS�ļ�����
        If lngReportID = 0 And str��鱨��ID = "" Then Exit Sub
    End If
    
    '���ж��Ƿ���Լ�������
    Select Case CheckEPRReport(lngҽ��ID, lngReportID, , , mblnMoved)
    Case 0
        MsgBox "��ҽ���ı���û����д��", vbInformation, gstrSysName
        Exit Sub
    Case 2
        strTmp = ""
        '����ҽ�����߱����ɫͨ����Ŀ���Բ鿴δ��ɵı���
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��־)) = 1 Then
            strTmp = "����鿴δ��ɱ���"
        Else
            If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "D" Then
                strSQL = "select 1 from Ӱ�����¼ a where a.��ɫͨ��=1 and a.ҽ��id=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                If Not rsTmp.EOF Then
                    strTmp = "����鿴δ��ɱ���"
                End If
            End If
        End If
        If InStr(GetInsidePrivs(pסԺҽ���´�), "����δ��ɱ���") > 0 Or strTmp <> "" Then
            MsgBox "ע�⣺��ҽ���ı��滹û����ʽǩ����", vbInformation, gstrSysName
        Else
            MsgBox "��ҽ���ı��滹û�����(û����ʽǩ�������ִ��)����û��Ȩ�޲�����", vbInformation, gstrSysName
            Exit Sub
        End If
    End Select
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_RIS����ID)) <> 0 Then
        If HaveRIS Then 'RIS�������
            i = gobjRis.ShowViewReport(mfrmParent.hwnd, lngҽ��ID, InStr(GetInsidePrivs(pסԺҽ���´�), ";�����ӡ;") > 0)
            If i = 0 Then Exit Sub
        End If
    End If
    
    'ִ�в���
    '�°�PACS���棬ֱ��ǿ��ʹ���°�PACS����༭��
    If str��鱨��ID <> "" Then
        Call CreateObjectPacs(mobjPublicPACS)
        Call mobjPublicPACS.zlDocShowReport(lngҽ��ID, , mblnAutoRead, mfrmParent)
    Else
        bln��ӡ = InStr(GetInsidePrivs(pסԺҽ���´�), ";�����ӡ;") > 0 And (mintPState = ps��Ժ Or mintPState = ps����)
        
        '������ĿӦ�õ���LIS�ӿ�
        If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E" Then
            Call InitObjLis(pסԺҽ��վ)
            If Not gobjLIS Is Nothing Then
                blnLis�ӿ� = True
            End If
        End If
        
        If lngMenu = conMenu_Edit_Compend * 10# + 1 Or (lngMenu = conMenu_Edit_Compend And lngViewMode = 1) Then
            '���ı���
            If blnLis�ӿ� Then
                strTmp = ""
                Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lngҽ��ID, 0, strTmp)
                If strTmp <> "" Then
                    MsgBox strTmp, vbInformation, gstrSysName
                    Exit Sub
                End If
            Else
                If mfrmParent.Name <> "frmPatiFeeQuery" Then
                    RaiseEvent ViewEPRReport(lngReportID, bln��ӡ)
                Else
                    objRichEPR.InitRichEPR gcnOracle, Me, glngSys, False
                    Call objRichEPR.ViewDocument(Me, lngReportID, bln��ӡ)
                End If
            End If
        Else
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������)) = 1 And lngMenu <> conMenu_Edit_Compend * 10# + 6 And Not (lngMenu = conMenu_Edit_Compend And lngViewMode = 6) Then
                '���༭��ʽ��ӡ��Ԥ������
                If blnLis�ӿ� Then
                    strTmp = ""
                    Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lngҽ��ID, 0, strTmp)
                    If strTmp <> "" Then
                        MsgBox strTmp, vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    RaiseEvent PrintEPRReport(lngReportID, lngMenu = conMenu_Edit_Compend * 10# + 3)
                End If
            Else
                bln������ = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������)) = 6 And vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "E"
                If Not bln������ Then bln�䷽�� = RowIs�䷽��(vsAdvice.Row)
                
                If bln������ Then
                    If blnLis�ӿ� Then
                        strTmp = ""
                        Call gobjLIS.LisSingleReprotBrowse(mfrmParent, lngҽ��ID, 1, strTmp)
                        If strTmp <> "" Then
                            MsgBox strTmp, vbInformation, gstrSysName
                            Exit Sub
                        End If
                    Else
                        '����LisWork��ӡ���鱨��
                        blnPrint = IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, True, False)
                        If Not Open_LIS_Report(Me, lngҽ��ID, mlng����ID, mblnMoved, blnPrint, Not bln��ӡ) Then
                            MsgBox "��ҽ���ı���Ϊ�°�LIS��������ʹ��(���������)���ܣ�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    '��ȡ���һ�η��͵�NO,����
                    If bln������ Or bln�䷽�� Then
                        '����ҽ��Ӧ�Լ�����Ŀ��NOΪ׼
                        strSQL = "Select ID From ����ҽ����¼ Where ���ID=[1] And Rownum=1"
                        strSQL = "Select ҽ��ID,NO,��¼���� From ����ҽ������ Where ҽ��ID=(" & strSQL & ") Order by ���ͺ� Desc"
                    Else
                        strSQL = "Select ҽ��ID,NO,��¼���� From ����ҽ������ Where ҽ��ID=[1] Order by ���ͺ� Desc"
                    End If
                    On Error GoTo errH
                                        If mblnMoved Then
                        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                    If Not rsTmp.EOF Then
                        strNO = NVL(rsTmp!NO): int���� = NVL(rsTmp!��¼����, 0)
                    End If
                    
                    '�������ʽ��ӡ��Ԥ������
                    strSQL = "Select ��� From �����ļ��б� Where ID=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�ļ�ID)))
                    If Not rsTmp.EOF Then
                        strBill = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-2"
                    End If
                    
                    If lngMenu = conMenu_Edit_Compend * 10# + 2 Then
                        If Not ReportPrintSet(gcnOracle, glngSys, strBill, Me) Then Exit Sub
                    End If
                    
                    If Not bln������ And Not bln�䷽�� Then
                        strFlagString = GetRPTPicture(mblnMoved, lngReportID, strBill, arrRPTPar)
                    End If
                    
                    If lngMenu <> conMenu_Edit_Compend * 10# + 2 And Not bln��ӡ Then
                        strTmp = "DisabledPrint=1"
                    Else
                        strTmp = "DisabledPrint=0"
                    End If
                    
                    'ҽ��IDΪ�ɼ���ʽ��ID������������ID
                    Call ReportOpen(gcnOracle, glngSys, strBill, Me, "NO=" & strNO, "����=" & int����, _
                        "ҽ��ID=" & lngҽ��ID, _
                        strFlagString, _
                        arrRPTPar(0), arrRPTPar(1), arrRPTPar(2), arrRPTPar(3), arrRPTPar(4), arrRPTPar(5), _
                        arrRPTPar(6), arrRPTPar(7), arrRPTPar(8), arrRPTPar(9), arrRPTPar(10), arrRPTPar(11), _
                        arrRPTPar(12), arrRPTPar(13), arrRPTPar(14), arrRPTPar(15), arrRPTPar(16), arrRPTPar(17), _
                        arrRPTPar(18), arrRPTPar(19), strTmp, _
                        IIF(lngMenu = conMenu_Edit_Compend * 10# + 2, 2, 1))
                End If
            End If
        End If
        '�Զ����Ϊ�Ѳ��ģ���ʿ���Ĳ���
        If mblnAutoRead And mint���� <> 1 Then Call FuncExecReportRead(True, True)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub FuncExecReportRead(ByVal blnRead As Boolean, Optional ByVal blnAuto As Boolean)
'���ܣ����õ�ǰ����Ϊ�Ѳ��ģ�����ȡ����ǰ����Ĳ���״̬
'������blnRead=���Ļ���ȡ���Ķ�״̬
'      blnAuto=����Ϊ����ʱ���Ƿ��Զ������ڵ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim strAdvice As String
    Dim strTmp As String
    Dim strErr As String
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_ID)) = 0 Then Exit Sub
        '�°�PACS�༭�����棬ֱ�ӵ��ýӿڱ������
        If .TextMatrix(.Row, COL_��鱨��ID) = "" Then
            If Val(.TextMatrix(.Row, COL_����ID)) = 0 Then Exit Sub
            If CheckOtherDeptPatiOpt = False Then Exit Sub
            
            If blnRead Then
                If Not blnAuto Then
                    If Val(.Cell(flexcpData, .Row, COL_����״̬)) = 1 Then Exit Sub '�Զ����ʱ���ƴ���
                    If MsgBox("��ȷ�ϸ���Ŀ�������Ѿ���ϸ�Ķ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                End If
                
                strSQL = "Zl_������ļ�¼_Insert(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_����ID)) & ")"
            Else
                If MsgBox("��ȷʵҪȡ���ñ���Ĳ���״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                strSQL = "Zl_������ļ�¼_Cancel(" & Val(.TextMatrix(.Row, COL_ID)) & "," & Val(.TextMatrix(.Row, COL_����ID)) & ",'" & UserInfo.���� & "')"
            End If
            Call InitObjLis(pסԺҽ��վ)
            On Error GoTo errH
            Call zlDatabase.ExecuteProcedure(strSQL, "FuncExecReportRead")
            If Not gobjLIS Is Nothing Then
                '������ñ�ǽӿ�
                strTmp = "Select ID From ����ҽ����¼ Where ID=[1] Or ���ID=[1] order by ���"
                Set rsTmp = zlDatabase.OpenSQLRecord(strTmp, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
                Do While Not rsTmp.EOF
                    strAdvice = strAdvice & "," & rsTmp!ID
                    rsTmp.MoveNext
                Loop
                If .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "6" Then
                    gobjLIS.WriteAdvicesLookState Mid(strAdvice, 2), IIF(blnRead, 1, 0)
                End If
            End If
            On Error GoTo 0
        Else
            Call CreateObjectPacs(mobjPublicPACS)
            Call mobjPublicPACS.zlDocViewStateUpdate(blnRead, Val(.TextMatrix(.Row, COL_ID)))
        End If
        
        '���ý���״̬
        If blnRead Then
            .Cell(flexcpData, .Row, COL_����״̬) = 1 '���Ѳ���
        Else
            On Error GoTo errH
            strSQL = "Select Count(1) as ���� From ������ļ�¼ Where ҽ��ID=[1] And ȡ��ʱ�� Is Null"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FuncExecReportRead", Val(.TextMatrix(.Row, COL_ID)))
            If NVL(rsTmp!����, 0) = 0 Then
                .Cell(flexcpData, .Row, COL_����״̬) = 0 '��δ����
            End If
        End If
        Call SetAdviceReportIcon(.Row)
        .TextMatrix(.Row, COL_����״̬) = "����"
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckAdviceIsAduit() As Boolean
'�ж�ҽ���Ƿ�˶�
    Dim strSQL As String, rsTmp As Recordset
    Dim strTmp As String
    Dim lngTmp As String
    
    If Val(gstrҽ���˶�) = 0 Then CheckAdviceIsAduit = True: Exit Function
    With vsAppend
        If .TextMatrix(.Row, COLSend("�������")) = "E" And .TextMatrix(.Row, COLSend("��������")) = "1" And Mid(gstrҽ���˶�, 2, 1) = "1" Or _
           .TextMatrix(.Row, COLSend("��������")) = "8" And .TextMatrix(.Row, COLSend("�������")) = "E" And Mid(gstrҽ���˶�, 1, 1) = "1" Or _
           .TextMatrix(.Row, COLSend("�������")) = "K" And Mid(gstrҽ���˶�, 1, 1) = "1" Then
            strTmp = IIF(.TextMatrix(.Row, COLSend("��������")) = "1", "Ƥ��", "��Ѫ")
            strSQL = "Select �˶��� From ����ҽ��ִ�� Where ҽ��id = [1] And ���ͺ� = [2]"
            On err GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COLSend("ҽ��ID"))), Val(.TextMatrix(.Row, COLSend("���ͺ�"))))
            If rsTmp.RecordCount = 1 Then
                If rsTmp!�˶��� & "" <> "" Then
                    CheckAdviceIsAduit = True
                Else
                    MsgBox "��ǰҽ����" & strTmp & "ҽ��������˶��˲�����ɡ�", vbInformation, gstrSysName
                End If
            ElseIf rsTmp.RecordCount > 1 Then
                lngTmp = rsTmp.RecordCount
                rsTmp.Filter = "�˶���<>''"
                If lngTmp <> rsTmp.RecordCount Then
                    MsgBox "��ǰҽ����" & strTmp & "ҽ��������δ�˶Ե�ִ�еǼǣ�����ȫ���˶��˲�����ɡ�", vbInformation, gstrSysName
                Else
                    CheckAdviceIsAduit = True
                End If
            Else
                MsgBox "��ǰҽ����" & strTmp & "ҽ���������¼ִ�������˶��˲�����ɡ�", vbInformation, gstrSysName
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
'���ܣ�ȷ��ִ�����
    Dim rsTmp As New ADODB.Recordset
    Dim lngҽ��ID As Long, lng���ͺ� As Long, lng���ID As Long
    Dim strSQL As String, strTest As String, blnTran As Boolean
    Dim str��� As String, int��� As Integer, strLabel As String
    Dim cnNew As ADODB.Connection, i As Long
    Dim strUserName As String, strOwner As String
    Dim dateInput As Date, blnIsAbnormal As Boolean
    Dim strSelect As String
    Dim strSelectInput As String
    Dim strTextInput As String
    Dim lngִ�п���ID As Long

    Dim curMoney As Currency, str��� As String, str����� As String

    With vsAppend
        lngҽ��ID = Val(.TextMatrix(.Row, COLSend("ҽ��ID")))
        lng���ͺ� = Val(.TextMatrix(.Row, COLSend("���ͺ�")))
        lng���ID = Val(.TextMatrix(.Row, COLSend("���ID")))
        lngִ�п���ID = Val(.Cell(flexcpData, .Row, COLSend("ִ�п���")))
        If Val(.Cell(flexcpData, .Row, COLSend("ִ��״̬"))) = 1 Then
            MsgBox "��ִ����Ŀ��ǰ�Ѿ�ִ����ɡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
        
        If Not CheckAdviceIsAduit Then Exit Sub
        
        '��鲡���Ƿ��������
        If Not CheckPatiIsAduit Then Exit Sub

        '�Ƿ��������δ�շѲ��˵���Ŀ:���ܼ��ʻ���,��ΪҪִ�к����,�����ſ��ܷ��͵������շ�
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_��Ч) = "����" And Val(.TextMatrix(.Row, COLSend("��¼����"))) = 1 And .Cell(flexcpData, .Row, COLSend("�Ʒ�״̬")) > 0 Then
            If Not ItemHaveCash(2, False, Val(.TextMatrix(.Row, COLSend("ҽ��ID"))), Val(.TextMatrix(.Row, COLSend("���ID"))), _
                Val(.TextMatrix(.Row, COLSend("���ͺ�"))), .TextMatrix(.Row, COLSend("�������")), .TextMatrix(.Row, COLSend("���ݺ�")), _
                    1, 0, 0, mblnMoved, CDate(.TextMatrix(.Row, COLSend("����ʱ��"))), "", "", blnIsAbnormal) Then
                If blnIsAbnormal Then
                    MsgBox "�ò��˻������쳣���ã����顣", vbInformation, gstrSysName
                Else
                    MsgBox "�ò��˻�����δ�շѵķ��ã����顣", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
        End If
        
        If Val(.TextMatrix(.Row, COLSend("��¼����"))) = 2 Then
            curMoney = GetAdviceMoney(IIF(lng���ID = 0, lngҽ��ID, lng���ID), lngҽ��ID, lng���ͺ�, str���, str�����, False, _
                IIF(Val(.TextMatrix(.Row, COLSend("�������"))) = 0, 2, 1))
            If curMoney > 0 Then
                'סԺ��Ժ���˷��ÿ���
                If Not PatiCanBilling(mlng����ID, mlng��ҳID, GetInsidePrivs(pסԺҽ������), pסԺҽ������) Then Exit Sub
                '���ʱ���
                If InitObjPublicExpense Then
                    If gobjPublicExpense.zlBillingWarn.zlBillingVerfyWarnCheck(Me, pסԺҽ������, "", .TextMatrix(.Row, COLSend("���ݺ�")), GetInsidePrivs(pסԺҽ������), mlng����ID) = False Then Exit Sub
                End If
                
                
                '����һ��ͨ���������֤,ֻ���������ʷ���
                If gdblԤ��������鿨 <> 0 And Val(.TextMatrix(.Row, COLSend("�������"))) = 1 Then
                    If Not zlDatabase.PatiIdentify(Me, glngSys, mlng����ID, curMoney, , , , IIF(-1 * gdblԤ��������鿨 >= Val(curMoney), False, True), , , (gdblԤ��������鿨 <> 0), (2 = gdblԤ��������鿨)) Then Exit Sub
                End If
            End If
        End If
    End With
    
    On Error GoTo errH

    '�ж��Ƿ�Ƥ��,����д���
    strSQL = "Select A.�������,A.Ƥ�Խ��,B.��������,Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID" & IIF(mbln��������ִ��, "(+)", "") & " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        '�Ѿ���д��Ƥ�Խ��������д
        If rsTmp!������� = "E" And NVL(rsTmp!��������) = "1" And IsNull(rsTmp!Ƥ�Խ��) Then
            '���������֤
            If mblnƤ����֤ Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "����дƤ�Խ��ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "Ƥ��ҽ�����", cnNew)
                If strUserName = "" Then Exit Sub
            End If
            '����
            For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(0), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(0), ",")(i) & "|0"
            Next
            '����
            For i = 0 To UBound(Split(Split(rsTmp!�걾��λ & "", ";")(1), ","))
                strSelect = strSelect & "," & Split(Split(rsTmp!�걾��λ & "", ";")(1), ",")(i) & "|0|2"
            Next
            strSelect = Mid(strSelect, 2)
            
            '��дƤ�Խ��
            str��� = zlCommFun.ShowMsgBox("Ƥ�Խ��", vsAdvice.TextMatrix(vsAdvice.Row, col_ҽ������) & "��^^����ݹ���������ѡ����Ӧ�İ�ť������", _
            "ȷ��(&O),?ȡ��(&C)", Me, vbQuestion, "Ƥ��ʱ��", dateInput, "yyyy-MM-dd HH:mm", "Ƥ�Խ��(&P):" & strSelect, strSelectInput, _
            "������Ӧ(&F)", 50, strTextInput, , True)
            
            If str��� = "" Then Exit Sub
            If strSelectInput = "" Then Exit Sub
            If mbln��ʿǩ�� Then
                If Not (Check����ǩ��) Then Exit Sub
            End If
            Call GetTestLabel(rsTmp!�걾��λ, strSelectInput, strLabel, int���)
            strTest = "ZL_����ҽ����¼_Ƥ��(" & lngҽ��ID & ",'" & strLabel & "'," & int��� & _
            ",'',to_date('" & dateInput & "','YYYY-MM-DD HH24:MI:SS')" & ",'" & strTextInput & "')"
        End If
    Else
        MsgBox "��Ӧ��ҽ����¼�����ڣ��޷���ɲ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If strTest = "" Then
        If MsgBox("ȷ�ϸ�ִ����Ŀִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If

    strSQL = "ZL_����ҽ��ִ��_Finish(" & lngҽ��ID & "," & lng���ͺ� & ",Null,0,'" & UserInfo.��� & "','" & UserInfo.���� & "'," & lngִ�п���ID & ")"

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
'���ܣ�ȡ��ִ�����
    Dim lng��ID As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim str������� As String, strSQL As String
    Dim byt��Դ As Byte, lngִ�п���ID As Long
    Dim strOwner As String, strUserName As String
    Dim cnNew As ADODB.Connection, rsTmp As ADODB.Recordset
    Dim bln���Ƥ�Խ�� As Boolean
    
    With vsAppend

        '��������ִ�вſ���ȡ��
        If Val(.Cell(flexcpData, .Row, COLSend("ִ��״̬"))) <> 1 Then
            MsgBox "��ִ����Ŀ��ǰ��������ִ��״̬������ȡ��ִ�С�", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
        
        '��鲡���Ƿ��������
        If Not CheckPatiIsAduit Then Exit Sub
        
        lngִ�п���ID = Val(.Cell(flexcpData, .Row, COLSend("ִ�п���")))
        lngҽ��ID = Val(.TextMatrix(.Row, COLSend("ҽ��ID")))
        lng���ͺ� = Val(.TextMatrix(.Row, COLSend("���ͺ�")))
        str������� = .TextMatrix(.Row, COLSend("�������"))
        lng��ID = IIF(Val(.TextMatrix(.Row, COLSend("���ID"))) = 0, lngҽ��ID, Val(.TextMatrix(.Row, COLSend("���ID"))))
    
        If Val(.TextMatrix(.Row, COLSend("��¼����"))) <> 1 Then
            If Val(.TextMatrix(.Row, COLSend("�������"))) = 0 Then
                byt��Դ = 2
            Else
                byt��Դ = 1
            End If
            '���ý����ж�
            If Not ItemCanCancel(lngҽ��ID, lng���ͺ�, lng��ID, str�������, False, mblnMoved, byt��Դ) Then Exit Sub
        End If
    End With
    
    If MsgBox("ȷʵҪ����ִ����Ŀȡ��ִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '�ж��Ƿ�Ƥ��,����д���
    strSQL = "Select A.�������,A.Ƥ�Խ��,B.��������,Nvl(B.�걾��λ,'����(+);����(-)') as �걾��λ From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID" & IIF(mbln��������ִ��, "(+)", "") & " And A.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        '�Ѿ���д��Ƥ�Խ��������д
        If rsTmp!������� = "E" And NVL(rsTmp!��������) = "1" And Not IsNull(rsTmp!Ƥ�Խ��) Then
            '���������֤
            If mblnƤ����֤ Then
                Set cnNew = New ADODB.Connection
                strUserName = zlDatabase.UserIdentify(Me, "��ȡ�����Ƥ��ҽ��ǰ�������������û�����������������֤��", glngSys, pסԺҽ������, "Ƥ��ҽ�����", cnNew)
                If strUserName = "" Then Exit Sub
                bln���Ƥ�Խ�� = True
            Else
                If MsgBox("�Ƿ����Ƥ�Խ����", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                    bln���Ƥ�Խ�� = False
                Else
                    bln���Ƥ�Խ�� = True
                End If
            End If
            strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngҽ��ID & "," & lng���ͺ� & "," & IIF(bln���Ƥ�Խ��, 1, 0) & ",0," & lngִ�п���ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
        Else
            strSQL = "ZL_����ҽ��ִ��_Cancel(" & lngҽ��ID & "," & lng���ͺ� & "," & "Null,0," & lngִ�п���ID & ",'" & UserInfo.��� & "','" & UserInfo.���� & "')"
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
    'Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) 'Ҫ����ִ��״̬
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
    Dim lng����id As Long, lngִ�п���ID As Long
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim blnOK As Boolean
    
    With vsAppend
        If Val(.Cell(flexcpData, .Row, COLSend("ִ��״̬"))) = 1 Then
            MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
            Exit Function
        End If
        If CheckDataMoved Then Exit Function
        
        lng����id = mlng����ID
        lngҽ��ID = Val(.TextMatrix(.Row, COLSend("ҽ��ID")))
        lng���ͺ� = Val(.TextMatrix(.Row, COLSend("���ͺ�")))
        
        RaiseEvent ExecLogNew(lngҽ��ID, lng���ͺ�, lng����id, blnOK)
        If blnOK Then
            If blnRefresh Then
                Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '����Ҫ����ִ��״̬
            End If
            FuncThingNew = True
        End If
    End With
End Function

Private Sub FuncThingModi()
    Dim lng����id As Long, lngҽ��ID As Long, lng���ͺ� As Long
    Dim strִ��ʱ�� As String, blnOK As Boolean
        
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub 'ֻ�ܲ������һ��ִ��

    If Val(gstrҽ���˶�) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) <> "" Then
        MsgBox "��ҽ�����Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsAppend
        If Val(.Cell(flexcpData, .Row, COLSend("ִ��״̬"))) = 1 Then
            MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
        
        lng����id = mlng����ID
        lngҽ��ID = Val(.TextMatrix(.Row, COLSend("ҽ��ID")))
        lng���ͺ� = Val(.TextMatrix(.Row, COLSend("���ͺ�")))
        strִ��ʱ�� = vsExec.Cell(flexcpData, vsExec.Row, 1)
    
        RaiseEvent ExecLogModi(lngҽ��ID, lng���ͺ�, lng����id, strִ��ʱ��, blnOK)
        If blnOK Then Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) '����Ҫ����ִ��״̬
    End With
End Sub

Private Sub FuncThingDel()
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim strִ��ʱ�� As String, strSQL As String
    
    If vsExec.TextMatrix(1, 0) = "" Then Exit Sub
    If vsExec.Row <> vsExec.FixedRows Then Exit Sub 'ֻ�ܲ������һ��ִ��

    With vsAppend
        If Val(.Cell(flexcpData, .Row, COLSend("ִ��״̬"))) = 1 Then '����Ͷ���ִͬ��״̬
            MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
            Exit Sub
        End If
        If Val(gstrҽ���˶�) > 0 And vsExec.TextMatrix(vsExec.FixedRows, COLExec("�˶���")) <> "" Then
            MsgBox "��ҽ�����Ѿ��˶ԣ���ȡ���˶Ժ����ԡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        If CheckDataMoved Then Exit Sub
            
        If MsgBox("ȷʵҪɾ������ִ�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lngҽ��ID = Val(.TextMatrix(.Row, COLSend("ҽ��ID")))
        lng���ͺ� = Val(.TextMatrix(.Row, COLSend("���ͺ�")))
        strִ��ʱ�� = vsExec.Cell(flexcpData, vsExec.Row, 1)
    
        strSQL = "ZL_����ҽ��ִ��_Delete(" & lngҽ��ID & "," & lng���ͺ� & ",To_Date('" & strִ��ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),0,0," & mlng����ID & ")"
        
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        On Error GoTo 0
        
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) 'Ҫ����ִ��״̬'����Ҫ����ִ��״̬
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
        If Decode(vsAppend.Tag, "�Ƽ�", True, "����", True, "ǩ��", True, False) Then
            Call SaveFlexState(vsAppend, App.ProductName & "\" & Me.Name)
        End If
    End If
    vsAppend.Tag = Item.Tag '���ڹ����������ָ��Ի�
    
    Call SetExecShow(False, False)
    
    If Item.Tag = "�Ƽ�" Then
        Call InitPriceTable
    ElseIf Item.Tag = "����" Then
        Call InitSendTable
        Call InitExecTable 'ʵ��ֻ��ִ��һ�μ���
    ElseIf Item.Tag = "ǩ��" Then
        Call InitSignTable
    ElseIf Item.Tag = "����" Then
        'NoneCode
    ElseIf Item.Tag = "����" Then
        'NoneCode
    ElseIf Item.Tag = "��ҩ" Then
    
    End If
    
    If Visible Then
        If Decode(Item.Tag, "�Ƽ�", True, "����", True, "ǩ��", True, False) Then
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
    
    picAppend.Tag = "��ִ��"
    If blnBar Then
        If blnBloodExec = False Then
            If picExec.Tag = "" Then
                lngH = vsAppend.Height - (vsAppend.Top + vsAppend.RowPos(vsAppend.Rows - 1) + vsAppend.RowHeight(vsAppend.Rows - 1) * 2)
                If lngH < picExec.Height Then
                    tbcAppend.Height = tbcAppend.Height + picExec.Height
                End If
                picExec.Visible = True: picExec.Tag = "�ɼ�"
                blnDo = True
            End If
            If picBlood.Tag = "�ɼ�" Then
                picBlood.Visible = False: picBlood.Tag = "": DkpBlood.Tag = ""
                blnDo = True: blnBlood = True
            End If
        Else
            If picBlood.Tag = "" Then
                picBlood.Visible = True: picBlood.Tag = "�ɼ�"
                Call DkpBlood_AttachPane(DkpBlood.Panes(1))
                If Not mobjFrmBlood Is Nothing Then
                    mobjFrmBlood.IsShowExec = mblnShowExec
                End If
                blnDo = True
            End If
            If picExec.Tag = "�ɼ�" Then
                picExec.Visible = False: picExec.Tag = ""
                blnDo = True
            End If
        End If
    Else
        If picExec.Tag = "�ɼ�" Then
            picExec.Visible = False: picExec.Tag = ""
            blnDo = True
        End If
        If picBlood.Tag = "�ɼ�" Then
            picBlood.Visible = False: picBlood.Tag = "": DkpBlood.Tag = ""
            blnDo = True
        End If
    End If
    
    If blnData Then
        If vsExec.Tag = "" Then '�ɼ�ʱTag=1
            lngH = vsAppend.Height - IIF(picExec.Tag = "�ɼ�", picExec.Height, 0) - (vsAppend.Top + vsAppend.RowPos(vsAppend.Rows - 1) + vsAppend.RowHeight(vsAppend.Rows - 1) * 2)
            If lngH < vsExec.Height + fraExecUD.Height Then
                tbcAppend.Height = tbcAppend.Height + fraExecUD.Height + vsExec.Height
            End If
            
            fraExecUD.Visible = True: vsExec.Visible = True: vsExec.Tag = "�ɼ�"
            blnDo = True
        End If
    Else
        If vsExec.Tag = "�ɼ�" Then
            fraExecUD.Visible = False: vsExec.Visible = False: vsExec.Tag = ""
            blnDo = True
        End If
        '��Ѫִ��blnDataʼ��ΪFalse
        If picBlood.Tag = "�ɼ�" Then
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
    TimShow.Enabled = (picBlood.Tag = "�ɼ�")
End Sub

Private Sub TimShow_Timer()
    Dim blnShowExec As Boolean
    Dim lngH As Long
    On Error GoTo ErrHand
    If picBlood.Visible = False Then Exit Sub
    If Not mobjFrmBlood Is Nothing Then
        blnShowExec = mobjFrmBlood.IsShowExec
        If DkpBlood.Tag <> IIF(blnShowExec, "�ɼ�", "���ɼ�") Then
            DkpBlood.Tag = IIF(blnShowExec, "�ɼ�", "���ɼ�")
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
    Dim blnExist As Boolean, blnSel As Boolean, bln��Ѫ As Boolean
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
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, COL_��ʼʱ��)
        End If
        If .Redraw <> flexRDNone Then
            If Val(.TextMatrix(NewRow, COL_ID)) <> 0 Then
                If mint���� = 1 And OldRow <> -1 And OldCol <> -1 And Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 And .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "1" Then
                    For i = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(i).Visible And tbcAppend(i).Tag = "����" Then
                            tbcAppend.Item(i).Selected = True
                            Exit For
                        End If
                    Next
                End If
            
                '��ʾ�����Ƿ������Ķ�
                If Val(.TextMatrix(NewRow, COL_����ID)) <> 0 Or .TextMatrix(NewRow, COL_��鱨��ID) <> "" Then
                    On Error GoTo errH
                    strSQL = "Select 1 From ������ļ�¼ Where ҽ��ID=[1]  And ������=[2] And ȡ��ʱ�� Is NULL"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", _
                        Val(.TextMatrix(NewRow, COL_ID)), UserInfo.����)
                    If Not rsTmp.EOF Then
                        If .TextMatrix(NewRow, COL_��鱨��ID) = "" Then
                            .Cell(flexcpData, NewRow, COL_����״̬) = 1
                        Else
                            '���ֲ��ĵ�
                            strSQL = "Select 1 From ����ҽ������ A Where not exists(select 1 from ������ļ�¼ B where B.ҽ��ID=A.ҽ��ID And A.��鱨��ID=B.��鱨��ID And B.������=[2] And B.ȡ��ʱ�� Is NULL) and A.ҽ��ID=[1] "
                            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "vsAdvice_AfterRowColChange", Val(.TextMatrix(NewRow, COL_ID)), UserInfo.����)
                            .Cell(flexcpData, NewRow, COL_����״̬) = IIF(Not rsTmp.EOF, 2, 1)
                        End If
                    Else
                        .Cell(flexcpData, NewRow, COL_����״̬) = 0
                    End If
                    On Error GoTo 0
                End If
                
                mblnȷ�ϻ��� = False
                If NewRow <> 0 And Val(.TextMatrix(NewRow, COL_�������)) <> 0 And Val(.TextMatrix(NewRow, COL_��������)) = 7 And .TextMatrix(NewRow, COL_�������) = "Z" And .TextMatrix(NewRow, COL_״̬) = "ֹͣ" Then
                    mblnȷ�ϻ��� = Getȷ�ϻ���(Val(.TextMatrix(NewRow, COL_ID)))
                End If
                
                '��ʾҽ�����ӱ�������
                If mblnAppend Then
                    '�жϵ��ݸ����Ƿ�������
                    blnSel = False: blnExist = False
                    Call ShowBillAppend(NewRow, blnExist)
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "����" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '�������������ظ�����
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '�жϸ�����Ϣ����ʾ
                    blnSel = False: blnExist = False
                    Call ShowAdvicePlan(NewRow, blnExist)
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "����" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '�������������ظ�����
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '�ж���Һ��ҩ����ʾ
                    blnSel = False: blnExist = False
                    If mint���� = 1 Then
                        Call ShowCompoundInfo(NewRow, blnExist)
                    End If
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "��ҩ" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '�������������ظ�����
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '�ж�ҽ���Ƿ����(ҽ�����ϲ���ʾ)
                    blnSel = False: blnExist = False: bln��Ѫ = False
                    If gblnѪ��ϵͳ And .TextMatrix(NewRow, COL_�������) = "K" Then
                        bln��Ѫ = True
                        With vsAdvice
                            '��Ѫҽ�����״̬=1��������Ѫ�Ʒ�Ѫ�����Ĵ��˶�ҽ����������Ѫҽ�������״̬=4������ҽ����δ����Ѫ�ּ�����ʱ����ʾΪ�ȴ���Ѫ
                            If Val(.TextMatrix(NewRow, COL_���״̬)) = 1 And Val(.TextMatrix(NewRow, COL_��鷽��)) = 1 Then
                                blnExist = True
                            Else
                                blnExist = InStr(",,2,3,4,5,6,", "," & .TextMatrix(NewRow, COL_���״̬) & ",") > 0 And Not (.TextMatrix(NewRow, COL_ҽ��״̬) = "4")
                            End If
                        End With
                    End If
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "ѪҺ" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '�������������ظ�����
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    blnSel = False: blnExist = False
                    If bln��Ѫ = False Then
                        With vsAdvice
                            blnExist = InStr(",2,3,4,5,", "," & .TextMatrix(NewRow, COL_���״̬) & ",") > 0
                            '����Ѫҽ��ʱ����Ѫ��ϵͳ�����Ϊ4�����״̬������ҽ����δ����Ѫ�ּ�����ʱ�� ���״̬Ϊ4ʱû����Ӧ�Ĳ�����¼<����ҽ��״̬>
                            If Val(.TextMatrix(NewRow, COL_���״̬)) = 4 And .TextMatrix(NewRow, COL_�������) = "K" Then
                                If Val(.TextMatrix(NewRow, COL_��־)) = 1 Or Not gbln��Ѫ�ּ����� Then blnExist = False
                            End If
                        End With
                    End If
                    
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "����" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '�������������ظ�����
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    '��ԤԼ��Ϣ����ʾ
                    blnSel = False: blnExist = False
                    Call ShowAdviceRISSch(NewRow, blnExist)
                    For intIdx = 0 To tbcAppend.ItemCount - 1
                        If tbcAppend(intIdx).Tag = "ԤԼ" Then
                            If tbcAppend(intIdx).Selected Then blnSel = True
                            tbcAppend(intIdx).Visible = blnExist
                            Exit For
                        End If
                    Next
                    If blnSel And Not blnExist Then
                        varDraw = .Redraw '�������������ظ�����
                        .Redraw = flexRDNone
                        tbcAppend.Item(0).Selected = True
                        .Redraw = varDraw
                    End If
                    
                    If tbcAppend.Selected.Tag = "�Ƽ�" Then
                        Call ShowPrice(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "����" Then
                        If NewRow <> 0 And Val(.TextMatrix(NewRow, COL_�������)) <> 0 And Val(.TextMatrix(NewRow, COL_��������)) = 7 And .TextMatrix(NewRow, COL_�������) = "Z" Then
                            vsAppend.ColHidden(COLSend("����ʱ��")) = False
                            vsAppend.ColHidden(COLSend("������")) = False
                            vsAppend.ColHidden(COLSend("����ʱ��")) = False
                        Else
                            vsAppend.ColHidden(COLSend("����ʱ��")) = True
                            vsAppend.ColHidden(COLSend("������")) = True
                            vsAppend.ColHidden(COLSend("����ʱ��")) = True
                        End If
                        Call ShowSendList(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "ǩ��" Then
                        Call ShowSignList(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "����" Then
                        'ǰ���ѹ̶���ȡ
                    ElseIf tbcAppend.Selected.Tag = "ԤԼ" Then
                        'ǰ���ѹ̶���ȡ
                    ElseIf tbcAppend.Selected.Tag = "����" Then
                        'ǰ���ѹ̶���ȡ
                    ElseIf tbcAppend.Selected.Tag = "��ҩ" Then
                        intIdx = IIF(vsAdvice.TextMatrix(NewRow, COL_��Ч) = "����", 0, 1)
                        Call mfrmCompoundMedicine.RefreshData(Val(vsAdvice.TextMatrix(NewRow, COL_���ID)), mlng����ID, mlng����ID, mlng��ҳID, mlng��������, intIdx, mclsMipModule, mfrmParent)
                    ElseIf tbcAppend.Selected.Tag = "����" Then
                        Call ShowOtherAppend(NewRow)
                    ElseIf tbcAppend.Selected.Tag = "ѪҺ" Then
                        If Not mobjFrmBloodList Is Nothing Then
                            Call mobjFrmBloodList.zlRefresh(Val(vsAdvice.TextMatrix(NewRow, COL_ID)), mlngFontSize, mblnMoved)
                        End If
                    End If
                End If
                
                '��ʾҽ���ɻ�������
                Call LoadRollList(NewRow)
                
                If (Not mblnShowExec) And mint���� = 1 And OldRow <> -1 And OldCol <> -1 And Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 And .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "1" And tbcAppend.Selected.Tag = "����" Then
                    mblnShowExec = Not mblnShowExec
                    Call SetExecShow(True, mblnShowExec)
                    Call vsAppend_AfterRowColChange(-1, -1, vsAppend.Row, vsAppend.Col)
                End If
                
            ElseIf mblnAppend Then
                Call ClearAppendData
            End If
            Call LoadBillList '��ʾ�ɴ�ӡ�����Ƶ���
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col_ҽ������ Or Col = col_���� Then
        vsAdvice.AutoSize Col, COL_�÷�
    ElseIf Col = COL_Ƥ�� Then
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
        ElseIf Col = COL_��ʾ Then 'Pass
            Cancel = True
        End If
    End If
End Sub

Private Sub vsAdvice_Click()
'���ܣ����ı���
    Dim lngMouseRow As Long, lngMouseCol As Long
    
    'PASS
    If mblnPass And Me.Visible Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap, 1)
    End If
    
    If mvarCond.����ģʽ <> 3 Then Exit Sub
    With vsAdvice
        lngMouseRow = .MouseRow
        lngMouseCol = .MouseCol
        
        If lngMouseRow > -1 And lngMouseCol > -1 Then
            If .Cell(flexcpFontUnderline, lngMouseRow, lngMouseCol) = True Then
                .Redraw = False
                Call FuncEPRReport(conMenu_Edit_Compend)
                .Cell(flexcpForeColor, lngMouseRow, COL_����״̬) = &H80& '����
                .Redraw = True
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_DblClick()
    Dim lngҽ��ID As Long
    Dim lngNo As Long
    Dim bln��Ѫ As Boolean
    'PASS
    If mblnPass Then
        Call gobjPass.zlPassAdviceMainPoint(mobjPassMap)
    End If
    '˫����ҽ����������뵥��ʽ�´�ĵ����鿴���� ��Ѫ�������������飬����
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_�������))
        
        If lngҽ��ID <> 0 And lngNo <> 0 Then
            If .TextMatrix(.Row, COL_�������) = "K" Then
                '��Ѫ
                If Val(Mid(gstrInUseApp, 3, 1)) = 1 Then
                    bln��Ѫ = Val(.TextMatrix(.Row, COL_��鷽��)) = 1
                    If gblnѪ��ϵͳ = True Then
                        Call frmApplyBloodNew.ShowMe(Me, mlng����ID, mlng��ҳID, 0, 2, lngҽ��ID, mlng����ID, mlng����ID, Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2), mintPState, , mrsDefine, mclsMipModule, , , , , mbytӤ��, , mlngǰ��ID, IIF(bln��Ѫ = True, 1, 0))
                    Else
                        Call frmApplyBlood.ShowMe(Me, mlng����ID, mlng��ҳID, 0, 2, lngҽ��ID, mlng����ID, mlng����ID, Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2), mintPState, , mrsDefine, mclsMipModule, , , , , mbytӤ��, , mlngǰ��ID)
                    End If
                End If
                
            ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                '����
                If Val(Mid(gstrInUseApp, 4, 1)) = 1 Then Call frmApplyOperation.ShowMe(Me, 0, 2, mlng����ID, mlng��ҳID, 0, lngҽ��ID, , , , , , , , , , , , mbytӤ��)
               
            ElseIf .TextMatrix(.Row, COL_�������) = "Z" And .TextMatrix(.Row, COL_��������) = "7" Then
                '����
                If Val(Mid(gstrInUseApp, 5, 1)) = 1 Then Call frmApplyConsultation.ShowMe(Me, lngҽ��ID, lngNo, 2, , mlng����ID, mlng��ҳID, , , , , , , , , , mbytӤ��)
                 
            ElseIf .TextMatrix(.Row, COL_�������) = "D" Then
                '���
                If Val(Mid(gstrInUseApp, 1, 1)) = 1 Then
                    Call ShowApply���(Me, lngNo)
                End If
                
            ElseIf .TextMatrix(.Row, COL_�������) = "E" And .TextMatrix(.Row, COL_��������) = "6" Then
                '����
            End If
        End If
    End With
End Sub

Private Function GetPatiInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As ADODB.Recordset
'���ܣ����ݲ���ID����ҳID��ȡ���˻�����Ϣ
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "Select A.סԺ��, A.��ǰ����, A.��������, Nvl(B.����, A.����) ����, Nvl(B.�Ա�, A.�Ա�) �Ա�, Nvl(B.����, A.����) ����, A.�����, A.������,b.�ѱ�" & vbNewLine & _
            "From ������Ϣ A, ������ҳ B" & vbNewLine & _
            "Where A.����id = B.����id And A.����id = [1] And B.��ҳid = [2]"

    Set GetPatiInfo = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng����ID, lng��ҳID)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT, vBrush As LOGBRUSH
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long

    With vsAdvice
        If Col <= .FixedCols - 1 Then
            '�����̶����еı����
            SetBkColor hDC, OS.SysColor2RGB(.BackColorFixed)

            '����߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Left + 1
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ϱ߱����
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Top + 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���±߱����
            vRect.Left = Left
            vRect.Top = Bottom - 1
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0

            '���ұ߱����
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom
            If Row = .Rows - 1 Then vRect.Bottom = vRect.Bottom - 1
            If Col = .FixedCols - 1 Then vRect.Right = vRect.Right - 1
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Else
            '����һ����ҩ������еı��߼�����
            lngLeft = COL_��Ч: lngRight = COL_��ʼʱ��
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_����: lngRight = COL_�÷�
            End If
            If Not Between(Col, lngLeft, lngRight) Then
                lngLeft = COL_Ƥ��: lngRight = COL_Ƥ��
            End If
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
            
            If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
            
            vRect.Left = Left '������߱����
            vRect.Right = Right - 1 '�����ұ߱����
            If Row = lngBegin Then
                vRect.Top = Bottom - 1 '���б�����������
                vRect.Bottom = Bottom
            Else
                If Row = lngEnd Then
                    vRect.Top = Top
                    vRect.Bottom = Bottom - 1 '���б����±���
                Else
                    vRect.Top = Top
                    vRect.Bottom = Bottom
                End If
                'Ϊ��֧��Ԥ�����
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
        If Button = 0 And lngRow > 0 Then  '���ģʽ���Ը���
            If .MouseCol = col_���� Then
                If Val(fraMore.Tag) <> lngRow Then
                    If InStr(.TextMatrix(lngRow, col_����), "����ҽ��") = 0 Then
                     
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
                        
                        fraMore.Left = .Left + .ColPos(col_����) + IIF(.ColWidth(col_����) > .ColWidthMax, .ColWidthMax, .ColWidth(col_����)) - fraMore.Width
                        
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
                If .MouseCol = COL_F��־ Then
                    If Val(.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
                        strPrompt = "����¼���ҽ��"
                    ElseIf Val(.TextMatrix(lngRow, COL_��־)) = 1 Then
                        strPrompt = "����ҽ��"
                    ElseIf Val(.TextMatrix(lngRow, COL_��־)) = 2 Then
                        strPrompt = "��¼ҽ��"
                    ElseIf .TextMatrix(lngRow, COL_Ƶ��) = "��Ҫʱ" Or .TextMatrix(lngRow, COL_Ƶ��) = "��Ҫʱ" Then
                        strPrompt = "����ҽ��"
                    End If
                     
                     '����п�����ҩ�����Ϣ��������ʾ
                    If Val(.TextMatrix(lngRow, COL_ҽ��״̬)) = 1 Then
                        Select Case Val(.TextMatrix(lngRow, COL_���״̬))
                        Case 1
                            If .TextMatrix(lngRow, COL_�������) = "K" And Val(.TextMatrix(lngRow, COL_��鷽��)) = 1 Then '��Ѫҽ�����
                                strPrompt = "��Ѫҽ�����˶�"
                            Else
                                strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "F", "����", "K", "��Ѫ", "������ҩ") & "�����"
                            End If
                        Case 2
                            If Not (.TextMatrix(lngRow, COL_�������) = "K" And Val(.TextMatrix(lngRow, COL_��鷽��)) = 1) Then
                                strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "F", "����", "K", "��Ѫ", "������ҩ") & "���ͨ��"
                            End If
                        Case 3
                            strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "F", "����", "K", "��Ѫ", "������ҩ") & "���δͨ��:" & GetKSSAuditQuestion(Val(.TextMatrix(lngRow, COL_ID)))
                        Case 7
                            strPrompt = Decode(.TextMatrix(lngRow, COL_�������), "F", "����", "K", "��Ѫ", "������ҩ") & "��ǩ��"
                        Case 4
                            If gblnѪ��ϵͳ = False Then strPrompt = "��Ѫ��Ѫ�����"
                        Case 5
                            If gblnѪ��ϵͳ = False Then strPrompt = "��ѪѪ��������Ѫ"
                        End Select
                    End If
                ElseIf .MouseCol = COL_����״̬ Then
                    If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then strPrompt = "����δ��"
                    If Val(.TextMatrix(lngRow, COL_����ID)) <> 0 Or .TextMatrix(lngRow, COL_��鱨��ID) <> "" Or _
                        Val(.TextMatrix(lngRow, COL_RIS����ID)) <> 0 Or Val(.TextMatrix(lngRow, COL_LIS����ID)) <> 0 Then
                        
                        If Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 0 Then
                            strPrompt = "����δ�ģ�����鿴"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 1 Then
                            strPrompt = "�������ģ�����鿴"
                        ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 2 Then
                            strPrompt = "���沿�����ģ�����鿴"
                        End If
                    End If
                ElseIf .MouseCol = COL_F���� Then
                    strPrompt = GetAdviceReportTip(lngRow)
                End If
            End If
            If .MouseRow > -1 And .MouseCol > -1 And mvarCond.����ģʽ = 3 And .MouseCol = COL_����״̬ Then
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
            ElseIf mvarCond.����ģʽ = 3 And strPrompt = "" Then
                Call zlCommFun.ShowTipInfo(.hwnd, "")
                mlngPromptRow = 0
            ElseIf mlngPromptRow <> 0 And lngRow <> mlngPromptRow Then
            '����֮ǰ����ʾ����
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
'���ܣ���ʾĳ��ҽ������ϸ����
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
                If .TextMatrix(0, i) = "ȷ��ͣ��ʱ��" Then
                    vsfAdivceDetail.TextMatrix(j - 1, 0) = "ȷ��ʱ��" & "��"
                Else
                    vsfAdivceDetail.TextMatrix(j - 1, 0) = .TextMatrix(0, i) & "��"
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
                PicAdviceDetail.Top = fraMore.Top - 10 '���ⶥ�˺ͱ�����غ�
            End If
            
            Call SetPicAdviceDetailEffect
            If PicAdviceDetail.Visible = False Then PicAdviceDetail.Visible = True
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetPicAdviceDetailEffect()
    Dim lngR As Long
    
    '�߿�API=RoundRect
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, 0)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    PicAdviceDetail.Line (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (0, Screen.TwipsPerPixelY)-(0, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.Line (PicAdviceDetail.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(PicAdviceDetail.Width - Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    PicAdviceDetail.PSet (Screen.TwipsPerPixelX, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    PicAdviceDetail.PSet (PicAdviceDetail.Width - Screen.TwipsPerPixelX * 2, PicAdviceDetail.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
           
    '��״
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
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    
    If mlng����ID = 0 Then Exit Sub
    strSQL = "Select NVL(B.����,A.����) ����,NVL(B.�Ա�,A.�Ա�) �Ա�,NVL(B.����,A.����) ���� ,B.סԺ��,B.��Ժ���� as ����,B.��Ժ����,B.��Ժ����" & _
        " From ������Ϣ A,������ҳ B Where A.����ID=B.����ID And A.����ID=[1] And B.��ҳID=[2]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mlng��ҳID)
    If rsTmp.EOF Then Exit Sub
    
    '��ͷ
    objOut.Title.Text = "����ҽ���嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "���ˣ�" & NVL(rsTmp!����) & " �Ա�" & NVL(rsTmp!�Ա�) & " ���䣺" & NVL(rsTmp!����)
    objRow.Add "סԺ�ţ�" & NVL(rsTmp!סԺ��) & " ���ţ�" & NVL(rsTmp!����)
    objOut.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��Ժ���ڣ�" & Format(NVL(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
    objRow.Add "��Ժ���ڣ�" & Format(NVL(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
    objOut.UnderAppRows.Add objRow
    
    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = vsAdvice
    
    '���
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
    
    'ҽ���嵥
    '-----------------------------------------------------
    Call InitAdviceTable
    Call InitColumnSelect '��ʼ����ѡ����
    
    'CommandBars
    '-----------------------------------------------------
    Call GetFilterSetting '���ع��˲���
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
        .InsertItem(0, "��������", picMain.hwnd, 0).Tag = "����������"
        .InsertItem(1, " ��  �� ", picMain.hwnd, 0).Tag = "����"
        .InsertItem(2, " ��  �� ", picMain.hwnd, 0).Tag = "����"
        .InsertItem(3, " ��  �� ", picMain.hwnd, 0).Tag = "����"
    End With
    tbcMain.Item(tbcMain.ItemCount - 1).Selected = True
    tbcMain.Item(mvarCond.����ģʽ).Selected = True
    If mvarCond.����ģʽ = 3 Then mbln���� = True
    
    With tbcAppend
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
        End With
        .InsertItem(0, "ҽ���Ƽ�����", picAppend.hwnd, 0).Tag = "�Ƽ�"
        .InsertItem(1, "ҽ�����ͼ�¼", picAppend.hwnd, 0).Tag = "����"
        If Not gobjESign Is Nothing Then  '����ǩ����¼
            .InsertItem(2, "ҽ��ǩ����¼", picAppend.hwnd, 0).Tag = "ǩ��"
        End If
        .InsertItem(3, "���븽��", rtfAppend.hwnd, 0).Tag = "����"
        .InsertItem(4, "�������", rtfInfo.hwnd, 0).Tag = "����"
        
        If gstr��Һ�������� <> "" Then
            Set mfrmCompoundMedicine = New frmCompoundMedicine
            .InsertItem(4, "��Һ��ҩ��¼", mfrmCompoundMedicine.hwnd, 0).Tag = "��ҩ"
        End If
        .InsertItem(5, "ԤԼ��Ϣ", rtfSche.hwnd, 0).Tag = "ԤԼ" 'RISԤԼ��Ϣ
        .InsertItem(6, "������Ϣ", rtfOther.hwnd, 0).Tag = "����"  '����ҩ�������Ϣ
        If gblnѪ��ϵͳ = True Then
            If InitObjBlood = True Then
                Set mobjFrmBloodList = gobjPublicBlood.zlGetBloodListInfo
                .InsertItem(7, "ѪҺ��Ϣ", mobjFrmBloodList.hwnd, 0).Tag = "ѪҺ"  'ѪҺ��Ѫ��Ϣ
            End If
        End If
        '��Ϊ����ͬ,���Ҫ�л��ص�1��;�����ݲ�Ӱ���ٶ�
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    mblnAppend = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "AppendData", 1)) <> 0
    tbcAppend.Visible = mblnAppend: fraAdviceUD.Visible = mblnAppend
    If mblnAppend Then
        strTab = zlDatabase.GetPara("ҽ�����б�", glngSys, pסԺҽ���´�, "")
        If strTab <> "" Then
            For i = 0 To tbcAppend.ItemCount - 1
                If tbcAppend(i).Visible And tbcAppend(i).Tag = strTab Then
                    tbcAppend.Item(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
    
    '����ѪҺ����
    If gblnѪ��ϵͳ = True Then
        With DkpBlood
            .Options.UseSplitterTracker = False 'ʵʱ�϶�
            .Options.ThemedFloatingFrames = True
            .Options.AlphaDockingContext = True
            .Options.HideClient = True
            
            Set objPane = .CreatePane(1, 100, 100, DockLeftOf, Nothing)
            objPane.Title = "��Ѫִ�еǼ�"
            objPane.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable
        End With
    End If
    
    '�ָ����Ի�����
    '-----------------------------------------------------
    Call RestoreWinState(Me, App.ProductName)
    vsAdvice.ColWidth(COL_F��־) = 11 * Screen.TwipsPerPixelX
    vsAdvice.ColWidth(COL_F����) = 11 * Screen.TwipsPerPixelX
    
    '������ʼ��
    '-----------------------------------------------------
    mstr����IDs = ""
    mMainPrivs = gMainPrivs '������ģ��Ȩ��
    Set mfrmEdit = Nothing
    ReDim marrRollList(0)
    Set mobjReport = New clsReport
    Set mrsDefine = InitAdviceDefine
    
    Call GetLocalSetting
    mblnAutoRead = Val(zlDatabase.GetPara("�Զ���Ǳ������״̬", glngSys, pסԺҽ���´�, "1", , , intType)) = 1
    mbln��������ִ�� = Val(zlDatabase.GetPara("������Ҫ����ִ��", glngSys)) = 1
    'ҽ����ӡģʽ
    mlngPrintType = Val(zlDatabase.GetPara("ҽ������ӡģʽ", glngSys, pסԺҽ���´�))
    'ת�Ƴ�Ժ��ӡ
    mlngPrintPos = Val(zlDatabase.GetPara("ת�ƺͳ�Ժ��ӡ", glngSys, pסԺҽ������, 1))
    
    mstr�����Ժ��� = zlDatabase.GetPara("Ҫ��������Ժ���", glngSys, pסԺҽ���´�)
    
    mblnAutoReadEnabled = Not ((intType = 3 Or intType = 15))
    mblnHaveAuditPriv = HaveAuditPriv
        
    If gblnKSSStrict Then Call CheckKSSPrivilege(1)
    If mint���� = 0 Then Call InitObjLis(pסԺҽ��վ)
    On Error Resume Next
    Set gobjExchange = CreateObject("zlExchange.clsExchange")
    If Not gobjExchange Is Nothing Then Call gobjExchange.Init(gcnOracle)
    err.Clear: On Error GoTo 0
End Sub

Private Sub GetLocalSetting()
'���ܣ���ȡ���ز���
    'ִ������
    mbln���� = Val(zlDatabase.GetPara("ҽ��ִ������", glngSys, pסԺҽ���´�)) <> 0
    'Ƥ����֤
    mblnƤ����֤ = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, pסԺҽ������)) <> 0
    '���뵥��ӡģʽ
    mint���뵥��ӡģʽ = Val(zlDatabase.GetPara("��Ѫ���뵥��ӡģʽ", glngSys, pסԺҽ������, "1"))

    mblnҽ����λ��� = Val(zlDatabase.GetPara("ҽ�����Ĭ�϶�λ�����һ��", glngSys, pסԺҽ���´�)) = 1
    
    mblnΣ��ֵ = InStr(GetInsidePrivs(pסԺҽ��վ), ";Σ��ֵ����;") > 0
    
    mbln��ʿǩ�� = Val(zlDatabase.GetPara("У��ҽ������ǩ��", glngSys, pסԺҽ������)) <> 0 And gintCA <> 0 And Mid(gstrESign, 2, 1) = "1"
    
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
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

    Set objBar = cbsSub.Add("�ڲ�������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 24, 24
    objBar.Visible = False  'ֻ���ڲ�����ʱ����ʾ(zlDefCommandBars)
    

    Set objBar = cbsSub.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    objBar.SetIconSize 16, 16
    With objBar.Controls
        Set objControl = .Add(xtpControlLabel, ID_ʱ���ǩ, "ʱ��")   'ҽ��ʱ��
        Set objCustom = .Add(xtpControlCustom, ID_ʱ��, "ʱ��")
            objCustom.Handle = cboTime.hwnd
        Set objControl = .Add(xtpControlButton, ID_����ҽ��, "����ҽ��")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_δ����ֹʱ��, "δ����ֹʱ��")
            objControl.ToolTipText = "��ʾδ��ִ����ֹʱ��ĳ���ҽ��"
        Set objControl = .Add(xtpControlButton, ID_����ҽ��, "����ҽ��")
        
        '----------------����ҳ��
        Set objControl = .Add(xtpControlButton, ID_ȫ��, "ȫ��")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.IconId = 1 '��ʼʱ����ͼ��
        Set objControl = .Add(xtpControlButton, ID_����, "����")
        Set objControl = .Add(xtpControlButton, ID_����, "����")
            objControl.IconId = 1
        '----------------
        
        Set objPopup = .Add(xtpControlPopup, ID_Ӥ��, "����ҽ��")
            objPopup.ID = ID_Ӥ��: objPopup.BeginGroup = True
            objPopup.IconId = 2608
        Set objControl = .Add(xtpControlButton, ID_����, "������")
            objControl.ToolTipText = "��ʾ���һ���������ҽ��"
        Set objControl = .Add(xtpControlButton, ID_δ����, "δ����")
            objControl.BeginGroup = True
            objControl.ToolTipText = "����ʾ������δ���ʵĻ��۷��õ�ҽ��"
        Set objControl = .Add(xtpControlButton, ID_����, "�����´�")
            objControl.ToolTipText = "ֻ��ʾҽ�������´��ҽ��"
            
        Set objControl = .Add(xtpControlButton, ID_�Ǳ���ҽ��, "��Ҫ����")
            objControl.ToolTipText = "��ʾ��Ҫ��д�����ҽ��,�Ͳ���Ҫ��������ѡ������ѡ��һ����"
            objControl.IconId = 11
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, ID_�Ǳ���ҽ��, "����Ҫ����")
            objControl.ToolTipText = "��ʾ����Ҫ��д�����ҽ��,����Ҫ��������ѡ������ѡ��һ����"
            objControl.IconId = 11
            
        Set objControl = .Add(xtpControlButton, ID_δ������, "δ������")
            objControl.ToolTipText = "��ʾδ������"
            objControl.BeginGroup = True
            mvarCond.δ������ = True
            
        Set objControl = .Add(xtpControlButton, ID_�ѳ�����, "�ѳ�����")
            objControl.ToolTipText = "��ʾ�ѳ�����"
            mvarCond.�ѳ����� = True
        Set objControl = .Add(xtpControlButton, ID_ҽ����ɫʾ��, "ҽ����ɫʾ��")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButton, ID_���, "���")
            objControl.BeginGroup = True
            objControl.Flags = xtpFlagRightAlign
        Set objControl = .Add(xtpControlButton, ID_��ϸ, "��ϸ")
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
    
    'ȱʡҽ��ʱ��
    cboTime.Clear
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "����"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "[ָ��..]"
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 16, 16
    End With
    Set cbsExec.Icons = zlCommFun.GetPubIcons
    cbsExec.EnableCustomization False
    cbsExec.ActiveMenuBar.Visible = False
    
    Set objBar = cbsExec.Add("������", xtpBarTop)
    objBar.EnableDocking xtpFlagHideWrap '+ xtpFlagStretched
    objBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objBar.ContextMenuPresent = False
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, ID_��ʾִ��, "��ʾִ������")
        Set objControl = .Add(xtpControlButton, ID_���ִ��, "ִ�����")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_Complete
        Set objControl = .Add(xtpControlButton, ID_ȡ�����, "ȡ�����")
            objControl.IconId = conMenu_Edit_Untread
        Set objControl = .Add(xtpControlButton, ID_ִ�м�¼, "��¼ִ�����")
            objControl.BeginGroup = True
            objControl.IconId = conMenu_Manage_ThingAdd
        Set objControl = .Add(xtpControlButton, ID_ִ�е���, "����ִ�����")
            objControl.IconId = conMenu_Manage_ThingModi
        Set objControl = .Add(xtpControlButton, ID_ִ��ɾ��, "ɾ��ִ�����")
            objControl.IconId = conMenu_Manage_ThingDel
        Set objControl = .Add(xtpControlButton, ID_�˶�, "�˶�")
            objControl.IconId = conMenu_Manage_ThingAudit
        Set objControl = .Add(xtpControlButton, ID_ȡ���˶�, "ȡ���˶�")
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
'���ܣ�����������İ���,���ڴ���ҽ�������ȼ�
'˵����
'1.��ҽ���Ӵ���δ����ʱ,�Ӵ���CommandBar���ȼ���Ч
'2.������CommandBar��KeyDown�¼������˵ļ������ټ�����¼�
    
    If Not Me.Visible Then Exit Sub '�������Ӵ���ʱ�Իἤ��
    If mlng����ID = 0 Then Exit Sub
    Call ActiveHotKey(KeyCode, Shift)
End Sub

Private Sub ActiveHotKey(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    Dim lngID As Long
    Dim intTab As Integer
    
    If Not Me.Visible Then Exit Sub
    If mlng����ID = 0 Then Exit Sub
    intTab = -1
    
    If Shift = vbCtrlMask And KeyCode >= vbKey0 And KeyCode <= vbKey5 Then
        lngID = ID_Ӥ�� * 100# + KeyCode - vbKey0 + 1
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
                lngID = ID_Ӥ�� * 100#
            Case vbKeyJ
                lngID = ID_����
            Case vbKeyK
                lngID = ID_����
            Case vbKeyU
                If mvarCond.����ģʽ = 3 Then
                    lngID = ID_ȫ��
                Else
                    lngID = ID_����ҽ��
                End If
            Case vbKeyX
                lngID = ID_���
            Case vbKeyY
                lngID = ID_����
            Case vbKeyQ
                lngID = ID_����
        End Select
    ElseIf KeyCode = vbKeyEscape Then '�ر���ѡ����
        If vsColumn.Visible Then
            vsColumn.Visible = False
            If vsAdvice.Enabled Then vsAdvice.SetFocus
        End If
    ElseIf Shift = vbAltMask And KeyCode = vbKeyC Then '����ѡ����
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
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "AppendData", IIF(mblnAppend, 1, 0)
    If mblnAppend And Not tbcAppend.Selected Is Nothing Then
        Call zlDatabase.SetPara("ҽ�����б�", tbcAppend.Selected.Tag, glngSys, pסԺҽ���´�)
    End If
    Call SaveFilterSetting
    Call SaveWinState(Me, App.ProductName)
    
    '��ҳ��������ֹ
    Call CreatePlugInOK(IIF(mint���� = 1, pסԺҽ������, pסԺҽ���´�), mint����)
    If Not gobjPlugIn Is Nothing Then
        On Error Resume Next
        If mint���� = 0 Then 'ҽ��վ����
            Call gobjPlugIn.Terminate(glngSys, pסԺҽ���´�, 0)
        ElseIf mint���� = 1 Then '��ʿվ����
            Call gobjPlugIn.Terminate(glngSys, pסԺҽ������, 1)
        ElseIf mint���� = 2 Then 'ҽ��վ����
            Call gobjPlugIn.Terminate(glngSys, pסԺҽ���´�, 2)
        End If
        Call zlPlugInErrH(err, "Terminate")
        err.Clear: On Error GoTo 0
    End If
    Set mclsMipModule = Nothing
    Set mrsΣ��ֵ = Nothing
    mblnΣ��ֵ = False
    mlngΣ��ֵID = 0
End Sub

Private Sub RefreshData()
'���ܣ�ˢ������
    If mlng����ID = 0 Then
        '���ҽ���嵥
        Call ClearAdviceData
        Call ClearAppendData
        mlngBabyDept = 0
    Else
        '��ʾҽ���嵥
        Call LoadAdvice
    End If
End Sub

Private Sub Refresh����()
'���ܣ��ڱ���ҳ�治ͬ����֮���л�ʱ�����ˢ�£������¶����ݿ����ñ������غ���ʾ
    Dim i As Long
    Dim blnTmp As Boolean
    Dim lngҽ��ID As Long
    Dim strFormat As String
    Dim strSameDay As String
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))      '��¼��ǰ��������ڵ�ǰ����ˢ��ҽ����Ӧ�ò���
        
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_������)) <> 0 Then
                If mvarCond.���� = 0 Then ' ȫ��
                    blnTmp = True
                ElseIf mvarCond.���� = 1 Then ' ���
                    blnTmp = .TextMatrix(i, COL_�������) = "D"
                ElseIf mvarCond.���� = 2 Then '����
                    blnTmp = (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "C")
                ElseIf mvarCond.���� = 3 Then ' ����
                    blnTmp = Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "D" Or .TextMatrix(i, COL_�������) = "C")
                End If
                
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                
                .RowHidden(i) = Not blnTmp
            Else
                .RowHidden(i) = True: .RowHeight(i) = 0
            End If
            
            '���ӹ���δ���ı�����ѳ��ı���
            If .RowHidden(i) = False Then
                blnTmp = IIF(.TextMatrix(i, COL_����״̬) = "δ��", mvarCond.δ������, mvarCond.�ѳ�����)
                If blnTmp And .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                .RowHidden(i) = Not blnTmp
            End If
        Next
    End With
    Call LocatedDefaultAdviceRow(lngҽ��ID)
End Sub

Private Sub LocatedDefaultAdviceRow(Optional ByVal lngҽ��ID As Long)
'���ܣ�ҽ���嵥��ȱʡ��λ�������ҽ��id����ҽ��id��λ
    'ȱʡ��λ����ǰѡ���ҽ��Ϊ��ʾ����λ������λ�����һ�С�
    Dim i As Long
    
    With vsAdvice
        .Redraw = flexRDNone
        .Row = .Rows - 1
        If lngҽ��ID <> 0 Then
            lngҽ��ID = .FindRow(CStr(lngҽ��ID), , COL_ID)
            If lngҽ��ID <> -1 Then
                If Not .RowHidden(lngҽ��ID) Then .Row = lngҽ��ID
            End If
        End If
        If mint���� = 1 Then
            If lngҽ��ID = -1 Or lngҽ��ID = 0 Then
                vsAdvice.Row = IIF(mblnҽ����λ���, vsAdvice.Rows - 1, vsAdvice.FixedRows)
            End If
        End If
        If .RowHidden(.Row) Then    '��λ���������еĴ���
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
'���ܣ���ȡҽ��������������
    Dim strPar As String
    
    mvarCond.Ӥ�� = 0
    mvarCond.δ���� = False
    mblnHideFilter = Val(zlDatabase.GetPara("���������Զ�����", glngSys, pסԺҽ���´�, "0")) <> 0
    mvarCond.���� = Val(zlDatabase.GetPara("����ҽ������", glngSys, pסԺҽ���´�, "0")) <> 0
    mvarCond.���� = Val(zlDatabase.GetPara("����ҽ������", glngSys, pסԺҽ���´�, "1")) <> 0
    
    strPar = Val(zlDatabase.GetPara("��ʾģʽ", glngSys, pסԺҽ���´�, "0"))
    mvarCond.��ʾģʽ = IIF(Val(strPar) = 0, 0, 1)
    
    mlngBaby = Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0")) - 1
    
    strPar = Val(zlDatabase.GetPara("ҽ�����˷�ʽ", glngSys, pסԺҽ���´�, "0"))
    If InStr(",0,1,2,3,", "," & strPar & ",") > 0 Then
        mvarCond.����ģʽ = Val(strPar)
    Else
        mvarCond.����ģʽ = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("����鿴����", glngSys, pסԺҽ���´�, "0"))
    If InStr(",0,1,2,3,", "," & strPar & ",") > 0 Then
        mvarCond.���� = Val(strPar)
    Else
        mvarCond.���� = 0
    End If
    
    strPar = Val(zlDatabase.GetPara("ҽ����ʾ����", glngSys, pסԺҽ���´�, "0"))
    mvarCond.ҽ����ʾ = IIF(Val(strPar) = 0, 0, 1)
    
    mvarCond.δ����ֹʱ�� = Val(zlDatabase.GetPara("ҽ����ʾ����δ����ֹʱ��", glngSys, pסԺҽ���´�, "1")) = 1
    
    strPar = Val(zlDatabase.GetPara("ҽ����ʾ������Ҫ", glngSys, pסԺҽ���´�, "0"))
    If strPar = "1" Then
        mvarCond.�Ǳ���ҽ�� = True: mvarCond.�Ǳ���ҽ�� = False
    ElseIf strPar = "2" Then
        mvarCond.�Ǳ���ҽ�� = False: mvarCond.�Ǳ���ҽ�� = True
    Else
        mvarCond.�Ǳ���ҽ�� = True: mvarCond.�Ǳ���ҽ�� = True
    End If
End Sub

Private Sub SaveFilterSetting()
'���ܣ�����ҽ��������������
    Dim strPar As String
    
    Call zlDatabase.SetPara("����ҽ������", IIF(mvarCond.����, 1, 0), glngSys, pסԺҽ���´�)
    Call zlDatabase.SetPara("����ҽ������", IIF(mvarCond.����, 1, 0), glngSys, pסԺҽ���´�)
    Call zlDatabase.SetPara("��ʾģʽ", mvarCond.��ʾģʽ, glngSys, pסԺҽ���´�)
    Call zlDatabase.SetPara("ҽ�����˷�ʽ", mvarCond.����ģʽ, glngSys, pסԺҽ���´�)
    Call zlDatabase.SetPara("����鿴����", mvarCond.����, glngSys, pסԺҽ���´�)
    Call zlDatabase.SetPara("���������Զ�����", IIF(mblnHideFilter, 1, 0), glngSys, pסԺҽ���´�)
    Call zlDatabase.SetPara("ҽ����ʾ����", mvarCond.ҽ����ʾ, glngSys, pסԺҽ���´�)
    Call zlDatabase.SetPara("ҽ����ʾ����δ����ֹʱ��", IIF(mvarCond.δ����ֹʱ��, 1, 0), glngSys, pסԺҽ���´�)
    
    If mvarCond.�Ǳ���ҽ�� And Not mvarCond.�Ǳ���ҽ�� Then
        strPar = "1"
    ElseIf Not mvarCond.�Ǳ���ҽ�� And mvarCond.�Ǳ���ҽ�� Then
        strPar = "2"
    Else
        strPar = "0"
    End If
    Call zlDatabase.SetPara("ҽ����ʾ������Ҫ", strPar, glngSys, pסԺҽ���´�)
End Sub

Private Sub ClearAppendData()
'���ܣ�������ӱ������븽�������
    Dim blnSel As Boolean, intIdx As Integer
    Dim varDraw As RedrawSettings
    
    If vsAppend.FixedRows = 2 Then vsAppend.RemoveItem 0
    vsAppend.Rows = vsAppend.FixedRows
    vsAppend.Rows = vsAppend.FixedRows + 1
    vsAppend.Row = vsAppend.FixedRows
    
    If rtfAppend.Visible Then rtfAppend.Text = ""
    If rtfInfo.Visible Then rtfInfo.Text = ""
    
    For intIdx = 0 To tbcAppend.ItemCount - 1
        If InStr("����,����,��ҩ,ԤԼ,����,ѪҺ", tbcAppend(intIdx).Tag) > 0 Then
            If tbcAppend(intIdx).Selected Then blnSel = True
            tbcAppend(intIdx).Visible = False
        End If
    Next
   
    If blnSel Then
        varDraw = vsAdvice.Redraw '�������������ظ�����
        vsAdvice.Redraw = flexRDNone
        tbcAppend.Item(0).Selected = True
        vsAdvice.Redraw = varDraw
    End If
    
    ReDim marrRollList(0)
End Sub

Private Sub InitPriceTable()
'���ܣ���ʼ���Ƽ��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "�Ƽ�ҽ��,2000,1;���,650,1;�շ���Ŀ,2500,1;��λ,500,4;�Ƽ�����,850,1;����,900,7;ִ�п���,1000,1;��������,800,1;����,450,4;�շѷ�ʽ,1500,1"
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
'���ܣ���ʼ�������嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "���ͺ�;����ʱ��,1530,1;���ݺ�,850,1;����ҽ��,1800,1;�շ���Ŀ,1800,1;��������,850,1;�Ʒ�״̬,850,1;" & _
        "ִ��״̬,850,1;״̬˵��,1800,1;ִ�п���,1000,1;ִ����,800,1;ִ��ʱ��,1530,1;���ִ��ʱ��,1530,1;ִ��˵��,1800,1;�״�ʱ��,1530,1;ĩ��ʱ��,1530,1;������,800,1;ҽ��ID;���ID;��¼����;�������;��¼״̬;�������;��������;��������;���ʱ��;��Ѫ����;����ʱ��,1530,1;������,800,1;����ʱ��,1530,1"
    arrHead = Split(strHead, ";")
    With vsAppend
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1  '��ʽ������vsAppend_AfterRowColChange
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
        
        .MergeCells = flexMergeRestrictAll  '�Զ�����MergeCellsFixedΪ��ͬ��ʽ
        .MergeCompare = flexMCIncludeNulls
    End With
End Sub

Private Sub InitExecTable()
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
    
    strHead = "Ҫ��ʱ��,1530,1;ִ��ʱ��,1530,1;��������,850,1;ִ��ժҪ,2500,1;ִ����,750,1;�Ǽ�ʱ��,1530,1;�Ǽ���,750,1;ִ�н��,1000,1;�˶���,750,1;�˶�ʱ��,1530,1;˵��,500,1;��Դ,600,1"
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
'���ܣ���ʼ��ǩ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long, blnDo As Boolean
    
    strHead = "ǩ������,1150,1;ǩ��ʱ��,1900,1;ǩ����,800,1;ʱ���,1900,1"
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
'���ܣ����ҽ���嵥����
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Editable = flexEDNone
End Sub

Private Sub InitColumnSelect()
'���ܣ�����ҽ���嵥ԭʼ����ʾ״̬��ʼ����ѡ����
    Dim lngRow As Long, i As Long
    
    vsColumn.Rows = vsColumn.FixedRows
    With vsAdvice
        For i = .FixedCols To .Cols - 1
            If Not (.ColHidden(i) Or .ColWidth(i) = 0) Then
                If .TextMatrix(0, i) <> "" And Not (i = COL_����״̬ Or i = COL_�걾״̬) Then  '�����,Ƥ��
                    vsColumn.Rows = vsColumn.Rows + 1
                    lngRow = vsColumn.Rows - 1
                    vsColumn.TextMatrix(lngRow, 1) = .TextMatrix(0, i)
                    vsColumn.RowData(lngRow) = i
                    
                    '�̶���ʾ��
                    If InStr(",��ʼʱ��,ҽ������,����ҽ��,", "," & .TextMatrix(0, i) & ",") > 0 Then
                        vsColumn.TextMatrix(lngRow, 0) = 1
                        vsColumn.Cell(flexcpForeColor, lngRow, 0, lngRow, 1) = vsColumn.BackColorFixed
                    End If
                    
                    'Ĭ�����ؿ���ʱ��
                     If InStr(",����ʱ��,", "," & .TextMatrix(0, i) & ",") > 0 Then
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
'���ܣ���ʼ��ҽ���嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "ID;���ID;���;Ӥ��ID;ҽ��״̬;�������;��������;�������;��־;" & _
              ",240,4;��Ч,500,4;��Чʱ��,1530,1;,200,7;ҽ������,3000,1;����,4000,1;,375,1;����,850,1;����,850,1;����,450,1;Ƶ��,1000,1;�÷�,1000,1;ҽ������,1000,1;ִ��ʱ��,1000,1;" & _
              "��ֹʱ��,1530,1;ִ�п���,1000,1;ִ������,850,1;�ϴ�ִ��,1560,1;״̬,500,4;����ҽ��,850,1;����ʱ��,1530,1;У�Ի�ʿ,850,1;У��ʱ��,1530,1;ͣ��ҽ��,850,1;" & _
              "ͣ��ʱ��,1530,1;ͣ����ʿ,850,1;ȷ��ͣ��ʱ��,1530,1;����ҩ��,850,1;����״̬,700,4;�걾״̬,850,1;" & _
              "������ĿID;�Թܱ���;ִ�б��;���δ�ӡ;ǰ��ID;ǩ����;�ļ�ID;������;����ID;�շ�ϸĿID;������λ;��������ID;���״̬;�������;" & _
              "��˱��;��ΣҩƷ;�걾��λ;��ҩĿ��;��鱨��ID;�������״̬;���������;RISԤԼID;RIS����ID;LIS����ID;RISԤԼ״̬;������Ŀ����;��鷽��;Σ��ֵID;�׵���"
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
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .ColData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        'δ���ú�����ҩʱ�����в��ɼ�������������̫Ԫͨʱ������gbytPass=1 or 3 ʱ �ɼ�
        .ColHidden(COL_��ʾ) = True
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .ColWidth(COL_F��־) = 11 * Screen.TwipsPerPixelX
        .ColWidth(COL_F����) = 11 * Screen.TwipsPerPixelX
    End With
End Sub

Private Sub SetAdviceColVisible()
'���ܣ�����ҽ������еĿɼ��Ժͱ�ͷ����
    Dim i As Long
    
    '������ʾģʽ������ʾ��
    With vsAdvice
        If (mvarCond.����ģʽ = 1 Or mvarCond.����ģʽ = 2) And mvarCond.��ʾģʽ = 0 Then
            .ColHidden(COL_��Ч) = True
        Else
            .ColHidden(COL_��Ч) = False
        End If
        
        .ColHidden(col_ҽ������) = mvarCond.��ʾģʽ = 0
        .ColHidden(col_����) = mvarCond.��ʾģʽ = 1
        .ColHidden(COL_Ƥ��) = False
        .ColHidden(COL_����) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_����) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_����) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_Ƶ��) = mvarCond.��ʾģʽ = 0
        .ColHidden(COL_ִ��ʱ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ִ��ʱ��) = "Detail"
        .ColHidden(COL_ִ������) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ִ������) = "Detail"
        .ColHidden(COL_�ϴ�ִ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_�ϴ�ִ��) = "Detail"
        .ColHidden(COL_״̬) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_״̬) = "Detail"
        .ColHidden(COL_����ʱ��) = True
        .ColHidden(COL_У�Ի�ʿ) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_У�Ի�ʿ) = "Detail"
        .ColHidden(COL_У��ʱ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_У��ʱ��) = "Detail"
        .ColHidden(COL_ͣ��ҽ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ͣ��ҽ��) = "Detail"
        .ColHidden(COL_ͣ��ʱ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ͣ��ʱ��) = "Detail"
        .ColHidden(COL_ͣ����ʿ) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ͣ����ʿ) = "Detail"
        .ColHidden(COL_ȷ��ͣ��ʱ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_ȷ��ͣ��ʱ��) = "Detail"
        .ColHidden(COL_����ҩ��) = mvarCond.��ʾģʽ = 0: .Cell(flexcpData, 0, COL_����ҩ��) = "Detail"
        .ColHidden(COL_��ΣҩƷ) = True
        .ColHidden(COL_�걾��λ) = True
        .ColHidden(COL_��˱��) = True
        .ColHidden(COL_��ҩĿ��) = True
        .ColHidden(COL_��鱨��ID) = True
        .ColHidden(COL_�������״̬) = True
        .ColHidden(COL_���������) = True
        .ColHidden(COL_��) = True
        .ColHidden(COL_�걾״̬) = True
        
        If mvarCond.����ģʽ = 3 Then '���Ǳ��濨Ƭ�Ȳ�����ʾ
            For i = COL_��ʼʱ�� + 1 To COL_�걾��λ
                .ColHidden(i) = True
            Next
            .ColHidden(COL_��Ч) = True
            .ColHidden(COL_��ʼʱ��) = False
            .ColHidden(col_����) = False
            .ColHidden(COL_ִ�п���) = False
            .ColHidden(COL_����ҽ��) = False
            .TextMatrix(0, COL_����ҽ��) = "����ҽ��"
            .ColHidden(COL_����״̬) = mfrmParent Is Nothing    '���Ӳ�������δ����������,��ֹ��ʾ����״̬
            .ColWidth(COL_����״̬) = 700
            .TextMatrix(0, COL_����״̬) = "����"
            .ColHidden(COL_�걾״̬) = False
            .ColWidth(COL_�걾״̬) = 850
        Else
            .ColHidden(COL_��) = False
            .TextMatrix(0, COL_����ҽ��) = "����ҽ��"
            If mvarCond.����ģʽ = 0 And mvarCond.��ʾģʽ = 0 Then .ColHidden(COL_��Ч) = False
            .ColHidden(COL_�÷�) = False
            .ColHidden(COL_ҽ������) = False
            .ColHidden(COL_��ֹʱ��) = (mvarCond.��ʾģʽ = 0 And mvarCond.����ģʽ = 2)
            .ColHidden(COL_����״̬) = True
            .TextMatrix(0, COL_����״̬) = "����״̬"
        End If
        'ֻ�г���ʱ����������
        If mvarCond.��ʾģʽ = 1 Then .ColHidden(COL_����) = mvarCond.����ģʽ = 1 Or Not mbln����
    End With
End Sub

Private Function LoadAdvice(Optional ByVal blnRefreshNotify As Boolean) As Boolean
'���ܣ����ݵ�ǰ�������ö�ȡ����ʾҽ���嵥
'������blnRefreshNotifyˢ��ҽ������(F5�ֶ�ˢ��,�¿�ҽ����ֹͣҽ��������ҽ��ʱ)
    Dim rsTmp As ADODB.Recordset
    Dim rsѪ�� As ADODB.Recordset
    Dim strSQL As String, strWhere As String
    Dim i As Long, j As Long
    Dim strFormat As String, strTmp As String
    Dim bln��ҩ;�� As Boolean, bln��ҩ�÷� As Boolean
    Dim bln�ɼ����� As Boolean, bln��Ѫ;�� As Boolean, blnFirst As Boolean
    Dim str״̬SQL As String, lngҽ��ID As Long
    Dim strδ���� As String
    Dim blnDo As Boolean, strCurr As String, strTime As String
    Dim strҽ����Ч As String, strҽ��״̬ As String
    Dim strSameDay As String, strGroupBy As String
    Dim strPreDay1 As String, strPreDay2 As String
    If mlng����ID = 0 Then Exit Function

    Screen.MousePointer = 11

    On Error GoTo errH
    
    lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))    '��¼��ǰ��������ڵ�ǰ����ˢ��ҽ����Ӧ�ò���
    If mvarCond.ҽ��ID <> 0 And lngҽ��ID = 0 Then
        lngҽ��ID = mvarCond.ҽ��ID
    End If

    'ҽ����������
    If mstrӤ�� <> "" Then
        If mblnFirstBaby = False Then
            mvarCond.Ӥ�� = mlngBaby
            mbytӤ�� = IIF(mvarCond.Ӥ�� = -1, 0, mvarCond.Ӥ��)
            Call zlDatabase.SetPara("����Ӥ������", mvarCond.Ӥ��, glngSys, pסԺҽ���´�)
            mblnFirstBaby = True
        End If

        'ĸӤ����Ĵ���
        If mlngBabyDept <> mlngӤ������ID Then
            If mlngӤ������ID <> 0 Then
                If (mvarCond.Ӥ�� = -1 Or mvarCond.Ӥ�� = 0) And (mlngӤ������ID = mlngҽ������ID Or mlngӤ������ID = mlngҽ������ID) Then
                    'Ӥ�����Ҳ���Ĭ��ѡ��Ӥ��
                    mvarCond.Ӥ�� = 1: mbytӤ�� = mvarCond.Ӥ��
                ElseIf (mvarCond.Ӥ�� = -1 Or mvarCond.Ӥ�� = 1) And (mlng����ID = mlngҽ������ID Or mlng����ID = mlngҽ������ID) Then
                    '���˿��Ҳ���Ĭ��ѡ����
                    mvarCond.Ӥ�� = 0: mbytӤ�� = mvarCond.Ӥ��
                End If
            End If
            mlngBabyDept = mlngӤ������ID
        End If
    Else
        mlngBabyDept = 0
        mblnFirstBaby = False
    End If
    strWhere = ""
    If mvarCond.Ӥ�� <> -1 Then
        strWhere = strWhere & " And Nvl(A.Ӥ��,0)=[4]"
    End If
        
    If mvarCond.����ģʽ = 1 Then
        strWhere = strWhere & " And A.ҽ����Ч=0"
    ElseIf mvarCond.����ģʽ = 2 Then
        strWhere = strWhere & " And A.ҽ����Ч=1"
    End If
    
    If mvarCond.ҽ����ʾ = 0 And mvarCond.����ģʽ <> 3 Then  '����ҽ����������ڱ���ҳ�治������ҽ���ķ�Χ��ֱ��ȡ��������˵����б��档
        strWhere = strWhere & " And Nvl(A.ҽ��״̬,0)<>4  And (A.ҽ����Ч=0 and " & _
            IIF(mvarCond.δ����ֹʱ��, " (a.ִ����ֹʱ��>[3] or a.ִ����ֹʱ�� is null) ", " a.ִ����ֹʱ�� is null ") & _
            " or A.ҽ����Ч=1 and A.��ʼִ��ʱ�� >=[6])"
    End If
    
    If Not (mvarCond.����ģʽ <> 3 And mvarCond.ҽ����ʾ = 0) Then
        If mvarCond.��ʼʱ�� <> CDate(0) And mvarCond.����ʱ�� <> CDate(0) Then
            strWhere = strWhere & " And A.����ʱ��+0 Between [7] And [8]"
        End If
    End If
    
    'ֻ��ʾ����δ���ʷ��õ�ҽ��
    If mvarCond.δ���� Then
        strδ���� = _
        " And Exists" & vbNewLine & _
                 " (Select 1" & vbNewLine & _
                 "       From (Select Nvl(C.���id, C.ID) As ҽ��id" & vbNewLine & _
                 "              From ����ҽ������ A, סԺ���ü�¼ B, ����ҽ����¼ C" & vbNewLine & _
                 "              Where A.ҽ��id = C.ID And A.NO = B.NO And A.��¼���� = B.��¼���� And A.��¼���� = 2 And B.��¼״̬ = 0 And" & vbNewLine & _
                 "                    C.����id = [1] And C.��ҳid = [2]" & IIF(mvarCond.Ӥ�� <> -1, " And Nvl(C.Ӥ��, 0) = [4]", "") & ")" & vbNewLine & _
                 "       Where A.ID = ҽ��id Or A.���id = ҽ��id)"
    End If
    
    'ҽ��վ  �����´�
    If mlngǰ��ID <> 0 And mvarCond.���� Then
        strWhere = strWhere & " And Nvl(A.ǰ��ID,0)<>0 and (A.ǰ��ID in (Select /*+cardinality(x,10)*/ x.Column_Value From Table(Cast(f_Num2list([5]) As zlTools.t_Numlist)) X) or a.��������ID=[9])"
    End If
    
    'ҽ����¼��������������,��������,��鲿λ,��ҩ�巨'�������÷�����
    str״̬SQL = "Decode(A.ҽ��״̬,1,'�¿�',2,'����',3,'У��',4,'����',5,'����',6,'��ͣ',7,'����',8,'ֹͣ',9,'ȷ��ֹͣ')"
    strSQL = _
        " Select /*+ RULE */ A.ID,A.���ID,A.���,Nvl(A.Ӥ��,0) as Ӥ��ID,A.ҽ��״̬,Nvl(A.�������,'*') as �������,B.��������,C.�������,A.������־ as ��־,A.����� as ��ʾ," & _
        " Decode(Nvl(A.ҽ����Ч,0),0,'����','����') as ��Ч,To_Char(A.��ʼִ��ʱ��,'YYYY-MM-DD HH24:MI') as ��ʼʱ��,Null as ��,A.ҽ������,Null as ����,A.Ƥ�Խ�� as Ƥ��," & _
        " Decode(A.�ܸ�����,NULL,NULL,Decode(A.�������,'E',Decode(B.��������,'4',A.�ܸ�����||'��',A.�ܸ�����||B.���㵥λ),'4',A.�ܸ�����||G.���㵥λ,'5',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,'6',Round(A.�ܸ�����/D.סԺ��װ,5)||D.סԺ��λ,A.�ܸ�����||B.���㵥λ)) as ����," & _
        " Decode(A.�״�����,Null,'',A.�״�����||Decode(A.�������,'4',G.���㵥λ,B.���㵥λ)||':')||Decode(A.��������,NULL,NULL,decode(sign(1-A.��������),1,'0'||A.��������,A.��������)||Decode(A.�������,'4',G.���㵥λ,B.���㵥λ)) as ����," & _
        " A.����,A.ִ��Ƶ�� as Ƶ��,Decode(A.�������,'E',Decode(Instr('2468',Nvl(B.��������,'0')),0,NULL,B.����),NULL) as �÷�,A.ҽ������,A.ִ��ʱ�䷽�� as ִ��ʱ��," & _
        " To_Char(A.ִ����ֹʱ��,'YYYY-MM-DD HH24:MI') as ��ֹʱ��,Nvl(E.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'-')) as ִ�п���," & _
        " Decode(Instr('567E',Nvl(A.�������,'*')),0,NULL,A.ִ������) as ִ������,To_Char(A.�ϴ�ִ��ʱ��,'YYYY-MM-DD HH24:MI') as �ϴ�ִ��," & str״̬SQL & " as ״̬," & _
        " A.����ҽ��,To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as ����ʱ��,A.У�Ի�ʿ,To_Char(A.У��ʱ��,'YYYY-MM-DD HH24:MI') as У��ʱ��,A.ͣ��ҽ��," & _
        " To_Char(A.ͣ��ʱ��,'YYYY-MM-DD HH24:MI') as ͣ��ʱ��,A.ȷ��ͣ����ʿ as ͣ����ʿ,To_Char(A.ȷ��ͣ��ʱ��,'YYYY-MM-DD HH24:MI') as ȷ��ͣ��ʱ��,D.����ҩ��,D.�Ƿ���������,Decode(Max(NVL(y.����״̬,0)),Min(NVL(y.����״̬,0)),Max(NVL(y.����״̬,0)),2) As ����״̬,null as �걾״̬,A.������ĿID," & _
        " B.�Թܱ���,A.ִ�б��,A.���δ�ӡ,A.ǰ��ID,Decode(A.�¿�ǩ��ID,NULL,0,1) as ǩ����,M.�����ļ�ID as �ļ�ID,Nvl(N.ͨ��,0) as ������,Max(y.����id) As ����id," & _
        " A.�շ�ϸĿID,B.���㵥λ as ������λ,A.��������ID,A.���״̬,A.�������," & _
        " A.��˱��,d.��ΣҩƷ,A.�걾��λ,A.��ҩĿ�� ,Max(y.��鱨��id)||'' As ��鱨��id,J.״̬ as �������״̬,J.����� as ���������,f.ԤԼID as RISԤԼID,Max(y.RISID) As RIS����ID,Max(y.����ID) as LIS����ID,f.�Ƿ���� as RISԤԼ״̬,b.���� as ������Ŀ����,Max(a.��鷽��) as ��鷽��,max(h.Σ��ֵid) as Σ��ֵID"
    strSQL = strSQL & _
        " From ����ҽ����¼ A,���ű� E,ҩƷ���� C,ҩƷ��� D,������ĿĿ¼ B,�շ���ĿĿ¼ G,����ҽ������ Y,��������Ӧ�� M,�����ļ��б� N,���������ϸ I,��������¼ J,RIS���ԤԼ F,����Σ��ֵҽ�� H" & _
        " Where A.������ĿID=B.ID(+) And A.ִ�п���ID=E.ID(+) And A.������ĿID=C.ҩ��ID(+) And a.ID = i.ҽ��ID(+) And I.��ID = J.ID(+) and (I.����ύ =1 Or I.��ID is NULL) and a.id=f.ҽ��ID(+) and a.id=h.ҽ��ID(+)" & _
        " And A.�շ�ϸĿID=D.ҩƷID(+) And A.�շ�ϸĿID=G.ID(+) And A.ID=Y.ҽ��ID(+) And (Not(A.������� IN ('F','G','D','E') And A.���ID is Not NULL) Or A.�������='E' And B.��������='8')" & _
        " And A.������ĿID=M.������ĿID(+) And M.Ӧ�ó���(+)=2 And M.�����ļ�ID=N.ID(+) And N.����(+)=7 And A.����ID=[1] And A.��ҳID=[2] And A.��ʼִ��ʱ�� is Not NULL And Nvl(A.ҽ��״̬,0)<>-1" & _
        IIF(mint���� = 2, "", " And A.������Դ<>3") & strWhere & strδ����
    strGroupBy = _
        " Group By a.Id,a.���id,a.���,a.Ӥ��,a.ҽ��״̬,a.�������,b.��������,c.�������,a.������־,a.�����,a.ҽ����Ч,a.��ʼִ��ʱ��,a.ҽ������,a.Ƥ�Խ��," & _
        " a.�ܸ�����,a.�״�����,g.���㵥λ,d.סԺ��װ,d.סԺ��λ,a.��������,a.����,a.ִ��Ƶ��,a.ҽ������,b.����,a.ִ������,a.ִ��ʱ�䷽��,a.ִ����ֹʱ��,e.����,a.�ϴ�ִ��ʱ��," & _
        " a.����ʱ��,a.����ҽ��,a.У�Ի�ʿ,a.У��ʱ��,a.ͣ��ҽ��,a.ͣ��ʱ��,a.ȷ��ͣ����ʿ,a.ȷ��ͣ��ʱ��,a.������Ŀid,b.�Թܱ���,a.ִ�б��,a.���δ�ӡ,a.ǰ��id,a.�¿�ǩ��id," & _
        " m.�����ļ�id,n.ͨ��,a.�շ�ϸĿid,b.���㵥λ,a.��������id,a.���״̬,a.�������,a.��˱��,d.����ҩ��,d.��ΣҩƷ,a.�걾��λ,a.��ҩĿ��,J.״̬,J.�����,f.ԤԼID,f.�Ƿ����,b.����,D.�Ƿ���������"
    '������ʾ��ʽ����
    If mdat���� <> CDate("1900-01-01") Then
        If mvarCond.���� Then
            'ֻ��ʾ���һ������֮���ҽ��
            strSQL = strSQL & " And (Nvl(A.������־,1)=1 Or A.ҽ��״̬ IN(1,2)) " & strGroupBy & " Order by Ӥ��ID,���"
        Else
            '��ʾ����ǰ��ָ�
            strSQL = _
                " Select * From (" & strSQL & " And Nvl(A.������־,1)=0 And A.ҽ��״̬ Not IN(1,2) " & strGroupBy & " Order by Ӥ��ID,���)" & _
                " Union ALL" & _
                " Select -Null as ID,-Null as ���ID,-Null as ���,-Null as Ӥ��ID,-Null as ҽ��״̬,Null as �������,Null as ��������,Null as �������,-Null as ��־,-Null as ��ʾ," & _
                " Null as ��Ч,Null as ��ʼʱ��,Null as ��,Null as ҽ������,Null as ����,Null as Ƥ��,Null as ����,Null as ����,-Null as ����,Null as Ƶ��,Null as �÷�,Null as ҽ������,Null as ִ��ʱ��," & _
                " Null as ��ֹʱ��,Null as ִ�п���,Null as ִ������,Null as �ϴ�ִ��,Null as ״̬,Null as ����ҽ��,Null as ����ʱ��,Null as У�Ի�ʿ,Null as У��ʱ��,Null as ͣ��ҽ��," & _
                " Null as ͣ��ʱ��,Null as ͣ����ʿ,Null as ȷ��ͣ��ʱ��,Null as ����ҩ��,-Null as �Ƿ���������,-Null as ����״̬,null as �걾״̬,-Null as ������ĿID,Null as �Թܱ���,-Null as ִ�б��,-Null as ���δ�ӡ,-Null as ǰ��ID,-Null as ǩ����,-Null as �ļ�ID," & _
                " -Null as ������,-Null as ����ID,-Null as �շ�ϸĿID, Null as ������λ, -Null as ��������ID,-Null as ���״̬,-Null as �������,-Null as ��˱��," & _
                " -Null as ��ΣҩƷ,-Null as �걾��λ,-NULL AS ��ҩĿ��,-NULL AS ��鱨��id,-Null as �������״̬,-Null as ���������,-null as RISԤԼID,-null as RIS����ID,-null as LIS����ID,-null as RISԤԼ״̬,null as ������Ŀ����,-null as ��鷽��,-null as Σ��ֵID From Dual" & _
                " Union ALL" & _
                " Select * From (" & strSQL & " And (Nvl(A.������־,1)=1 Or A.ҽ��״̬ IN(1,2)) " & strGroupBy & " Order by Ӥ��ID,���)"
        End If
    Else
        strSQL = strSQL & strGroupBy & " Order by Ӥ��ID,���"
    End If

    '������ʷ�ռ䴦��
    If mblnMoved Then
        strSQL = Replace(strSQL, "/*+ RULE */", "/*+driving_site(a) driving_site(y)*/")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    strCurr = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mlng��ҳID, CDate(strCurr), mvarCond.Ӥ��, IIF(mstrǰ��IDs = "", "0", mstrǰ��IDs), CDate(Format(strCurr, "yyyy-MM-dd 00:00:00")), mvarCond.��ʼʱ��, mvarCond.����ʱ��, mlng�������ID)
    
    If Not rsTmp.EOF Then
        strSQL = "Select a.ҽ��id,decode(a.��ѪѪ��,1,'A',2,'B',3,'AB',4,'O','') As Ѫ�� From ��Ѫ�����¼ A, ����ҽ����¼ B Where ҽ��id = b.Id And b.����ID=[1] and b.��ҳID=[2] And a.��ѪѪ��>0 and b.�������='K'"
        Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Name, mlng����ID, mlng��ҳID)
        
        With vsAdvice
            .Redraw = False
            .MergeCells = flexMergeNever
            Call ClearAdviceData
            Call AddDataToVsf(rsTmp)
            '����ÿ��ҽ��
            i = .FixedRows
            Do While i <= .Rows - 1
                .Cell(flexcpData, i, COL_��ʼʱ��) = CStr(.TextMatrix(i, COL_��ʼʱ��))    '������ҩ�ӿڵ���ʱȡ��
                .Cell(flexcpData, i, COL_����״̬) = Val(.TextMatrix(i, COL_����״̬)) '�������״ֵ̬
                If mvarCond.��ʾģʽ = 0 Then
                    '���ģʽ�´������ڵ���ʾ
                    strFormat = Format(.TextMatrix(i, COL_��ʼʱ��), "yyyy-MM-dd")
                    If strFormat = Format(strCurr, "yyyy-MM-dd") Then
                        .TextMatrix(i, COL_��ʼʱ��) = "�� �� " & Format(.TextMatrix(i, COL_��ʼʱ��), "HH:mm")
                    Else
                        If strFormat = strPreDay1 Then
                            .TextMatrix(i, COL_��ʼʱ��) = "�� �� " & Format(.TextMatrix(i, COL_��ʼʱ��), "HH:mm")
                        ElseIf strFormat = strPreDay2 Then
                            .TextMatrix(i, COL_��ʼʱ��) = "ǰ �� " & Format(.TextMatrix(i, COL_��ʼʱ��), "HH:mm")
                        Else
                            .TextMatrix(i, COL_��ʼʱ��) = Format(.TextMatrix(i, COL_��ʼʱ��), "MM-dd HH:mm")
                        End If
                    End If
                End If
                
                If .TextMatrix(i, COL_�������) = "K" And gblnѪ��ϵͳ Then
                    strSQL = "select zl_Get_��Ѫִ��Ѫ��([1]) as Ѫ�� from dual"
                    Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(i, COL_ID)))
                    If Not rsѪ��.EOF Then
                        If rsѪ��!Ѫ�� & "" <> "" Then .TextMatrix(i, COL_Ƥ��) = "(" & rsѪ��!Ѫ�� & ")"
                    End If
                End If
                
                '��ҩ����ҩ��һЩ����
                bln��ҩ;�� = False: bln��ҩ�÷� = False: bln�ɼ����� = False: bln��Ѫ;�� = False
                If .TextMatrix(i, COL_�������) = "E" Then
                    If Val(.TextMatrix(i - 1, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                        If InStr(",5,6,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ;�� = True
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    '��ʾ��ҩ�ĸ�ҩ;��+����
                                    .TextMatrix(j, COL_�÷�) = .TextMatrix(i, COL_�÷�) & .TextMatrix(i, COL_ҽ������)

                                    If mvarCond.��ʾģʽ = 0 Then    '�ϲ��÷���:�÷� Ƶ�� ����
                                        strFormat = .TextMatrix(j, COL_�÷�)
                                        strTmp = .TextMatrix(j, COL_Ƶ��)
                                        If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                                        strTmp = .TextMatrix(j, COL_����)
                                        If strTmp <> "" Then
                                            strFormat = strFormat & IIF(strFormat <> "", ",", "") & "��" & strTmp & "��"
                                        End If
                                        .TextMatrix(j, COL_�÷�) = strFormat
                                    End If

                                    '��ʾ��ҩ��ִ������
                                    If Val(.TextMatrix(j, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                        If Val(.TextMatrix(j, COL_ִ�б��)) = 2 Then
                                            .TextMatrix(j, COL_ִ������) = "��ȡҩ"
                                        Else
                                            .TextMatrix(j, COL_ִ������) = "�Ա�ҩ"
                                        End If
                                    ElseIf Val(.TextMatrix(j, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                        .TextMatrix(j, COL_ִ������) = "��Ժ��ҩ"
                                    Else
                                        .TextMatrix(j, COL_ִ������) = IIF(Val(.TextMatrix(j, COL_ִ�б��)) = 1, "��ȡҩ", "")
                                    End If
                                    
                                    'Σ��ֵID��ֻ��������ҽ�����ģ����Ƶ�ҩƷ����
                                    .TextMatrix(j, COL_Σ��ֵID) = .TextMatrix(i, COL_Σ��ֵID)

                                    If mvarCond.��ʾģʽ = 0 Then
                                        If .TextMatrix(j, COL_Ƥ��) <> "" Then
                                            If Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "1") Then
                                                .TextMatrix(j, col_����) = .TextMatrix(j, col_����) & "," & .TextMatrix(j, COL_Ƥ��)
                                            End If
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        ElseIf InStr(",7,C,", .TextMatrix(i - 1, COL_�������)) > 0 Then
                            bln��ҩ�÷� = .TextMatrix(i - 1, COL_�������) = "7"    '��ҩ�÷���
                            bln�ɼ����� = .TextMatrix(i - 1, COL_�������) = "C"    '�ɼ�������

                            '�ɼ���ʽ�Ĺ�����һ���ĵ�һ��������ͬ
                            If bln�ɼ����� Then
                                j = .FindRow(.TextMatrix(i, COL_ID), .FixedRows, COL_���ID)
                                If j <> -1 Then
                                    .TextMatrix(i, COL_�Թܱ���) = .TextMatrix(j, COL_�Թܱ���)
                                End If
                                .TextMatrix(i, COL_��ʼʱ��) = .TextMatrix(j, COL_��ʼʱ��)
                                .Cell(flexcpData, i, COL_��ʼʱ��) = CStr(.TextMatrix(j, COL_��ʼʱ��))
                                .Cell(flexcpData, i, COL_Ƥ��) = .TextMatrix(i, COL_Ƥ��)
                                .TextMatrix(i, COL_Ƥ��) = "" '���������ʱ��ID�����ϲ���ʾ
                            End If

                            '��ʾ��ҩ�䷽�������ϵ�ִ�п���
                            .TextMatrix(i, COL_ִ�п���) = .TextMatrix(i - 1, COL_ִ�п���)

                            If bln��ҩ�÷� Then
                                '��ʾ��ҩ�䷽ִ������
                                If Val(.TextMatrix(i - 1, COL_ִ������)) = 5 And Val(.TextMatrix(i, COL_ִ������)) <> 5 Then
                                    If Val(.TextMatrix(i - 1, COL_ִ�б��)) = 2 Then
                                        .TextMatrix(i, COL_ִ������) = "��ȡҩ"
                                    Else
                                        .TextMatrix(i, COL_ִ������) = "�Ա�ҩ"
                                    End If
                                ElseIf Val(.TextMatrix(i - 1, COL_ִ������)) <> 5 And Val(.TextMatrix(i, COL_ִ������)) = 5 Then
                                    .TextMatrix(i, COL_ִ������) = "��Ժ��ҩ"
                                Else
                                    .TextMatrix(i, COL_ִ������) = IIF(Val(.TextMatrix(i - 1, COL_ִ�б��)) = 1, "��ȡҩ", "")
                                End If
                            Else
                                .TextMatrix(i, COL_ִ������) = ""
                            End If

                            'ɾ����ζ��ҩ��,�Լ���������еļ�����Ŀ
                            strTmp = ""
                            For j = i - 1 To .FixedRows Step -1
                                If Val(.TextMatrix(j, COL_���ID)) = Val(.TextMatrix(i, COL_ID)) Then
                                    .TextMatrix(i, COL_������) = .TextMatrix(j, COL_������)    '���顢�䷽������ҽ��Ϊ׼
                                    .TextMatrix(i, COL_�ļ�ID) = .TextMatrix(j, COL_�ļ�ID)
                                    If bln��ҩ�÷� Then  '��ζ��ҩ��ID��¼������������ҩɾ��ʹ��
                                        strTmp = strTmp & IIF(strTmp = "", .TextMatrix(j, COL_ID), "," & .TextMatrix(j, COL_ID))
                                    End If
                                    .RemoveItem j: i = i - 1
                                Else
                                    If bln��ҩ�÷� Then
                                        .Cell(flexcpData, i, COL_���ID) = strTmp
                                    End If
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf .TextMatrix(i - 1, COL_�������) = "K" And Val(.TextMatrix(i - 1, COL_ID)) = Val(.TextMatrix(i, COL_���ID)) Then
                        bln��Ѫ;�� = True
                        '��ʾ��Ѫ;��
                        .TextMatrix(i - 1, COL_�÷�) = .TextMatrix(i, COL_�÷�) & .TextMatrix(i, COL_ҽ������)
                    Else
                        .TextMatrix(i, COL_ִ������) = ""
                    End If
                End If
                '����ҽ��������ﲡ�������й���
                If .TextMatrix(i, COL_�������) = "Z" And .TextMatrix(i, COL_��������) = "7" And .TextMatrix(i, COL_����ID) <> "" Then
                     .TextMatrix(i, COL_����ID) = ""
                End If
                '����ɼ��еĵ�һЩ��ʶ:�ſ����ɼ�����ʱδɾ������
                If Not (bln��ҩ;�� Or bln��Ѫ;��) And .TextMatrix(i, COL_�������) <> "7" Then
                    '�иߣ�Ϊ��֧��zl9PrintMode:Resize֮��,ȡRowHeight����С��RowHeightMin
                    If .RowHeight(i) < .RowHeightMin Then .RowHeight(i) = .RowHeightMin
                    
                    'ֻ��ʾ��ı����ҽ��
                    If mvarCond.����ģʽ = 3 Then
                        If Val(.TextMatrix(i, COL_������)) = 0 Then .RowHidden(i) = True: .RowHeight(i) = 0
                        '��ʾ���ֱ����ҽ��
                        If mvarCond.���� = 1 Then ' ���
                            If Not .TextMatrix(i, COL_�������) = "D" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.���� = 2 Then '����
                            If Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "C") Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        ElseIf mvarCond.���� = 3 Then ' ����
                            If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" Or .TextMatrix(i, COL_�������) = "D" Or .TextMatrix(i, COL_�������) = "C" Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    Else
                        'ֻ��ʾ��ı����ҽ��
                        If mvarCond.�Ǳ���ҽ�� And Not mvarCond.�Ǳ���ҽ�� And Val(.TextMatrix(i, COL_������)) = 0 Then
                            .RowHidden(i) = True: .RowHeight(i) = 0
                        ElseIf Not mvarCond.�Ǳ���ҽ�� And mvarCond.�Ǳ���ҽ�� And Val(.TextMatrix(i, COL_������)) <> 0 Then
                            .RowHidden(i) = True: .RowHeight(i) = 0
                        End If
                    End If
                    
                    '����ҽ���ָ�
                    If Val(.TextMatrix(i, COL_ID)) = 0 Then
                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = "�������� ����ҽ��(" & Format(mdat����, "yyyy-MM-dd HH:mm") & ") ��������"
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed
                        .Cell(flexcpAlignment, i, .FixedCols, i, .Cols - 1) = 4

                        .MergeRow(i) = True
                        .MergeCells = flexMergeFree
                    End If

                    '����С��������,��δ�뵽�취
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If
                    If Left(.TextMatrix(i, COL_����), 1) = "." Then
                        .TextMatrix(i, COL_����) = "0" & .TextMatrix(i, COL_����)
                    End If

                    'ҽ����ɫ
                    blnDo = False
                    If Val(.TextMatrix(i, COL_ҽ��״̬)) = 2 Then
                        'У������
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H80&    '���
                        blnDo = True
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 4 Then
                        '������
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '��ɫ
                        .Cell(flexcpFontStrikethru, i, .FixedCols, i, .Cols - 1) = True
                        blnDo = True
                    ElseIf InStr(",8,9,", Val(.TextMatrix(i, COL_ҽ��״̬))) > 0 Then
                        '��ֹͣ,��ȷ��ֹͣ:����������ֹʱ������ж�
                        If strCurr >= .TextMatrix(i, COL_��ֹʱ��) Or .TextMatrix(i, COL_��Ч) = "����" Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '��ɫ
                            blnDo = True
                        ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 8 And strCurr < .TextMatrix(i, COL_��ֹʱ��) Then
                            '����,ֹͣ��,ֹͣʱ��δ����һ�����
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080    'ǳ��
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 6 Then
                        '����ͣ
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 6), "yyyy-MM-dd HH:mm")
                        If strCurr >= strTime Then
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H8000&    '����
                            blnDo = True
                        Else
                            '����,��ͣ��,��ͣʱ��δ����һ�����
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HFF8080    'ǳ��
                            blnDo = True
                        End If
                    ElseIf Val(.TextMatrix(i, COL_ҽ��״̬)) = 7 Then
                        '������
                        strTime = Format(GetAdviceTime(Val(.TextMatrix(i, COL_ID)), 7), "yyyy-MM-dd HH:mm")
                        If strCurr < strTime Then
                            '����,���ú�,����ʱ��δ����һ�����
                            .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H4AAD00    'ǳ��
                            blnDo = True
                        End If
                    End If
                    If Not blnDo Then
                        If Val(.TextMatrix(i, COL_ҽ��״̬)) <> 1 And Val(.TextMatrix(i, COL_ID)) <> 0 Then
                            '��ͨ��У��(Ҳ���������Ķ��״̬)
                            If Format(.TextMatrix(i, COL_�ϴ�ִ��), "YYYY-MM-DD") >= Format(strCurr, "YYYY-MM-DD") Then  '�����ѷ��͵�(�������ܷ��͵�����)
                                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HA08000               '����
                            Else
                                .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &HC00000    '����
                            End If
                        End If
                    End If

                    'У�Ժ���ǰ����ҽ����ɫ��ʾ
                    If .TextMatrix(i, COL_�������) = "Z" And (Val(.TextMatrix(i, COL_��������)) = 4 Or Val(.TextMatrix(i, COL_��������)) = 14) _
                       And InStr(",-1,1,2,4,", Val(.TextMatrix(i, COL_ҽ��״̬))) = 0 Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed    '��ɫ
                    End If

                    '���ͺ�ת��ҽ����ɫ��ʾ
                    If .TextMatrix(i, COL_�������) = "Z" And Val(.TextMatrix(i, COL_��������)) = 3 And Val(.TextMatrix(i, COL_ҽ��״̬)) = 8 Then
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = vbRed    '��ɫ
                    End If

                    '���龫ҩƷ��ʶ:��ҩ�䷽�����ζ��ҩ������
                    If .TextMatrix(i, COL_�������) <> "" Then
                        If InStr(",����ҩ,����ҩ,����ҩ,����I��,����II��,", .TextMatrix(i, COL_�������)) > 0 Then
                            .Cell(flexcpFontBold, i, col_ҽ������) = True
                            .Cell(flexcpFontBold, i, col_����) = True
                        End If
                    End If

                    'Ƥ�Խ����ʶ
                    If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "1" And .TextMatrix(i, COL_Ƥ��) <> "" Then
                        j = GetSkinTestResult(Val(.TextMatrix(i, COL_������ĿID)), .TextMatrix(i, COL_Ƥ��))
                        .Cell(flexcpForeColor, i, COL_Ƥ��) = Decode(j, 1, vbRed, -1, vbBlue, .Cell(flexcpForeColor, i, COL_Ƥ��))
                    End If


                    'ͼ�괦��
                    '�����м���ӡ״̬��ʶ
                    Call SetAdviceReportIcon(i)

                    '����¼��
                    If Val(.TextMatrix(i, COL_������ĿID)) = 0 And Val(.TextMatrix(i, COL_ID)) <> 0 Then
                        Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("����").Picture
                    End If
                    '����ҽ��
                    If .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ" Or .TextMatrix(i, COL_Ƶ��) = "��Ҫʱ" Then
                        Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("����").Picture
                    End If
                    '������־:һ����ҩֻ��ʾ�ڵ�һ��
                    blnFirst = True
                    If InStr(",5,6,", .TextMatrix(i, COL_�������)) > 0 Then
                        If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(i - 1, COL_���ID)) Then
                            blnFirst = False
                        End If
                    End If
                    If blnFirst Then
                        If Val(.TextMatrix(i, COL_��־)) = 1 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("����").Picture
                        ElseIf Val(.TextMatrix(i, COL_��־)) = 2 Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("��¼").Picture
                        End If

                        If Val(.TextMatrix(i, COL_ҽ��״̬)) < 2 Then   '�¿����ݴ��ҽ��
                            Select Case Val(.TextMatrix(i, COL_���״̬))
                                '0-������ˣ�1-����ˣ�2-���ͨ����3-���δͨ��
                            Case 1
                                If .TextMatrix(i, COL_�������) = "K" And Val(.TextMatrix(i, COL_��鷽��)) = 1 Then
                                    '��Ѫҽ�����ͼ�굥����ʾ(��������ҽ���˶�)
                                    Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�˶�").Picture
                                Else
                                    Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                                End If
                            Case 2
                                If Not (.TextMatrix(i, COL_�������) = "K" And Val(.TextMatrix(i, COL_��鷽��)) = 1) Then
                                    Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���ͨ��").Picture
                                End If
                            Case 3
                                Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���δͨ��").Picture
                            Case 4, 5
                                If gblnѪ��ϵͳ = False Then
                                    Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                                End If
                            Case 7
                                Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("��ǩ��").Picture
                            Case Else
                            End Select
                            .Cell(flexcpPictureAlignment, i, COL_F��־) = 4
                        End If
                        '�������ϵͳ
                        If .TextMatrix(i, COL_�������״̬) = "0" Then
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("�����").Picture
                        ElseIf .TextMatrix(i, COL_�������״̬) = "2" Or .TextMatrix(i, COL_���������) = "1" Then
                            '��ʱ�������ϸ���
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���ͨ��").Picture
                        ElseIf .TextMatrix(i, COL_���������) = "2" Then
                            ' ���ϸ�
                            Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("���δͨ��").Picture
                        End If
                    End If

                    'δ��ҽ����ʶ
                    If Val(.TextMatrix(i, COL_ִ�б��)) = -1 Then
                        Set .Cell(flexcpPicture, i, COL_F��־) = frmIcons.imgFlag.ListImages("δ��").Picture
                        .Cell(flexcpForeColor, i, .FixedCols, i, .Cols - 1) = &H808080    '��ɫ
                    End If


                    'Pass:�����������ʾ��ʾ��
                    If mblnPass Then
                        If .TextMatrix(i, COL_��ʾ) <> "" Then
                            Call gobjPass.zlPassSetWarnLight(mobjPassMap, i, Val(.TextMatrix(i, COL_��ʾ)))
                            .TextMatrix(i, COL_��ʾ) = ""
                        End If
                    End If
                End If

                If bln��ҩ;�� Or bln��Ѫ;�� Then
                    .RemoveItem i
                Else
                    '���ģʽ�����ҽ������
                    If mvarCond.��ʾģʽ = 0 And mvarCond.����ģʽ <> 3 Then
                        strFormat = .TextMatrix(i, col_ҽ������)
                        If .TextMatrix(i, COL_�������) <> "Z" And Val(.TextMatrix(i, COL_������ĿID)) <> 0 And InStr(strFormat, "����ҽ��") = 0 Then
                            'ҽ�����ݶ����а����������ʱ�������ظ����
                            mrsDefine.Filter = "�������='" & .TextMatrix(i, COL_�������) & "'"
                            If Not (.TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "1") Then
                                strFormat = strFormat & .TextMatrix(i, COL_Ƥ��)
                            End If

                            If Not (InStr("5,6,7", .TextMatrix(i, COL_�������)) = 0 And .TextMatrix(i, COL_Ƶ��) = "һ����") Then
                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_����)
                                    If strTmp <> "" Then strFormat = strFormat & ",��" & strTmp
                                End If

                                blnDo = True
                                If mrsDefine.RecordCount > 0 Then blnDo = InStr(mrsDefine!ҽ������, "[����]") = 0
                                If blnDo Then
                                    strTmp = .TextMatrix(i, COL_����)
                                    If strTmp <> "" Then strFormat = strFormat & ",ÿ��" & strTmp
                                End If
                            End If
                        End If
                        .TextMatrix(i, col_����) = strFormat


                        '�ϲ��÷���:�÷� Ƶ�� ����(һ����ҩ����ǰ���Ѵ���)
                        If .TextMatrix(i, COL_�������) <> "Z" And Val(.TextMatrix(i, COL_������ĿID)) <> 0 And InStr(strFormat, "����ҽ��") = 0 Then
                            
                            '���ģʽ�³�ҩƷ��������Ŀ��������ҽ������ʾ�÷�
                            If .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) = 0 Or _
                                InStr(",5,6,7,", "," & .TextMatrix(i, COL_�������) & ",") > 0 And Val(.TextMatrix(i, COL_���ID)) <> 0 Then
                                strFormat = .TextMatrix(i, COL_�÷�)
                            Else
                                strFormat = ""
                            End If
                            
                            '���� '��� '��Ѫ '���� '����ȼ� ���ģʽ�²���ʾƵ��
                            If .TextMatrix(i, COL_�������) = "E" And Val(.TextMatrix(i, COL_��������)) = 6 Or _
                                .TextMatrix(i, COL_�������) = "D" And Val(.TextMatrix(i, COL_���ID)) = 0 Or _
                                .TextMatrix(i, COL_�������) = "K" And Val(.TextMatrix(i, COL_���ID)) = 0 Or _
                                .TextMatrix(i, COL_�������) = "F" And Val(.TextMatrix(i, COL_���ID)) = 0 Or _
                                .TextMatrix(i, COL_�������) = "H" And Val(.TextMatrix(i, COL_��������)) = 1 Then
                                strTmp = ""
                            Else
                                strTmp = .TextMatrix(i, COL_Ƶ��)
                            End If
                            If strTmp <> "" Then strFormat = strFormat & IIF(strFormat <> "", ",", "") & strTmp

                            strTmp = .TextMatrix(i, COL_����)
                            If strTmp <> "" Then
                                strFormat = strFormat & IIF(strFormat <> "", ",", "") & "��" & strTmp & "��"
                            End If

                            .TextMatrix(i, COL_�÷�) = strFormat
                        End If
                        
                        '���ģʽ�£�������ֹʱ����ʾΪ�ա�
                        If .TextMatrix(i, COL_��Ч) = "����" Then .TextMatrix(i, COL_��ֹʱ��) = ""
                    End If
                    
                    If mvarCond.����ģʽ = 3 Then
                        '����Ǳ���ҳǩ�£����� �� ����Ϊ�գ����¸�ֵ
                        .TextMatrix(i, col_����) = .TextMatrix(i, col_ҽ������)
                        If Val(.TextMatrix(i, COL_����ID)) = 0 And .TextMatrix(i, COL_��鱨��ID) = "" And Val(.TextMatrix(i, COL_RIS����ID)) = 0 And Val(.TextMatrix(i, COL_LIS����ID)) = 0 Then
                            .TextMatrix(i, COL_����״̬) = "δ��"
                        Else
                            .TextMatrix(i, COL_����״̬) = "����"
                            If Val(.Cell(flexcpData, i, COL_����״̬)) = 0 Then  'δ��
                                .Cell(flexcpForeColor, i, COL_����״̬, i, COL_����״̬) = &HFF0000     '��ɫ
                            ElseIf Val(.Cell(flexcpData, i, COL_����״̬)) = 2 Then  '�����Ѷ�
                                .Cell(flexcpForeColor, i, COL_����״̬, i, COL_����״̬) = &HFF00FF     '��ɫ
                            Else
                                .Cell(flexcpForeColor, i, COL_����״̬, i, COL_����״̬) = &H80&      '����
                            End If
                            .Cell(flexcpFontUnderline, i, COL_����״̬, i, COL_����״̬) = True
                        End If
                        '���ӹ���δ���ı�����ѳ��ı���
                        If .RowHidden(i) = False Then
                            If Not IIF(.TextMatrix(i, COL_����״̬) = "δ��", mvarCond.δ������, mvarCond.�ѳ�����) Then
                                .RowHidden(i) = True: .RowHeight(i) = 0
                            End If
                        End If
                    End If
                    i = i + 1
                End If
            Loop
            
            '����ҽ�����ݵ�Ԫ���ͼ��,����ǩ����ʶ�����δ�ӡ,Σ��ֵ
            For i = 1 To .Rows - 1
                Call SetAdviceIcon(i)
            Next
            
            '�Զ������и�
            If mvarCond.��ʾģʽ = 0 And mvarCond.����ģʽ <> 3 Then
                If InStr("2505,3345,1005,1335", .ColWidth(COL_�÷�)) > 0 Then .ColWidth(COL_�÷�) = IIF(mlngFontSize = 9, 2505, 3345)   '�û�δ�ĸ��п�ʱ������
                .AutoSize col_����, COL_�÷�
                .ColWidth(COL_��ʼʱ��) = IIF(mlngFontSize = 9, 1130, 1510)
            Else
                If InStr("2505,3345,1005,1335", .ColWidth(COL_�÷�)) > 0 Then .ColWidth(COL_�÷�) = IIF(mlngFontSize = 9, 1005, 1335)
                .AutoSize col_ҽ������, COL_�÷�
                .ColWidth(COL_��ʼʱ��) = IIF(mlngFontSize = 9, 1530, 2040)
            End If

            '�̶���ͼ�����:����Ϊ�ж���,��Ȼ���߿�ʱ����������
            .Cell(flexcpPictureAlignment, .FixedRows, 0, .Rows - 1, .FixedCols - 1) = 4
            '����ǩ��ͼ�����
            .Cell(flexcpPictureAlignment, .FixedRows, col_ҽ������, .Rows - 1, col_ҽ������) = 0
            Call SetTagһ����ҩ
            Call Set�걾״̬
            .Redraw = True
        End With
    Else
        Call ClearAdviceData
        Call ClearAppendData
    End If
    Call SetAdviceColVisible
    'ֻ������ʱ���ú�ɫ����
    vsAdvice.GridColor = IIF(mvarCond.����ģʽ = "2", &H8080FF, vsAdvice.GridColorFixed)
    
    imgColSel.Visible = (mvarCond.��ʾģʽ = 1 And mvarCond.����ģʽ <> 3)
    
    Call LocatedDefaultAdviceRow(lngҽ��ID)
    
    Screen.MousePointer = 0
    LoadAdvice = True
    If Not mfrmParent Is Nothing Then
        '�°滤ʿվ����ʱ��Ĭ������һ����ɫ��afterrowcolchange���޷����á�
        If mfrmParent.Name = "frmInNurseRoutine" Then
            If vsAdvice.Col >= vsAdvice.FixedCols Then
                vsAdvice.ForeColorSel = vsAdvice.Cell(flexcpForeColor, vsAdvice.Row, COL_��ʼʱ��)
            End If
        End If
    End If
    '�Զ�ˢ��ҽ����������
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
'���ܣ����ݵ�ǰ�е���������ҽ�����ݵ�ͼ���ʶ
'˵����ע���ǵ������ã�����һ������
    Dim intͼ���� As Integer 'ҽ�����������ͼ�����
    
    intͼ���� = 1
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_��˱��)) = 2 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = img16.ListImages("ͣ������").Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = img16.ListImages("ͣ������").Picture
    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_ǩ����)) = 1 And Val(vsAdvice.TextMatrix(lngRow, COL_���δ�ӡ)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = img16dbl.ListImages(1).Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = img16dbl.ListImages(1).Picture
    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_ǩ����)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgSign.ListImages("ǩ��").Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgSign.ListImages("ǩ��").Picture
    ElseIf Val(vsAdvice.TextMatrix(lngRow, COL_���δ�ӡ)) = 1 Then
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = img16.ListImages("���δ�ӡ").Picture
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = img16.ListImages("���δ�ӡ").Picture
    Else
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = Nothing
        Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = Nothing
        intͼ���� = 0
    End If
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_��ΣҩƷ)) > 0 Then
        If vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) Is Nothing Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture
            intͼ���� = 1
        Else
            If vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) <> frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture Then
                pictmp.Cls
                pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
                pictmp.PaintPicture frmIcons.imgQuestion.ListImages("��ΣҩƷ").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
                Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
                intͼ���� = 2
            End If
        End If
    End If
    
    'Σ��ֵͼ��
    If Val(vsAdvice.TextMatrix(lngRow, COL_Σ��ֵID)) > 0 Then
        If intͼ���� = 0 Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture
        ElseIf intͼ���� = 1 Then
            pictmp.Cls
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
            intͼ���� = 2
        ElseIf intͼ���� = 2 Then
            pictmp.Cls
            pictmp.Width = 720
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, 480, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("Σ��ֵ").Picture, 480, 0, 240, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
            pictmp.Width = 480
            intͼ���� = 3
        End If
    End If
    
    '�׵���ͼ��
    If Val(vsAdvice.TextMatrix(lngRow, COL_�׵���)) > 0 Then
        If intͼ���� = 0 Then
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = frmIcons.imgQuestion.ListImages("�׵���").Picture
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = frmIcons.imgQuestion.ListImages("�׵���").Picture
        ElseIf intͼ���� = 1 Then
            pictmp.Cls
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, pictmp.Width / 2, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, pictmp.Width / 2, 0, pictmp.Width / 2, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
            intͼ���� = 2
        ElseIf intͼ���� = 2 Then
            pictmp.Cls
            pictmp.Width = 720
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, 480, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, 480, 0, 240, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
            pictmp.Width = 480
            intͼ���� = 3
        ElseIf intͼ���� = 3 Then
            pictmp.Cls
            pictmp.Width = 960
            pictmp.PaintPicture vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������), 0, 0, 720, pictmp.Height
            pictmp.PaintPicture frmIcons.imgQuestion.ListImages("�׵���").Picture, 720, 0, 240, pictmp.Height
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_ҽ������) = pictmp.Image
            Set vsAdvice.Cell(flexcpPicture, lngRow, col_����) = pictmp.Image
            pictmp.Width = 480
            intͼ���� = 4
        End If
    End If
End Sub

Private Function RowIs�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���ҩ�䷽��
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='7' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs�䷽�� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIs������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���������
'˵����ָ����Ϊ��ʾ��,�����="E"
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
        
    On Error GoTo errH
    
    strSQL = "Select ID From ����ҽ����¼ Where Rownum=1 And �������='C' And ���ID=[1]"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    If Not rsTmp.EOF Then RowIs������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ShowPrice(ByVal lngRow As Long) As Boolean
'���ܣ���ȡָ��ҽ���ļƼ�,�����ݵ�ǰ�������շ� ��ϵ���и���
    Dim rs������Ŀ As New ADODB.Recordset
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim strҽ��IDs As String, str�շ�ϸĿIDs As String, str�����շ� As String
    Dim strSQL As String, i As Long, j As Long
    Dim bln�䷽�� As Boolean, bln������ As Boolean, blnLoad As Boolean
    Dim lng���˿���ID As Long, lngִ�п���ID As Long
    Dim dblPrice As Double, lng����ID As Long
    Dim lngҽ��ID As Long, lng���ID As Long
    Dim strPriceType As String

    On Error GoTo errH
    
    With vsAppend
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIs�䷽��(lngRow)
            bln������ = RowIs������(lngRow)
        End If
        
        lngҽ��ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
        lng���ID = Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
                                    
        blnLoad = True
        
        'ҩƷ�����ĵļƼ�
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "4" Then
            '���ļƼ�
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������,0 as �շѷ�ʽ," & _
                " A.�շ�ϸĿID,1 as סԺ��װ,C.���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,Nvl(B.����,D.ȱʡ�۸�),D.�ּ�) as ����,A.ִ�п���ID,0 as ����,C.��� as �շ����" & _
                " From ����ҽ����¼ A,����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1] And A.ID=B.ҽ��ID(+) And A.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0) Not IN(0,5)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "5", "6", "7") & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
                
                blnLoad = False
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            '��,����ҩ:���ܰ������ҽ��,����1��סԺ��װ�ĵ���
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������,0 as �շѷ�ʽ," & _
                " C.ID as �շ�ϸĿID,B.סԺ��װ,B.סԺ��λ as ���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.סԺ��װ as ����," & _
                " A.ִ�п���ID,0 as ����,C.��� as �շ����" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID And Nvl(A.ִ������,0) Not IN(0,5)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "5", "6", "7") & _
                " And (A.�շ�ϸĿID is NULL Or A.�շ�ϸĿID=B.ҩƷID)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
                
                '��һ����ҩ(�����)�ĵ�һ��ҩ�в���ʾ��ҩ;���ļƼ�
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_���ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
        ElseIf bln�䷽�� Then
            '�в�ҩ:һ����Ӧ�й���¼����д���շ�ϸĿID
            strSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������,0 as �շѷ�ʽ," & _
                " C.ID as �շ�ϸĿID,B.סԺ��װ,B.סԺ��λ as ���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*B.סԺ��װ as ����," & _
                " A.ִ�п���ID,0 as ����,C.��� as �շ����" & _
                " From ����ҽ����¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.�������='7' And A.���ID=[1]" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "5", "6", "7") & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID And C.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3)" & _
                " And D.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0) Not IN(0,5)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        End If
        
        '��ȡ���мƼ�(ȡ���¼۸�)����ҩƷ��������ļƼ�,�������ҽ���Ƽ�
        '���Ƽ�,�ֹ��Ƽ۵�ҽ������ȡ
        '��Union��ʽ������������
        If blnLoad Then
            '�����¿���ҽ�������ݲ���ҽ���Ƽ���ȡ
            If InStr(",1,2,-1,", vsAdvice.TextMatrix(lngRow, COL_ҽ��״̬)) = 0 Then
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0) as ��������,Nvl(B.�շѷ�ʽ,0) as �շѷ�ʽ," & _
                    " B.�շ�ϸĿID,1 as סԺ��װ,C.���㵥λ,B.����,Decode(C.�Ƿ���,1,B.����,Sum(D.�ּ�)) as ����," & _
                    " Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID,Nvl(B.����,0) as ����,C.��� as �շ����" & _
                    " From ����ҽ����¼ A,����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                    " Where A.������� Not IN('4','5','6','7') And A.ID=B.ҽ��ID" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "5", "6", "7") & _
                    " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ�ϸĿID=C.ID And B.�շ�ϸĿID=D.�շ�ϸĿID" & _
                    " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                    " And (A.ID=[1]" & IIF(lng���ID <> 0, " Or A.ID=[2]", "") & " Or A.���ID=[1])" & _
                    " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0),Nvl(B.�շѷ�ʽ,0)," & _
                    " B.�շ�ϸĿID,C.���,C.���㵥λ,B.����,C.�Ƿ���,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID),Nvl(B.����,0)"
            Else
            '�¿���ҽ�������������շ� ��ϵ��ȡ(��ҩ�����ʾΪ0)
            '�������۲��ˣ����ÿ��Ҳ��������
            '���ֶ�Ӧ�ļƼۣ�
            '   1.���յķ��ã�ֻ������Ŀ������գ�Ŀǰֻ�д��Ի������������
            '   2.�����ķ��ã����Ǿ���ļ�鲿λ�ͼ�鷽����
            '   3.�����ķ��ã��Ǽ�鲿λ�ͷ�����(ע�����걾��д�ڱ걾��λ��)
                lng����ID = 0 '�����Թܷ���,ֻ��ȡ�Թܶ�Ӧ�����ķ���
                If vsAdvice.TextMatrix(lngRow, COL_�Թܱ���) <> "" Then
                    lng����ID = GetTubeMaterial(vsAdvice.TextMatrix(lngRow, COL_�Թܱ���))
                End If
                
                str�����շ� = "Select * From (" & _
                    "Select C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,c.���ÿ���id" & _
                    " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                    " From �����շѹ�ϵ C,����ҽ����¼ A Where (A.ID=[1]" & IIF(lng���ID <> 0, " Or A.ID=[2]", "") & " Or A.���ID=[1]) And A.������ĿID+0=C.������ĿID" & _
                    "   And (a.���id Is Null And a.ִ�б�� In (1, 2) And c.�������� = 1 Or" & vbNewLine & _
                    "   a.�걾��λ = c.��鲿λ And a.��鷽�� = c.��鷽�� And Nvl(c.��������, 0) = 0 Or" & vbNewLine & _
                    "   (a.��鷽�� Is Null or a.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=a.������ĿID And Z.��������='4')) And Nvl(c.��������, 0) = 0 And c.��鲿λ Is Null And c.��鷽�� Is Null)" & _
                    "      And (C.���ÿ���ID is Null or C.���ÿ���ID = Nvl(A.ִ�п���ID,[4]) And C.������Դ = " & IIF(mlng�������� = 1, 1, 2) & ")" & _
                    " ) Where Nvl(���ÿ���id, 0) = Top"
                
                strSQL = strSQL & IIF(strSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0) as ��������,Nvl(B.�շѷ�ʽ,0) as �շѷ�ʽ," & _
                    " B.�շ���ĿID as �շ�ϸĿID,1 as סԺ��װ,C.���㵥λ,B.�շ����� as ����,Decode(C.�Ƿ���,1,Sum(D.ȱʡ�۸�),Sum(D.�ּ�)) as ����," & _
                    " A.ִ�п���ID,Nvl(B.������Ŀ,0) as ����,C.��� as �շ����" & _
                    " From ����ҽ����¼ A,(" & str�����շ� & ") B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                    " Where A.������� Not IN('4','5','6','7') And A.ҽ��״̬ IN(-1,1,2) And A.������ĿID+0=B.������ĿID" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "5", "6", "7") & _
                    " And (A.���ID is Null And A.ִ�б�� IN(1,2) And B.��������=1" & _
                    "       Or A.�걾��λ=B.��鲿λ And A.��鷽��=B.��鷽�� And Nvl(B.��������,0)=0" & _
                    "       Or (A.��鷽�� is Null or a.������� = 'E' And Exists(Select 1 From ������ĿĿ¼ Z Where Z.id=a.������ĿID And Z.��������='4')) And Nvl(B.��������,0)=0 And B.��鲿λ is Null And B.��鷽�� is Null)" & _
                    " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ���ĿID=C.ID And B.�շ���ĿID=D.�շ�ϸĿID" & _
                    " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                    " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) And C.������� IN(" & IIF(mlng�������� = 1, 1, 2) & ",3)" & _
                    " And (Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And B.�շ���ĿID=[3] Or Not(Nvl(B.�շѷ�ʽ,0)=1 And C.���='4' And [3]<>0))" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null) And (A.ID=[1]" & IIF(lng���ID <> 0, " Or A.ID=[2]", "") & " Or A.���ID=[1])" & _
                    " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0),Nvl(B.�շѷ�ʽ,0)," & _
                    " B.�շ���ĿID,C.���,C.���㵥λ,B.�շ�����,C.�Ƿ���,A.ִ�п���ID,Nvl(B.������Ŀ,0)"
            End If
        End If
        strSQL = strSQL & " Order by ���,��������,����,�շ����"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ���Ƽ�", "H����ҽ���Ƽ�")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lngҽ��ID, lng���ID, lng����ID, mlng����ID, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        '��ʾ�Ƽ�����
        If Not rsTmp.EOF Then
            'ȷ����ʾ����
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '��ȡ������Ŀ,�շ�ϸĿ��Ϣ
            For i = 1 To rsTmp.RecordCount
                If InStr("," & strҽ��IDs & ",", "," & rsTmp!ID & ",") = 0 Then strҽ��IDs = strҽ��IDs & "," & rsTmp!ID
                If InStr("," & str�շ�ϸĿIDs & ",", "," & rsTmp!�շ�ϸĿID & ",") = 0 Then str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & rsTmp!�շ�ϸĿID
                rsTmp.MoveNext
            Next
            strҽ��IDs = Mid(strҽ��IDs, 2)
            str�շ�ϸĿIDs = Mid(str�շ�ϸĿIDs, 2)
                        
            If mblnMoved Then
            'ͨ��DBLink���ӵ�Զ����ʷ�����⴦��ʹ��f_Num2list���޷���������,��driving_site�����Զ�̴���Ƶ���ǰ������
                
                strSQL = "Select /*+driving_site(a)*/ B.ID,B.���,C.���� as �������,B.����,B.�걾��λ" & _
                      " From H����ҽ����¼ A,������ĿĿ¼ B,������Ŀ��� C" & _
                      " Where A.������ĿID=B.ID And B.���=C.���� And A.ID "
                If InStr(strҽ��IDs, ",") > 0 Then              '����SQL�����֣�ֻ�ò��ð󶨱���
                    strSQL = strSQL & " In(" & strҽ��IDs & ")"
                Else
                    strSQL = strSQL & " = [1]"
                End If
            Else
              strSQL = "Select /*+cardinality(d,10)*/ B.ID,B.���,C.���� as �������,B.����,B.�걾��λ" & _
                  " From ����ҽ����¼ A,������ĿĿ¼ B,������Ŀ��� C,Table(f_Num2list([1])) D" & _
                  " Where A.ID = D.Column_Value And A.������ĿID=B.ID And B.���=C.����"
            End If
            Set rs������Ŀ = zlDatabase.OpenSQLRecord(strSQL, Me.Name, strҽ��IDs)
            
            strSQL = "Select A.ID,A.���,B.���� as �������,A.����," & _
                " A.����,A.���,A.����,A.��������,A.�Ƿ���" & _
                " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,Table(f_Num2list([1])) D" & _
                " Where A.���=B.���� And A.ID = D.Column_Value"
            strSQL = "Select /*+ Rule*/ A.ID,A.���,A.�������,A.����,Nvl(B.����,A.����) as ����," & _
                " A.���,A.����,A.��������,A.�Ƿ���,C.��������" & _
                " From (" & strSQL & ") A,�շ���Ŀ���� B,�������� C" & _
                " Where A.ID=C.����ID(+) And A.ID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=[2]"
            Set rs�շ�ϸĿ = zlDatabase.OpenSQLRecord(strSQL, Me.Name, str�շ�ϸĿIDs, IIF(gbytҩƷ������ʾ = 0, 1, 3))
            
            '��ʾÿ������
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs������Ŀ.Filter = "ID=" & rsTmp!������ĿID
                rs�շ�ϸĿ.Filter = "ID=" & rsTmp!�շ�ϸĿID
                
                '�Ƽ�ҽ��
                If rsTmp!������� = "4" Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��������-" & rs������Ŀ!����
                ElseIf InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "ҩƷҽ��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��ҩ;��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And vsAdvice.TextMatrix(lngRow, COL_�������) = "K" Then
                    .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��Ѫ;��-" & rs������Ŀ!����
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "�ɼ�����-" & rs������Ŀ!����
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��ҩ�巨-" & rs������Ŀ!����
                    Else
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��ҩ�÷�-" & rs������Ŀ!����
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "������Ŀ-" & rs������Ŀ!����
                    ElseIf rsTmp!������� = "D" Then
                        '��λ������
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��鲿λ-" & NVL(rsTmp!�걾��λ) & "(" & NVL(rsTmp!��鷽��) & ")"
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "��������-" & rs������Ŀ!����
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = "������Ŀ-" & rs������Ŀ!����
                    End If
                Else
                    If NVL(rsTmp!��������, 0) = 1 Then
                        '���Ի����м��շ���
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!���� & "(" & Decode(NVL(rsTmp!ִ�б��, 0), 1, "����", 2, "����", "") & "����)"
                    Else
                        .TextMatrix(i, COLPrice("�Ƽ�ҽ��")) = rs������Ŀ!������� & "ҽ��-" & rs������Ŀ!����
                    End If
                End If
                
                '���
                .TextMatrix(i, COLPrice("���")) = rs�շ�ϸĿ!�������
                '�շ���Ŀ:���/����
                .TextMatrix(i, COLPrice("�շ���Ŀ")) = rs�շ�ϸĿ!����
                If Not IsNull(rs�շ�ϸĿ!����) Then
                    .TextMatrix(i, COLPrice("�շ���Ŀ")) = .TextMatrix(i, COLPrice("�շ���Ŀ")) & "(" & rs�շ�ϸĿ!���� & ")"
                End If
                If Not IsNull(rs�շ�ϸĿ!���) Then
                    .TextMatrix(i, COLPrice("�շ���Ŀ")) = .TextMatrix(i, COLPrice("�շ���Ŀ")) & " " & rs�շ�ϸĿ!���
                End If
                
                '���㵥λ:ҩ��ҩƷΪסԺ��λ,��ҩ��ҩƷΪ�ۼ۵�λ
                .TextMatrix(i, COLPrice("��λ")) = NVL(rsTmp!���㵥λ)
                '�Ƽ�����:ҩ��ҩƷΪ1,��ҩ��ҩƷΪ��Ӧ�ۼ���
                .TextMatrix(i, COLPrice("�Ƽ�����")) = FormatEx(rsTmp!����, 5)
                
                'ִ�п���
                lngִ�п���ID = NVL(rsTmp!ִ�п���ID, 0)
                If rs�շ�ϸĿ!��� = "4" And NVL(rs�շ�ϸĿ!��������, 0) = 1 Or _
                    InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 And InStr(",5,6,7,", rs������Ŀ!���) = 0 Then
                    lng���˿���ID = mlng����ID
                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, rs�շ�ϸĿ!���, rs�շ�ϸĿ!ID, 4, lng���˿���ID, 0, 2, lngִ�п���ID, , , 2)
                End If
                
                '���۴���
                If InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 Then
                    If NVL(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                        '��ҩƷʱ��
                        If InStr(",5,6,7,", rs������Ŀ!���) > 0 Then
                            'ҩ��ҩƷ����һ��סԺ��װ��סԺʱ��
                            .TextMatrix(i, COLPrice("����")) = CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, NVL(rsTmp!סԺ��װ, 1), , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            .TextMatrix(i, COLPrice("����")) = Format(Val(.TextMatrix(i, COLPrice("����"))) * NVL(rsTmp!סԺ��װ, 0), gstrDecPrice)
                        Else
                            '��ҩ��ҩƷ��������ۼ��������ۼ�ʵ��
                            .TextMatrix(i, COLPrice("����")) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, NVL(rsTmp!����, 0), , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                        End If
                    Else
                        'ҩ��ҩƷΪסԺ����,��ҩҩƷΪ�ۼ�
                        .TextMatrix(i, COLPrice("����")) = Format(NVL(rsTmp!����), gstrDecPrice)
                    End If
                ElseIf rs�շ�ϸĿ!��� = "4" And NVL(rs�շ�ϸĿ!��������, 0) = 1 And NVL(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                    'ʱ�����ĵĵ��ۺ�ҩƷһ������
                    .TextMatrix(i, COLPrice("����")) = Format(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, NVL(rsTmp!����, 0), , , , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gstrDecPrice)
                Else
                    .TextMatrix(i, COLPrice("����")) = Format(NVL(rsTmp!����), gstrDecPrice)
                End If
                
                'ִ�п���
                If lngִ�п���ID <> 0 Then
                    .TextMatrix(i, COLPrice("ִ�п���")) = Sys.RowValue("���ű�", lngִ�п���ID, "����")
                End If
                
                '��ʾҽ����������
                If Val(rsTmp!�շ�ϸĿID & "") <> 0 Then
                    strPriceType = GetPriceType(mlng����ID, Val(rsTmp!�շ�ϸĿID & ""), mint����, mlng�������� = 1)
                End If
                '��������
                If strPriceType = "" Then
                    .TextMatrix(i, COLPrice("��������")) = NVL(rs�շ�ϸĿ!��������)
                Else
                    .TextMatrix(i, COLPrice("��������")) = strPriceType
                End If

                
                '������Ŀ
                .TextMatrix(i, COLPrice("����")) = IIF(NVL(rsTmp!����, 0) = 0, "", "��")
                
                '�շѷ�ʽ
                .TextMatrix(i, COLPrice("�շѷ�ʽ")) = getChargeMode(Val(NVL(rsTmp!�շѷ�ʽ, 0)))
                
                dblPrice = dblPrice + Format(Val(.TextMatrix(i, COLPrice("�Ƽ�����"))) * Val(.TextMatrix(i, COLPrice("����"))), "0.00000")
                
                rsTmp.MoveNext
            Next
        End If
        
        '�ϼ���
        If .Rows > 2 Then
            .MergeCol(COLPrice("�Ƽ�ҽ��")) = True
            .MergeCol(COLPrice("���")) = True
            
            .Rows = .Rows + 1
            .Cell(flexcpText, .Rows - 1, COLPrice("�Ƽ�ҽ��"), .Rows - 1, COLPrice("��λ")) = "�ϼ�"
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("�Ƽ�ҽ��"), .Rows - 1, COLPrice("��λ")) = 4
            .Cell(flexcpText, .Rows - 1, COLPrice("�Ƽ�����"), .Rows - 1, COLPrice("����")) = Format(dblPrice, gstrDecPrice)
            .Cell(flexcpAlignment, .Rows - 1, COLPrice("�Ƽ�����"), .Rows - 1, COLPrice("����")) = 7
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
'���ܣ���ʾָ����ҽ���ķ��ͼ�¼
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strSub As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strExe1 As String, strExe2 As String, strState As String
    Dim bln�䷽�� As Boolean, bln������ As Boolean
    Dim strKey As String, lngKey As Long
    Dim rsִ�� As ADODB.Recordset
    Dim str���ͺ� As String, strTab As String
    Dim bln״̬˵�� As Boolean
    Dim lng��Ѫ As Long
    Dim j As Long
    
    On Error GoTo errH
        lng��Ѫ = -1
    With vsAppend
        '��¼ԭ��λ��
        lngKey = -1
        If .Row >= .FixedRows Then
            strKey = .TextMatrix(.Row, COLSend("���ͺ�")) & "," & .TextMatrix(.Row, COLSend("ҽ��ID")) & "," & .TextMatrix(.Row, COLSend("�շ���Ŀ"))
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_��Ч) = "����" And 1 <> mlng�������� Then
            strTab = "סԺ���ü�¼"
        Else
            If GetAdviceFeeKind(Val(vsAdvice.TextMatrix(lngRow, COL_ID))) = 2 Then  'סԺҽ��վ�������ɷ��͵�����
                strTab = "סԺ���ü�¼"
            Else
                strTab = "������ü�¼"
            End If
        End If
    
        .Redraw = False
        If .FixedRows = 2 Then .RemoveItem 0
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        
        If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
            .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If mbln��������ִ�� And Val(vsAdvice.TextMatrix(lngRow, COL_ҽ��״̬)) = 4 And Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID)) = 0 Then
            Call SetExecShow(False, mblnShowExec)
           .Redraw = True: ShowSendList = True: Exit Function
        End If
        
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIs�䷽��(lngRow)
            bln������ = RowIs������(lngRow)
        End If
                
        strExe1 = "Decode(Nvl(A.ִ��״̬,0),0,'δִ��',1,'��ȫִ��',2,'����ִ��')"
        strExe2 = "Decode(Nvl(B.ִ��״̬,0),0,'δִ��',1,'ִ�����',2,'�ܾ�ִ��',3,'����ִ��')"
        strState = "Decode(A.ִ��״̬,9,'�շ��쳣',Decode(A.��¼����,1,Decode(A.��¼״̬,0,'�շѻ���',1,'���շ�',3,'���˷�'),2,Decode(A.��¼״̬,0,'���ʻ���',1,'�Ѽ���',3,'������'),'δ�Ʒ�'))"
        
        'ҩ����Ӧ��ҩƷ�Ƽ۰�סԺ��װ��ʾ,��ҩ����Ӧ��ҩƷ�Ƽ۰����۵�λ��ʾ
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            If Not RowInһ����ҩ(lngRow, lngBegin, lngEnd) Then lngBegin = lngRow
            '��ҩ����:��д�˷��ͼ�¼,�������޶�Ӧ����(���Ա�ҩ,��ҽ���й��)
            strSub = "Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                " From " & strTab & " A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL And A.�շ���� IN('5','6','7')" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.ҽ�����=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, strTab, "H" & strTab)
            ElseIf zlDatabase.DateMoved(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, strTab, "H" & strTab)
            End If
                
            strSQL = _
                " Select B.ҽ��ID,C.���ID,C.�걾��λ,C.��鷽��,B.����ʱ��,B.NO,B.��¼����,B.�������,A.�շ�ϸĿID," & _
                " Nvl(A.סԺ��λ,D.סԺ��λ) as ��λ," & _
                " Nvl(A.����/Nvl(A.סԺ��װ,1),B.��������/Nvl(D.����ϵ��,1)/Nvl(D.סԺ��װ,1)) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID,Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬," & _
                " B.�״�ʱ��,B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�'," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.״̬˵��,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������," & _
                " B.ִ��״̬ as ִ��״̬ID,B.�Ʒ�״̬ as �Ʒ�״̬ID,A.��¼״̬,NVL(B.���ʱ��,A.ִ��ʱ��) as ���ʱ��,NVL(B.�����,A.ִ����) as �����,B.ִ��˵��,B.����ʱ��,B.������,B.����ʱ��" & _
                " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,ҩƷ��� D" & _
                " Where B.ҽ��ID=C.ID And C.�շ�ϸĿID=D.ҩƷID(+)" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And A.ҽ�����(+)=B.ҽ��ID" & _
                " And C.ID=[1]"

            '��һ����ҩ�����в���ʾ��ҩ;���ķ���
            If lngRow = lngBegin Then
                '��ҩ;������:��д�˷��ͼ�¼(������),����һ���з���
                strSub = "Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                    " From " & strTab & " A,ҩƷ��� B" & _
                    " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                    " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[2]"
                If mblnMoved Then
                    strSub = Replace(strSub, strTab, "H" & strTab)
                ElseIf zlDatabase.DateMoved(mvInDate) Then
                    strSub = strSub & " Union ALL " & Replace(strSub, strTab, "H" & strTab)
                End If
                    
                strSQL = strSQL & " Union ALL " & _
                    " Select B.ҽ��ID,C.���ID,C.�걾��λ,C.��鷽��,B.����ʱ��,B.NO,B.��¼����,B.�������,A.�շ�ϸĿID," & _
                    " Decode(Nvl(Instr('567',A.�շ����),0),0,Decode(A.�շ����,'4',F.���㵥λ,D.���㵥λ),Nvl(A.סԺ��λ,E.סԺ��λ)) as ��λ," & _
                    " Nvl(A.����/Nvl(A.סԺ��װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.סԺ��װ,1)) as ��������," & _
                    " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID,Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬," & _
                    " B.�״�ʱ��,B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�'," & strState & ") as �Ʒ�״̬," & _
                    " B.������,B.״̬˵��,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������," & _
                    " B.ִ��״̬ as ִ��״̬ID,B.�Ʒ�״̬ as �Ʒ�״̬ID,A.��¼״̬ ,NVL(B.���ʱ��,A.ִ��ʱ��) as ���ʱ��,NVL(B.�����,A.ִ����) as �����,B.ִ��˵��,B.����ʱ��,B.������,B.����ʱ��" & _
                    " From (" & strSub & ") A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D,ҩƷ��� E,�շ���ĿĿ¼ F" & _
                    " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID And C.�շ�ϸĿID=E.ҩƷID(+) And C.�շ�ϸĿID=F.ID(+)" & _
                    " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID" & _
                    " And C.ID=[2]"
            End If
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        Else
            '����ҽ��(�������ġ��䷽����飬����һ��ҽ��):��д�˷��ͼ�¼(������),����һ���з���
            '��ҩ�Ա�ҩҲ���޶�Ӧ����(��ҽ���й��)
            strSub = _
                " Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                " From " & strTab & " A,ҩƷ��� B" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=[1]"
            strSub = strSub & " Union ALL " & _
                " Select A.*,B.סԺ��װ,B.סԺ��λ" & _
                " From " & strTab & " A,ҩƷ��� B,����ҽ����¼ C" & _
                " Where A.��¼״̬ IN(0,1,3) And A.�۸񸸺� is NULL" & _
                " And A.�շ�ϸĿID=B.ҩƷID(+) And A.ҽ�����=C.ID" & _
                " And C.���ID=[1]"
            If mblnMoved Then
                strSub = Replace(strSub, strTab, "H" & strTab)
            ElseIf zlDatabase.DateMoved(mvInDate) Then
                strSub = strSub & " Union ALL " & Replace(strSub, strTab, "H" & strTab)
            End If
            
            strSQL = _
                " Select * From ����ҽ����¼ Where ID=[1]" & _
                " Union ALL " & _
                " Select * From ����ҽ����¼ Where ���ID=[1]"
            strSQL = _
                " Select B.ҽ��ID,C.ҽ������,C.���ID,C.�걾��λ,C.��鷽��,B.����ʱ��,B.NO,B.��¼����,B.�������,A.�շ�ϸĿID," & _
                " Decode(Nvl(Instr('567',A.�շ����),0),0,Decode(A.�շ����,'4',F.���㵥λ,D.���㵥λ),Nvl(A.סԺ��λ,E.סԺ��λ)) as ��λ," & _
                " Nvl(Nvl(A.����,1)*A.����/Nvl(A.סԺ��װ,1),B.��������/Nvl(E.����ϵ��,1)/Nvl(E.סԺ��װ,1)) as ��������," & _
                " Nvl(A.ִ�в���ID,B.ִ�в���ID) as ִ�в���ID,Decode(Nvl(Instr(',4,5,6,7,',A.�շ����),0),0," & strExe2 & "," & strExe1 & ") as ִ��״̬," & _
                " B.�״�ʱ��,B.ĩ��ʱ��,Decode(Nvl(B.�Ʒ�״̬,0),-1,'����Ʒ�',0,'δ�Ʒ�'," & strState & ") as �Ʒ�״̬," & _
                " B.������,B.״̬˵��,B.���ͺ�,B.��¼��� as �������,A.��� as �������,C.������ĿID,C.�������," & _
                " B.ִ��״̬ as ִ��״̬ID,B.�Ʒ�״̬ as �Ʒ�״̬ID,A.��¼״̬,B.���ʱ��,B.�����,B.ִ��˵��,B.����ʱ��,B.������,B.����ʱ��" & _
                " From (" & strSub & ") A,����ҽ������ B,(" & strSQL & ") C,������ĿĿ¼ D,ҩƷ��� E,�շ���ĿĿ¼ F" & _
                " Where B.ҽ��ID=C.ID And C.������ĿID=D.ID" & IIF(mbln��������ִ��, "(+)", "") & " And C.�շ�ϸĿID=E.ҩƷID(+) And C.�շ�ϸĿID=F.ID(+)" & _
                " And A.NO(+)=B.NO And A.��¼����(+)=B.��¼���� And 0+A.ҽ�����(+)=B.ҽ��ID"
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
                strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
            End If
        End If
        
        strSQL = "Select  A.�������,A.�������," & _
            " A.ҽ��ID,A.���ID,A.�������,F.���� as �������," & IIF(mbln��������ִ��, IIF(InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0, "D.����", "Nvl(d.����, a.ҽ������)"), "D.����") & " as ������Ŀ,A.�걾��λ,A.��鷽��,A.����ʱ��,A.NO,A.��¼����,A.�������," & _
            " Nvl(G.����,B.����)||Decode(B.����,NULL,NULL,'('||B.����||')')||Decode(B.���,NULL,NULL,' '||B.���) as �շ���Ŀ," & _
            " A.��λ,A.�������� as ����,C.���� as ִ�п���,A.ִ��״̬,A.�״�ʱ��,A.ĩ��ʱ��,A.�Ʒ�״̬,A.������,A.״̬˵��,A.���ͺ�," & _
            " A.ִ�в���ID,A.ִ��״̬ID,A.�Ʒ�״̬ID,A.��¼״̬,D.��������,H.��������,A.���ʱ��,a.�����,a.ִ��˵��,a.����ʱ��,a.������,a.����ʱ��" & _
            " From (" & strSQL & ") A,�շ���ĿĿ¼ B,���ű� C,������ĿĿ¼ D,������Ŀ��� F,�շ���Ŀ���� G,�������� H" & _
            " Where A.�շ�ϸĿID=B.ID(+) And A.ִ�в���ID=C.ID(+) And A.������ĿID=D.ID" & IIF(mbln��������ִ��, "(+)", "") & " And A.�������=F.����(+)" & _
            " And A.�շ�ϸĿID=H.����ID(+) And A.�շ�ϸĿID=G.�շ�ϸĿID(+) And G.����(+)=1 And G.����(+)=" & IIF(gbytҩƷ������ʾ = 0, 1, 3) & _
            " Order by A.���ͺ� Desc,A.�������,A.�������,A.�������"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)))
        
        If Not rsTmp.EOF Then
            strSQL = "Select Max(a.ִ��ʱ��) As ִ��ʱ��, a.ҽ��id, a.���ͺ� From ����ҽ��ִ�� A, ����ҽ������ B, ����ҽ����¼ C" & vbNewLine & _
                        "Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And b.ҽ��id=c.id and (c.id=[1] or c.���id=[1])" & vbNewLine & _
                        "Group By a.ҽ��id, a.���ͺ�"
            Set rsִ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Name, IIF(Val(vsAdvice.TextMatrix(lngRow, COL_���ID)) = 0, Val(vsAdvice.TextMatrix(lngRow, COL_ID)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID))))
            
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                If InStr(str���ͺ� & ",", "," & NVL(rsTmp!���ͺ�, 0) & ",") = 0 Then
                    str���ͺ� = str���ͺ� & "," & NVL(rsTmp!���ͺ�, 0)
                End If
                .TextMatrix(i, COLSend("���ͺ�")) = NVL(rsTmp!���ͺ�, 0)
                .TextMatrix(i, COLSend("����ʱ��")) = Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
                .Cell(flexcpData, i, COLSend("����ʱ��")) = Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
                
                '����ҽ��
                If rsTmp!������� = "4" Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "��������-" & rsTmp!������Ŀ
                ElseIf InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "ҩƷҽ��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "��ҩ;��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And vsAdvice.TextMatrix(lngRow, COL_�������) = "K" Then
                    .TextMatrix(i, COLSend("����ҽ��")) = "��Ѫ;��-" & rsTmp!������Ŀ
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "�ɼ�����-" & rsTmp!������Ŀ
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "��ҩ�巨-" & rsTmp!������Ŀ
                    Else
                        .TextMatrix(i, COLSend("����ҽ��")) = "��ҩ�÷�-" & rsTmp!������Ŀ
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "������Ŀ-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "D" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "��鲿λ-" & NVL(rsTmp!�걾��λ) & "(" & NVL(rsTmp!��鷽��) & ")"
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "��������-" & rsTmp!������Ŀ
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, COLSend("����ҽ��")) = "������Ŀ-" & rsTmp!������Ŀ
                    End If
                Else
                    .TextMatrix(i, COLSend("����ҽ��")) = rsTmp!������� & "ҽ��-" & rsTmp!������Ŀ
                End If
               
                .TextMatrix(i, COLSend("���ݺ�")) = NVL(rsTmp!NO)
                .TextMatrix(i, COLSend("�շ���Ŀ")) = NVL(rsTmp!�շ���Ŀ)
                .TextMatrix(i, COLSend("��������")) = FormatEx(NVL(rsTmp!����), 5) & NVL(rsTmp!��λ)
                .TextMatrix(i, COLSend("�Ʒ�״̬")) = NVL(rsTmp!�Ʒ�״̬)
                .TextMatrix(i, COLSend("ִ��״̬")) = NVL(rsTmp!ִ��״̬)
                .TextMatrix(i, COLSend("ִ�п���")) = NVL(rsTmp!ִ�п���)
                .TextMatrix(i, COLSend("�״�ʱ��")) = Format(NVL(rsTmp!�״�ʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("ĩ��ʱ��")) = Format(NVL(rsTmp!ĩ��ʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("����ʱ��")) = Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("������")) = NVL(rsTmp!������)
                .TextMatrix(i, COLSend("����ʱ��")) = Format(NVL(rsTmp!����ʱ��), "yyyy-MM-dd HH:mm")
                
                .TextMatrix(i, COLSend("������")) = NVL(rsTmp!������)
                .TextMatrix(i, COLSend("״̬˵��")) = NVL(rsTmp!״̬˵��)
                If rsTmp!״̬˵�� & "" <> "" Then
                    bln״̬˵�� = True
                End If
                '������,����ִ�д���
                .TextMatrix(i, COLSend("ҽ��ID")) = rsTmp!ҽ��ID
                .TextMatrix(i, COLSend("���ID")) = NVL(rsTmp!���ID)
                .TextMatrix(i, COLSend("��¼����")) = NVL(rsTmp!��¼����, 0)
                .TextMatrix(i, COLSend("�������")) = Val("" & rsTmp!�������)
                .TextMatrix(i, COLSend("��¼״̬")) = NVL(rsTmp!��¼״̬, 0)
                .TextMatrix(i, COLSend("�������")) = NVL(rsTmp!�������)
                .TextMatrix(i, COLSend("��������")) = NVL(rsTmp!��������)
                .TextMatrix(i, COLSend("��������")) = NVL(rsTmp!��������, 0)
                .TextMatrix(i, COLSend("���ʱ��")) = Format(NVL(rsTmp!���ʱ��), "yyyy-MM-dd HH:mm")
                .TextMatrix(i, COLSend("ִ��ʱ��")) = Format(NVL(rsTmp!���ʱ��), "yyyy-MM-dd HH:mm")
                rsִ��.Filter = "ҽ��ID=" & rsTmp!ҽ��ID & " And ���ͺ�=" & NVL(rsTmp!���ͺ�, 0)
                If Not rsִ��.EOF Then .TextMatrix(i, COLSend("���ִ��ʱ��")) = Format(NVL(rsִ��!ִ��ʱ��), "yyyy-MM-dd HH:mm")
                
                .TextMatrix(i, COLSend("ִ����")) = NVL(rsTmp!�����)
                .TextMatrix(i, COLSend("ִ��˵��")) = NVL(rsTmp!ִ��˵��)
                .TextMatrix(i, COLSend("��Ѫ����")) = Val(NVL(rsTmp!��鷽��))  '��Ѫ��ҽ������鷽���洢��0-��Ѫ��1-��Ѫ
                .Cell(flexcpData, i, COLSend("�Ʒ�״̬")) = CStr(rsTmp!�Ʒ�״̬ID)
                .Cell(flexcpData, i, COLSend("ִ��״̬")) = Val(NVL(rsTmp!ִ��״̬ID, 0))
                .Cell(flexcpData, i, COLSend("ִ�п���")) = Val("" & rsTmp!ִ�в���ID)
                
                '��λԭ����
                If NVL(rsTmp!���ͺ�, 0) & "," & NVL(rsTmp!ҽ��ID) & "," & NVL(rsTmp!�շ���Ŀ) = strKey Then
                    lngKey = i
                End If
                
                If Val("" & rsTmp!ִ�в���ID) = mlng����ID Or Val("" & rsTmp!ִ�в���ID) = mlng����ID Or vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ������) = "��Ժ��ҩ" Then
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = &HD0FFFF
                    If lngKey = -1 Then
                        '�������λ����ǰ���У����Զ���λ�������������ִ�е���Ŀ
                        lngKey = i
                    End If
                End If
                If vsAdvice.TextMatrix(lngRow, COL_�������) = "K" And rsTmp!������� & "" = "K" Then
                    If gblnѪ��ϵͳ Then
                        lng��Ѫ = i
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        
        If lng��Ѫ <> -1 Then
            '��Ѫҽ������������Ŀ����Ϣ
            strSQL = "select b.���� as ������Ŀ,a.������ as ����,b.���㵥λ as ��λ,a.������Ŀid from ��Ѫ������Ŀ a,������ĿĿ¼ b where a.������Ŀid=b.id and a.ҽ��id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
            For i = 1 To rsTmp.RecordCount
                If Val(rsTmp!������ĿID & "") <> Val(vsAdvice.TextMatrix(lngRow, COL_������ĿID)) Then
                    .AddItem ""
                    For j = .FixedCols To .Cols - 1
                        .TextMatrix(.Rows - 1, j) = .TextMatrix(lng��Ѫ, j)
                    Next
                    .Cell(flexcpData, .Rows - 1, COLSend("����ʱ��")) = .Cell(flexcpData, lng��Ѫ, COLSend("����ʱ��"))
                    .Cell(flexcpData, .Rows - 1, COLSend("�Ʒ�״̬")) = .Cell(flexcpData, lng��Ѫ, COLSend("�Ʒ�״̬"))
                    .Cell(flexcpData, .Rows - 1, COLSend("ִ��״̬")) = .Cell(flexcpData, lng��Ѫ, COLSend("ִ��״̬"))
                    .Cell(flexcpData, .Rows - 1, COLSend("ִ�п���")) = .Cell(flexcpData, lng��Ѫ, COLSend("ִ�п���"))
                    .TextMatrix(.Rows - 1, COLSend("����ҽ��")) = "��Ѫҽ��-" & rsTmp!������Ŀ
                    .TextMatrix(.Rows - 1, COLSend("��������")) = FormatEx(NVL(rsTmp!����), 5) & NVL(rsTmp!��λ)
                    .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, lng��Ѫ, .FixedCols)
                Else
                    .TextMatrix(lng��Ѫ, COLSend("��������")) = FormatEx(NVL(rsTmp!����), 5) & NVL(rsTmp!��λ)
                End If
                rsTmp.MoveNext
            Next
        End If
        
        If str���ͺ� <> "" Then
            .AddItem "", 0
            .FixedRows = 2
            .Cell(flexcpText, 0, 0, 0, .Cols - 1) = " ������ " & UBound(Split(str���ͺ�, ",")) & " ��"
            .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
            .MergeRow(0) = True
            
            .Row = IIF(lngKey = -1, .FixedRows, lngKey + 1): .Col = COLSend("����ҽ��")
        Else
            .Row = IIF(lngKey = -1, .FixedRows, lngKey): .Col = COLSend("����ҽ��")
        End If
        .MergeCells = flexMergeFree
        .MergeCol(COLSend("���ͺ�")) = True
        .MergeCol(COLSend("����ʱ��")) = True
        .MergeCol(COLSend("���ݺ�")) = True
        .MergeCol(COLSend("����ҽ��")) = True
        .MergeCol(COLSend("�շ���Ŀ")) = True
        .MergeCol(COLSend("�״�ʱ��")) = True
        .MergeCol(COLSend("ĩ��ʱ��")) = True
        .MergeCol(COLSend("������")) = True
        .MergeCol(COLSend("״̬˵��")) = True
        
        .ColHidden(COLSend("״̬˵��")) = Not bln״̬˵��
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

Private Function LoadExecList(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long) As Boolean
'���ܣ���ȡָ��ҽ����ִ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strPre As String
    Dim rsѪ�� As ADODB.Recordset
    Dim bln��Ѫ As Boolean
    Dim intѪ���� As Integer
    
    On Error GoTo errH
    
    '������Ŀһ��ִ��ʱ��ִ������Ǽǵ���һ����Ŀ�ϡ���ɢ����ִ��ʱ���Ǽǵ�������Ŀ�ϡ�
    strSQL = "Select A.Ҫ��ʱ��,A.ִ��ʱ��,A.��������,D.���㵥λ,A.ִ��ժҪ,A.ִ����,A.�Ǽ�ʱ��,A.�Ǽ���,DECODE(NVL(A.ִ�н��,1),0,'δִ��',1,'���',2,'�ܾ�',3,'���') As ִ�н��,a.�˶���,a.�˶�ʱ��,d.��������,d.���,a.˵��,a.��¼��Դ as ��Դ" & _
        " From ����ҽ��ִ�� A,����ҽ������ B,����ҽ����¼ C,������ĿĿ¼ D" & _
        " Where A.ҽ��ID=[1] And A.���ͺ�=[2]" & _
        " And A.ҽ��ID=B.ҽ��ID And A.���ͺ�=B.���ͺ� And B.ҽ��ID=C.ID And C.������ĿID=D.ID" & IIF(mbln��������ִ��, "(+)", "") & _
        " Order by A.�Ǽ�ʱ�� Desc"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ��ִ��", "H����ҽ��ִ��")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    With vsExec
        strPre = .Cell(flexcpData, .Row, 0)
        .Redraw = flexRDNone
        .Rows = vsExec.FixedRows
        .Rows = vsExec.FixedRows + 1
        .Row = .FixedRows: .Col = .FixedCols
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            '��Ѫҽ���������̱䶯 70823
            If gblnѪ��ϵͳ And Val(rsTmp!�������� & "") = 8 And rsTmp!��� = "E" Then
                strSQL = "select zl_Get_��Ѫִ�д���(���id) as ���� from ����ҽ����¼ where id = [1]"
                Set rsѪ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
                If Not rsѪ��.EOF Then intѪ���� = Val(rsѪ��!���� & "")
                bln��Ѫ = True
            End If
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Format(rsTmp!Ҫ��ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 1) = Format(rsTmp!ִ��ʱ��, "yyyy-MM-dd HH:mm")
                If bln��Ѫ Then
                    .TextMatrix(i, 2) = FormatEx(Val(rsTmp!�������� & "") * intѪ����, 0) & " ��"
                Else
                    .TextMatrix(i, 2) = FormatEx(rsTmp!��������, 5) & " " & NVL(rsTmp!���㵥λ)
                End If
                .TextMatrix(i, 3) = NVL(rsTmp!ִ��ժҪ)
                .TextMatrix(i, 4) = NVL(rsTmp!ִ����)
                .TextMatrix(i, 5) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 6) = NVL(rsTmp!�Ǽ���)
                .TextMatrix(i, 7) = rsTmp!ִ�н�� & ""
                .TextMatrix(i, 8) = NVL(rsTmp!�˶���)
                .TextMatrix(i, 9) = Format(rsTmp!�˶�ʱ��, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, 10) = NVL(rsTmp!˵��)
                .TextMatrix(i, 11) = IIF(1 = Val(rsTmp!��Դ & ""), "�ƶ���", "PC��")
                
                .Cell(flexcpData, i, 0) = Format(rsTmp!Ҫ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                .Cell(flexcpData, i, 1) = Format(rsTmp!ִ��ʱ��, "yyyy-MM-dd HH:mm:ss")
       
                If .Cell(flexcpData, i, 0) = strPre Then .Row = i
                rsTmp.MoveNext
            Next
            rsTmp.MoveFirst
        End If
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
    End With
    With vsAppend
        If Not (.TextMatrix(.Row, COLSend("�������")) = "E" And .TextMatrix(.Row, COLSend("��������")) = "1" And Mid(gstrҽ���˶�, 2, 1) = "1" Or _
            (.TextMatrix(.Row, COLSend("�������")) = "E" And .TextMatrix(.Row, COLSend("��������")) = "8" Or .TextMatrix(.Row, COLSend("�������")) = "K") And Mid(gstrҽ���˶�, 1, 1) = "1") Then
            
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
'���ܣ���ʾָ����ҽ����ǩ����¼
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
        
        strSQL = "Select A.ǩ��ID,A.��������,B.ǩ��ʱ��,B.ǩ����,B.ʱ���," & _
            " Decode(A.��������,1,'�¿�ҽ��',3,'У��ҽ��',4,'����ҽ��',8,'ֹͣҽ��','��������') as ǩ������" & _
            " From ����ҽ��״̬ A,ҽ��ǩ����¼ B Where A.ҽ��ID=[1] And A.ǩ��ID=B.ID Order by B.ǩ��ʱ��"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ��״̬", "H����ҽ��״̬")
            strSQL = Replace(strSQL, "ҽ��ǩ����¼", "Hҽ��ǩ����¼")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = Val(rsTmp!ǩ��ID)
                .TextMatrix(i, 0) = rsTmp!ǩ������
                .Cell(flexcpData, i, 0) = Val(rsTmp!��������)
                .TextMatrix(i, 1) = Format(rsTmp!ǩ��ʱ��, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 2) = rsTmp!ǩ����
                .TextMatrix(i, 3) = Format(NVL(rsTmp!ʱ���), "yyyy-MM-dd HH:mm:ss")
                Set .Cell(flexcpPicture, i, 0) = frmIcons.imgSign.ListImages("ǩ��").Picture
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
'���ܣ���ʾָ����ҽ���ĵ��ݸ�������
'���أ�blnExist=ҽ���Ƿ���ڵ��ݸ�������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    
    blnExist = False
    rtfAppend.Text = "": rtfAppend.SelStart = 0
    
    On Error GoTo errH
    
    strSQL = "Select ��Ŀ,���� From ����ҽ������ Where ҽ��ID=[1] Order by ����"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfAppend
            Do While Not rsTmp.EOF
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp!��Ŀ & "��" & NVL(rsTmp!����)
                lngIdx = .Find(rsTmp!��Ŀ & "��", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp!��Ŀ & "��")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
                
                rsTmp.MoveNext
            Loop
            
            '��궨λ�ڵ�һ�����븽��
            rsTmp.MoveFirst
            lngIdx = .Find(rsTmp!��Ŀ & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp!��Ŀ & "��")
            
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
'���ܣ���ʾָ����ҽ����ִ�а�����Ϣ
'���أ�blnExist=ҽ���Ƿ����ִ�а�����Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    blnExist = False
    rtfInfo.Text = "": rtfInfo.SelStart = 0
    
    On Error GoTo errH
    
    With vsAdvice
        If InStr("D,F,G,", .TextMatrix(lngRow, COL_�������)) > 0 Or _
            .TextMatrix(lngRow, COL_�������) = "E" And InStr(",0,6,", "," & .TextMatrix(lngRow, COL_��������) & ",") > 0 Then
            
            If .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_��������)) = 6 Then
                strSQL = "Select a.����ʱ��,a.ִ�м�,a.ִ��˵�� From ����ҽ������ a,����ҽ����¼ b " & _
                        "Where a.ҽ��ID = b.ID And b.���ID=[1] And (a.ִ��˵�� is Not Null Or a.����ʱ�� is Not Null) And Rownum=1"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            Else
                strSQL = "Select ����ʱ��,ִ�м�,ִ��˵�� From ����ҽ������ Where ҽ��ID=[1] And (ִ��˵�� is Not Null Or ����ʱ�� is Not Null)"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)))
            End If
            If Not rsTmp.EOF Then
                strSQL = ""
                
                If Not IsNull(rsTmp!����ʱ��) Then
                    strSQL = strSQL & vbCrLf & "����ʱ�䣺" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                End If
                If Not IsNull(rsTmp!ִ�м�) Then
                    strSQL = strSQL & vbCrLf & "ִ�м䣺" & rsTmp!ִ�м�
                End If
                strSQL = strSQL & vbCrLf & NVL(rsTmp!ִ��˵��)
                
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
'���ܣ���ʾָ����ҽ���������Ϣ
'˵����ֻ������״̬ͨ����δͨ����ҽ��
'���أ��Ƿ���������Ϣ
    Dim strSQL As String
    Dim int���� As Integer
    Dim rsTmp As ADODB.Recordset
    Dim str����Ա As String
    Dim strʱ�� As String
    Dim strδ��ԭ�� As String '��Ѫҽ������
    
    On Error GoTo errH

    str����Ա = "����ˣ�": strʱ�� = "���ʱ�䣺"
    With vsAdvice
        If gblnѪ��ϵͳ And .TextMatrix(lngRow, COL_�������) = "K" Then
            If Val(.TextMatrix(lngRow, COL_ִ�б��)) = -1 Then '��ȡ���δ�õ�ԭ��
                strSQL = "Select ������Ա,����ʱ��,����˵�� From ����ҽ��״̬ Where ҽ��id = [1] And �������� = [2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), 17)
                If Not rsTmp.EOF Then
                    strδ��ԭ�� = "δ��ԭ��" & rsTmp!����˵��
                    strδ��ԭ�� = strδ��ԭ�� & "(����Ա��" & rsTmp!������Ա & "  ����ʱ�䣺" & Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS") & ")"
                End If
            End If
        End If
        
        Select Case .TextMatrix(lngRow, COL_���״̬)
            Case 2
                If gblnѪ��ϵͳ And .TextMatrix(lngRow, COL_�������) = "K" Then  '��Ѫҽ���������̱䶯 70823
                    int���� = 15 'Ѫ�����ͨ��
                    str����Ա = "Ѫ������ˣ�"
                    strʱ�� = "Ѫ�����ʱ�䣺"
                Else
                    int���� = 11
                End If
            Case 3
                int���� = 12
            Case 4
                int���� = 11
            Case 5
                int���� = 14
                str����Ա = "Ѫ������ˣ�"
                strʱ�� = "Ѫ�����ʱ�䣺"
        End Select
        rtfOther.Text = ""
        strSQL = "Select ������Ա,����ʱ�� From ����ҽ��״̬ Where ҽ��id = [1] And �������� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, COL_ID)), int����)
    End With
    
    If Not rsTmp.EOF Then
        strSQL = ""
        Do While Not rsTmp.EOF
            strSQL = IIF(strSQL = "", "", strSQL & vbCrLf) & str����Ա & rsTmp!������Ա & vbCrLf & _
                strʱ�� & Format(rsTmp!����ʱ�� & "", "YYYY-MM-DD HH:MM:SS")
            rsTmp.MoveNext
        Loop
        If strδ��ԭ�� <> "" Then
            strSQL = IIF(strSQL = "", "", strSQL & vbCrLf) & strδ��ԭ��
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
'���ܣ���ʾָ���е���Һ��ҩ����
'���أ�blnExist=ҽ���Ƿ���ڷ��͵���Һ�������ĵ�ҩƷ
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    blnExist = False
    On Error GoTo errH
    
    'ֻ�л�ʿվ�ŵ��ñ�����
    If gstr��Һ�������� <> "" Then
        If InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            strSQL = "Select 1 From ��Һ��ҩ��¼ Where ҽ��id = [1] and nvl(����״̬,0)<>12 And Rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_���ID)))
            If Not rsTmp.EOF Then '����Һ���Ҳ�����͵��������ģ�����������ҩ��¼�������ͳ���������Һ�����ĵĲ����������Ƿ������ҩ��¼
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
'���ܣ���ʾָ����ҽ�����Ի��˵��������ڴ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim vRoll As TYPE_AdviceRoll
    Dim lngҽ��ID As Long
    
    ReDim marrRollList(0)
    If Val(vsAdvice.TextMatrix(lngRow, COL_ID)) = 0 Then
        LoadRollList = True: Exit Function
    End If
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_���ID)) <> 0 Then
        lngҽ��ID = Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
    Else
        lngҽ��ID = Val(vsAdvice.TextMatrix(lngRow, COL_ID))
    End If
    
    '�ɻ���ҽ�������ͷ���,ҽ������Ĳ�������(�緢�ͺ��Զ�ֹͣ)
    '�������ɻ����Զ�ֹͣ,���˷���ʱ�Զ�����ֹͣ
    strSQL = " And (A.ID=[1] Or A.���ID=[1])"
    strSQL = _
        " Select Distinct 0 as ���ͺ�,B.������Ա as ��Ա,B.����ʱ�� as ʱ��,B.��������," & _
        " Decode(B.��������,4,'����ҽ��',5,'����ҽ��',6,'��ͣҽ��',7,'����ҽ��',8,'ֹͣҽ��',9,'ȷ��ֹͣ',10,'Ƥ�Խ��',13,'ͣ������') as ����" & _
        " From ����ҽ����¼ A,����ҽ��״̬ B" & _
        " Where A.ID=B.ҽ��ID" & strSQL & _
        " And (Nvl(A.ҽ����Ч,0)=0 And B.�������� Not IN(1,2,3)" & _
            " Or Nvl(A.ҽ����Ч,0)=1 And B.�������� Not IN(1,2,3,8))" & _
        " Union ALL" & _
        " Select Distinct B.���ͺ�,B.������ as ��Ա,B.����ʱ�� as ʱ��,0 as ��������,'����ҽ��' as ����" & _
        " From ����ҽ����¼ A,����ҽ������ B" & _
        " Where A.ID=B.ҽ��ID" & strSQL & _
        " Order by ʱ�� Desc,���ͺ�"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, lngҽ��ID)
    If Not rsTmp.EOF Then
        ReDim marrRollList(rsTmp.RecordCount)
        For i = 1 To rsTmp.RecordCount
            With vRoll
                .�������� = rsTmp!��������
                .���ͺ� = rsTmp!���ͺ�
                .����ʱ�� = Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm:ss")
                .������Ա = rsTmp!��Ա
                .�������� = "������:" & rsTmp!��Ա & ",ʱ��:" & Format(rsTmp!ʱ��, "yyyy-MM-dd HH:mm") & ",����:" & rsTmp!����
            End With
            marrRollList(i) = vRoll '��0�ĸ�����
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
'���ܣ��жϵ�һ���������Ƿ���Ի���
    Dim vRoll As TYPE_AdviceRoll
    Dim blnEnabled As Boolean
    
    If UBound(marrRollList) < 1 Then Exit Function
    vRoll = marrRollList(1)
    
    '��Ժ�����ﲡ�˲�������˲���
    If mintPState = ps��Ժ Or mintPState = ps���� Then Exit Function
    
    'Ԥ��Ժ���˽����Ի��˳�Ժҽ������
    If mintPState = psԤ�� Then
        blnEnabled = False
        If vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "Z" _
            And InStr(",5,6,11,", Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��������))) > 0 Then
            If vRoll.�������� = 0 And vRoll.���ͺ� <> 0 Then
                blnEnabled = True
            End If
        End If
        If Not blnEnabled Then Exit Function
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_������ĿID)) = 0 And InStr(vRoll.��������, "����ҽ��") > 0 Then
        Exit Function
    End If
    
    'ҽ��ֻ�ܻ������ѵ����ϡ�ֹͣ,��ͣ������,���������Ͳ���
    If mint���� <> 1 Then
        If Not ((vRoll.�������� = 0 Or InStr("45678", vRoll.��������) > 0 Or vRoll.�������� = 13) And vRoll.������Ա = UserInfo.����) Then
            Exit Function
        ElseIf mint���� = 2 Then
            If InStr("," & mstrǰ��IDs & ",", "," & vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID) & ",") = 0 Or vRoll.������Ա <> UserInfo.���� Then
                 Exit Function
            End If
        ElseIf mint���� = 0 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ǰ��ID)) <> 0 Then Exit Function
        End If
    End If
    
    RollFirstEnabled = True
End Function

Private Function LoadBillList() As Boolean
'���ܣ���ʾָ���е�ҽ�����Ϳ��Դ�ӡ�����Ƶ����ڲ˵���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim objMenu As CommandBarPopup
    Dim strTmp As String
    Dim blnBlood As Boolean, intBloodState As Integer '�¿�ʱ��ӡ��Ѫҽ��
    
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
    
    '��Ѫҽ�����뵥��ӡģʽ=0���¿�����Դ�ӡ
    intBloodState = 8
    If mint���뵥��ӡģʽ = 0 And vsAdvice.TextMatrix(vsAdvice.Row, COL_�������) = "K" Then
        blnBlood = True
        strTmp = ",-1,4,"
        If InStr(",1,2,3,", "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) & ",") > 0 Then
            intBloodState = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬))
        End If
    Else
        strTmp = ",-1,1,2,4,"
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 _
       Or InStr(strTmp, "," & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ҽ��״̬)) & ",") > 0 Then
        LoadBillList = True: Exit Function
    End If

        
    On Error GoTo errH
    strTmp = ""
    With vsAdvice
        If mint���뵥��ӡģʽ = 1 Or blnBlood = True Then
            '��Ҫ��������Ѫ���뵥������Ѫ֪ͨ��
            If gblnѪ��ϵͳ = True Then
                strTmp = " Union All " & vbNewLine & _
                " Select '-17', Decode(���, 1, '��Ѫ���뵥', 'ȡѪ֪ͨ��') ����, '', '', ���" & vbNewLine & _
                " From (Select Decode(a.��������, '8', Nvl(a.ִ�з���, 0), 0)+1 ���" & vbNewLine & _
                "       From ������ĿĿ¼ a, ����ҽ����¼ b, ����ҽ����¼ c" & vbNewLine & _
                "       Where Instr(',8,9,', ',' || a.�������� || ',') > 0 And a.Id = b.������Ŀid And b.������� = 'E' And b.���id = c.Id And" & vbNewLine & _
                "             c.Id = [1] And c.������� Is Not Null And c.������� = 'K' And c.ҽ��״̬ = [2])"
            Else
                strTmp = " Union All " & _
                " Select '-17','��Ѫ���뵥','','',0 From ����ҽ����¼ A Where  a.ID=[1] And A.������� is not null And A.������� = 'K' And A.ҽ��״̬=[2]"
            End If
        End If
        strSQL = "Select Distinct D.���,D.����,D.˵��,B.NO,0 ���" & _
            " From ����ҽ����¼ A,����ҽ������ B,��������Ӧ�� C,�����ļ��б� D" & _
            " Where C.������ĿID=A.������ĿID And a.ID=b.ҽ��ID " & _
            " And C.Ӧ�ó���=2 And C.�����ļ�ID=D.ID And D.����=7 And (a.ID=[1] or A.���ID=[1])" & _
            " And (A.������� is null Or A.������� <>'K')" & _
            strTmp & _
            " Order by ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Decode(Val(.TextMatrix(.Row, COL_���ID)), 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_���ID))), intBloodState)
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
    End With
     '���ֻ��һ�����Ƶ��ݿ��ã���ֱ�Ӽ��뵽ҽ���˵���
    If rsTmp.RecordCount = 1 Then
        objPopup.Visible = False
        objPopup.Category = "���ж�"
        Set objPopup = objMenu
    End If
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, IIF(rsTmp.RecordCount = 1, "��ӡ:", "") & rsTmp!����)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                '��ҩ�ļ巨�÷����ݺź���ҩ��һ������������ʾ����ҩ�÷������԰ѵ��ݵ�NOƴ��ȥ
                '���С��0��ʾʹ�ò����̶�����
                If Val(rsTmp!��� & "") < 0 Then
                    objControl.Parameter = "ZL1_INSIDE_1254_" & Abs(Val(rsTmp!��� & "")) & IIF(Val(rsTmp!��� & "") = 0, "", "_" & Val(rsTmp!��� & "")) '��Ӧ���Զ��屨����
                Else
                    objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" & "|" & rsTmp!NO '��Ӧ���Զ��屨����
                End If
                'If i > 1 Then objControl.Enabled = False 'һ����Ŀֻ������һ�����Ƶ���
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
'���ܣ���ʾָ���е�ҽ�����Ϳ��Դ�ӡ�����Ƶ����ڲ˵���(���Ͳ˵���)
'      �п��ܳ�����ӡǰ���η��͵ĵ��ݣ�����Ҫѡ���ͼ�¼No����ӡ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objControl As CommandBarControl
    Dim objPopup As CommandBarPopup
    Dim strTmp As String
    
    If mcbsMain Is Nothing Then LoadBillListOld = True: Exit Function
    Set objPopup = mcbsMain.FindControl(, conMenu_Report_ClinicBill, False, True) '���ܹ���������
    If objPopup Is Nothing Then LoadBillListOld = True: Exit Function
    
    objPopup.CommandBar.Controls.DeleteAll
    
    If tbcAppend.Selected.Tag <> "����" Then LoadBillListOld = True: Exit Function
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) = 0 _
        Or Val(vsAppend.TextMatrix(vsAppend.Row, COLSend("���ͺ�"))) = 0 Then
        LoadBillListOld = True: Exit Function
    End If
    
    On Error GoTo errH
    
    With vsAppend
        '�ų����뵥�´����Ѫҽ������������
        '�������ҽ�����ͺ�ŵ��ã������ٴ��ж����뵥��ӡģʽ
'        If mint���뵥��ӡģʽ = 1 Then
            If gblnѪ��ϵͳ = True Then
                strTmp = " Union All " & vbNewLine & _
                " Select '-17', Decode(���, 1, '��Ѫ���뵥', 'ȡѪ֪ͨ��') ����, '', ���" & vbNewLine & _
                " From (Select Decode(a.��������, '8', Nvl(a.ִ�з���, 0), 0)+1 ���" & vbNewLine & _
                "       From ������ĿĿ¼ a, ����ҽ����¼ b, ����ҽ����¼ c" & vbNewLine & _
                "       Where Instr(',8,9,', ',' || a.�������� || ',') > 0 And a.Id = b.������Ŀid And b.������� = 'E' And b.���id = c.Id And" & vbNewLine & _
                "             c.Id = [3] And c.������� Is Not Null And c.������� = 'K' And c.ҽ��״̬ = 8)"
            Else
                strTmp = " Union All " & _
                " Select '-17','��Ѫ���뵥','',0 From ����ҽ����¼ A Where  a.ID=[3] And A.������� is not null And A.������� = 'K' And A.ҽ��״̬=8"
            End If
'        End If
        strSQL = "Select Distinct D.���,D.����,D.˵��,0 ���" & _
            " From ����ҽ������ A,����ҽ����¼ B,��������Ӧ�� C,�����ļ��б� D" & _
            " Where A.���ͺ�=[1] And A.NO=[2]" & _
            " And A.ҽ��ID=B.ID And B.������ĿID=C.������ĿID" & _
            " And C.Ӧ�ó���=2 And C.�����ļ�ID=D.ID And D.����=7" & _
            " And (b.������� is null Or b.������� <>'K')" & _
            strTmp & _
            " Order by ���"
        If mblnMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, Val(.TextMatrix(.Row, COLSend("���ͺ�"))), .TextMatrix(.Row, COLSend("���ݺ�")), Val(.TextMatrix(.Row, COLSend("ҽ��ID"))))
    End With
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Report_ClinicBill * 100# + i, rsTmp!����)
                If i <= 10 Then
                    objControl.Caption = objControl.Caption & "(&" & i - 1 & ")"
                ElseIf i <= 36 Then
                    objControl.Caption = objControl.Caption & "(&" & Chr(i - 11 + Asc("A")) & ")"
                End If
                '���С��0��ʾʹ�ò����̶�����
                If Val(rsTmp!��� & "") < 0 Then
                    objControl.Parameter = "ZL1_INSIDE_1254_" & Abs(Val(rsTmp!��� & "")) & IIF(Val(rsTmp!��� & "") = 0, "", "_" & Val(rsTmp!��� & "")) '��Ӧ���Զ��屨����
                Else
                    objControl.Parameter = "ZLCISBILL" & Format(rsTmp!���, "00000") & "-1" '��Ӧ���Զ��屨����
                End If
                'If i > 1 Then objControl.Enabled = False 'һ����Ŀֻ������һ�����Ƶ���
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
    Dim blnBloldExec As Boolean '��Ѫִ��
    If Not Me.Visible Then Exit Sub
    If NewRow = OldRow Then Exit Sub
    
    With vsAppend
        If NewCol >= .FixedCols And NewRow >= .FixedRows Then
            If .Redraw <> flexRDNone Then
                If tbcAppend.Selected.Tag = "����" And vsAppend.Cols = COLSend.Count Then
                    If mint���� = 1 And Val(.TextMatrix(NewRow, COLSend("���ͺ�"))) <> 0 And (InStr(",5,6,7,", .TextMatrix(NewRow, COLSend("�������"))) = 0 Or (mbln��������ִ�� And .TextMatrix(NewRow, COLSend("�������")) = "")) Then
                        If Not (.TextMatrix(NewRow, COLSend("�������")) = "Z" And Val(.TextMatrix(NewRow, COLSend("��������"))) <> 0) Then
                            If Val(.Cell(flexcpData, NewRow, COLSend("ִ�п���"))) = mlng����ID Or Val(.Cell(flexcpData, NewRow, COLSend("ִ�п���"))) = mlng����ID Or vsAdvice.TextMatrix(vsAdvice.Row, COL_ִ������) = "��Ժ��ҩ" Then
                                blnExec = True
                            End If
                            
                            If gblnѪ��ϵͳ And .TextMatrix(NewRow, COLSend("�������")) = "E" And Val(.TextMatrix(NewRow, COLSend("��������"))) = 8 Then
                                '�����̵���Ѫҽ��������Ѫִ�еǼ�(ѪҺҽ���ļ�鷽��=1-��Ѫ)
                                blnBloldExec = IsUseBloodAdvice
                                blnExec = True
                            End If
                        End If
                    End If
                    vsAppend.Enabled = False '�ؼ�λ�ñ仯,��λ�ñ仯,���������������Ч
                    '��Ҫ����Ѫ��ִ�к�����ҽ��ִ�У�ִ����ʾ״̬����һ��(֮ǰ����Ѫִ�У��л���������ҽ���Ĵ���)
                    If blnExec = True And blnBloldExec = False And picBlood.Tag = "�ɼ�" Then
                        If Not mobjFrmBlood Is Nothing Then mblnShowExec = mobjFrmBlood.IsShowExec
                    End If
                    Call SetExecShow(blnExec, mblnShowExec, blnBloldExec)
                    vsAppend.Enabled = True
                    Me.Refresh
                    
                    '��ȡִ���б�
                    If mblnShowExec And blnExec And blnBloldExec = False Then
                        Call LoadExecList(Val(.TextMatrix(NewRow, COLSend("ҽ��ID"))), Val(.TextMatrix(NewRow, COLSend("���ͺ�"))))
                    Else
                        vsExec.Rows = vsExec.FixedRows
                        vsExec.Rows = vsExec.FixedRows + 1
                        vsExec.Row = vsExec.FixedRows
                    End If
                    '��Ѫִ���б��ȡ
                    If blnExec = True And blnBloldExec = True Then
                        If Not mobjFrmBlood Is Nothing Then
                            Call mobjFrmBlood.zlRefresh(Me, glngSys, pסԺҽ������, Val(.TextMatrix(NewRow, COLSend("ҽ��ID"))), mlngҽ������ID, GetInsidePrivs(pסԺҽ������), 2, mlng����ID, mblnMoved, mlngFontSize)
                        End If
                    End If
                End If
                    
                '��ʾ�ɴ�ӡ�����Ƶ���
                Call LoadBillListOld
            End If
        End If
    End With
End Sub

Private Sub vsAppend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBarPopup
    
    With vsAppend
        If Button = 2 And tbcAppend.Selected.Tag = "����" Then
            If mcbsMain Is Nothing Then Exit Sub
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    Set objPopup = mcbsMain.FindControl(, conMenu_Report_ClinicBill, False, True) '���ܹ���������
                    If Not objPopup Is Nothing Then
                        '���û�����ݣ��������ݴ���1������¸��µ������ͼ�¼�ĵ��ݣ���Ϊ���ж�����ݵĻ���ѡ��ҽ���Ż���ֵ�
                        If objPopup.CommandBar.Controls.Count = 0 Or objPopup.CommandBar.Controls.Count > 1 Then Call LoadBillListOld
                        If objPopup.CommandBar.Controls.Count > 0 Then
                            'ShowPopup���ᴥ��InitCommandsPopup�¼�
                            mintBillPrint = 1
                            objPopup.CommandBar.ShowPopup
                        End If
                    End If
                End If
            End If
        ElseIf Button = 2 And tbcAppend.Selected.Tag = "ǩ��" Then
            If mcbsMain Is Nothing Then Exit Sub
            If Between(.MouseRow, .FixedRows, .Rows - 1) Then
                If Between(.MouseCol, .FixedCols, .Cols - 1) Then
                    Set objPopup = mcbsMain.FindControl(, conMenu_Tool_Sign, False, True) '���ܹ���������
                    If Not objPopup Is Nothing Then
                        If objPopup.CommandBar.Controls.Count > 0 Then
                            'ShowPopup���ᴥ��InitCommandsPopup�¼�
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
    
    '��Ϊ����ͬ,��ȡ����ʱ�ᶪʧ��,Resize��ָ�
    picAppend.Tag = "��ִ��"
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

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'˵��:PASS �е� ��RowInһ����ҩ�� ��˷�����ͬ,�޸Ĵ˷���Ҳ��Ҫͬ���޸� PASSͬ������
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
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
'���ܣ�����������������
'������Index=�����ӹ�������(0,1)
    Dim blnOK As Boolean, lngҽ��ID As Long
    Dim strCommon As String, intAtom As Integer
    
    '���÷��ò�������
    On Error Resume Next
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Sub
    If gobjInExse Is Nothing Then
        Set gobjInExse = CreateObject("zl9InExse.clsInExse")
        If gobjInExse Is Nothing Then Exit Sub
    End If
    err.Clear: On Error GoTo 0
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
    End With
        
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & OS.ComputerName
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    
    If Index = conMenu_Edit_ChargeDelApply Then
        'ҽ��վֻ���������빦�ܣ���ֻ������ҩƷ
        blnOK = gobjInExse.CallReCharge(mfrmParent, gcnOracle, gstrDBUser, glngSys, 0, IIF(mint���� = 1, 0, 2), mlng����ID, GetInsidePrivs(1133), mlng����ID, , lngҽ��ID)
    ElseIf Index = conMenu_Edit_ChargeDelAudit Then
        blnOK = gobjInExse.CallReCharge(mfrmParent, gcnOracle, gstrDBUser, glngSys, 1, 0, mlng����ID, GetInsidePrivs(1133), mlng����ID)
    End If
    
    Call GlobalDeleteAtom(intAtom)
    
    If blnOK Then RaiseEvent RequestRefresh(False)
End Sub

Private Sub FuncApplyModi()
'���ܣ��޸����뵥
    Dim strSQL As String, rsTmp As Recordset
    With vsAdvice
        '���ж��Ƿ����Զ������뵥
        strSQL = "Select �ļ�ID From ҽ�����뵥�ļ� Where ҽ��ID=[1] And RowNum<2"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_���ID)) = 0, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_���ID))))
        If rsTmp.RecordCount > 0 Then
            FuncApplyCustom 1, Val(rsTmp!�ļ�ID)
        Else
                        If Val(.TextMatrix(.Row, COL_ҽ��״̬)) <> 1 Then
                MsgBox "�������޸��ѷ��͵����롣", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(.TextMatrix(.Row, COL_��������)) = 6 And .TextMatrix(.Row, COL_�������) = "E" Then
                Call FuncApplyLIS(Val(.TextMatrix(.Row, COL_�������)))
            ElseIf .TextMatrix(.Row, COL_�������) = "D" Then
                Call FuncApplyPACS(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_�������)))
            ElseIf .TextMatrix(.Row, COL_�������) = "K" Then
                If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)) = 1 And Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��鷽��)) = 1 Then
                    Call FuncApplyBlood(4)
                Else
                    Call FuncApplyBlood(1)
                End If
            ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                Call FuncApplyOperation(1)
            ElseIf Val(.TextMatrix(.Row, COL_��������)) = 7 And .TextMatrix(.Row, COL_�������) = "Z" Then
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
'���ܣ��鿴���뵥
    Dim lngҽ��ID As Long
    Dim lngNo As Long
    Dim strSQL As String, rsTmp As Recordset
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        lngNo = Val(.TextMatrix(.Row, COL_�������))
        
        If lngҽ��ID <> 0 And lngNo <> 0 Then
            strSQL = "Select �ļ�ID From ҽ�����뵥�ļ� Where ҽ��ID=[1] And RowNum<2"
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(Val(.TextMatrix(.Row, COL_���ID)) = 0, lngҽ��ID, Val(.TextMatrix(.Row, COL_���ID))))
            If rsTmp.RecordCount > 0 Then
                FuncApplyCustom 2, Val(rsTmp!�ļ�ID)
            Else
                If .TextMatrix(.Row, COL_�������) = "K" Then
                    Call FuncApplyBlood(2)
                ElseIf .TextMatrix(.Row, COL_�������) = "F" Then
                    Call FuncApplyOperation(2)
                ElseIf Val(.TextMatrix(.Row, COL_��������)) = 7 And .TextMatrix(.Row, COL_�������) = "Z" Then
                    Call FuncApplyConsultation(2)
                ElseIf .TextMatrix(.Row, COL_�������) = "D" Then
                    '���
                    If Val(Mid(gstrInUseApp, 1, 1)) = 1 Then
                        Call ShowApply���(Me, lngNo)
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

Private Sub FuncApplyPACS(ByVal lngҽ��ID As Long, ByRef lng������� As Long)
'���ܣ����ü�����뵥
'������lngҽ��ID=�޸����뵥ʱ��ǰ�е�ҽ��ID,lng������� =��ǰ�޸��е��������

    Dim bln��ҽ As Boolean
    Dim str���� As String
    Dim blnSucceed As Boolean
    Dim strMsg As String
    Dim lngNo As Long
    
    If CheckAdviceAddModi(IIF(lngҽ��ID = 0, 0, 1)) = False Then Exit Sub
    
    If lngҽ��ID & "_" & lng������� = "0_0" Then
        If Not FuncPathAdd() Then Exit Sub
    End If
    
    '��ϼ��
    If InStr(mstr�����Ժ���, "D") > 0 Then
        bln��ҽ = Sys.DeptHaveProperty(mlng����ID, "��ҽ��")
        str���� = IIF(bln��ҽ, "2,12", "2")
        If Not ExistsDiagNoses(mlng����ID, mlng��ҳID, str����) Then
            strMsg = "���˵���Ժ��ϻ�û�����룬�������벡�˵���Ժ������´����ҽ����"
        End If
        If strMsg <> "" Then
            If InStr(";" & mMainPrivs & ";", ";��ҳ����;") > 0 Then
                vsAdvice.Refresh
                MsgBox strMsg & vbCrLf & vbCrLf & "�밴 [ȷ��] �������������档", vbInformation, gstrSysName
                blnSucceed = True
                RaiseEvent EditDiagnose(Me, mlng����ID, mlng��ҳID, mlng����ID, str����, blnSucceed)
                vsAdvice.Refresh
                If Not blnSucceed Then Exit Sub
            Else
                vsAdvice.Refresh
                MsgBox strMsg, vbInformation, gstrSysName
                vsAdvice.Refresh: Exit Sub
            End If
        End If
    End If
    lngNo = ApplyInPacs(Me, lng�������, mlng����ID, mlng��ҳID, Val(mbytӤ��), mlng��������, lngҽ��ID, mlngҽ������ID, mlng����ID, mlng����ID, mobjVBA, mobjScript, mrsDefine, mclsMipModule, , mlngǰ��ID)
    If lngNo <> 0 Then Call LoadAdvice
    
    If mlng·��״̬ = 1 And Not gobjPath Is Nothing And lngNo <> 0 Then
        Call FuncPathSet(lngNo)
    End If
End Sub

Private Sub FuncApplyLIS(ByVal lng������� As Long)
'���ܣ����ü�������������뵥�ͼ���ҽ��
'������lng�������=�޸����뵥ʱ���������
    Dim arrSQL As Variant, i As Long, blnTrans As Boolean, strSQL As String
    Dim strResult As String, strDiag As String, strDept As String, strErr As String
    Dim rsPati As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset 'ע��˱�����Ҫ����,��LisInfoTrans��������ֵ
    Dim rsTemp As ADODB.Recordset
    Dim lngҽ��ID As Long, lng���ID As Long, lng��� As Long
    Dim lngִ�п���ID As Long, lng�ɼ�����ID As Long, lng������ĿID As Long, lng�ɼ���ĿID As Long
    Dim str����Ƽ����� As String, str�ɼ��Ƽ����� As String, str����ִ������ As String, str�ɼ�ִ������ As String
    Dim str������Ŀ As String, str�ɼ����� As String, str�걾 As String, str���� As String
    Dim strCurDate As String, strҽ������ As String, strҽ��IDs As String, blnCancel As Boolean
    Dim strDelIDs As String, arrDelID() As String
    Dim Y As Long, j As Long
    Dim str���� As String, str���� As String
    Dim bln��ҽ As Boolean, blnSucceed As Boolean
    Dim str���� As String
    Dim arrAppend As Variant
    Dim lng��������ID As Long
    Dim lng������� As Long
    Dim str��� As String
    Dim lng��ҽ��ID As Long '����ҽ��ID����ֵ���˷ѣ�������ύ����ʱ�������ҽ��ID
    Dim strҽ��ID As String, str���ID As String
    Dim varID As Variant
    Dim strTmp As String
    Dim bln���Ѷ��� As Boolean
    Dim vMsg As VbMsgBoxResult
    Dim strItems As String
    Dim strTabAdvice As String
    Dim blnCheckItem As Boolean 'ҽ���ܿؼ��
    Dim rsPrice As ADODB.Recordset
    Dim strժҪ As String, strMsg As String
    Dim dat��ʼִ��ʱ�� As Date
    Dim str��ʼִ��ʱ�� As String
    Dim dat��ǰʱ�� As Date
    Dim datTurn As Date
    Dim rsLISInfo As ADODB.Recordset
    Dim lng������� As Long
    
    If lng������� = 0 Then
        If Not FuncPathAdd() Then Exit Sub
    End If
    
    If CheckAdviceAddModi(IIF(lng������� = 0, 0, 1), , datTurn) = False Then Exit Sub
    
    Set rsPati = GetPatiInfo(mlng����ID, mlng��ҳID)
    If rsPati.RecordCount = 0 Then
        MsgBox "δ����ȷ��ȡ������Ϣ��", vbInformation, gstrSysName
        Exit Sub
    End If
    If lng������� <> 0 Then
        strDiag = GetAdviceDiag(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)))
    End If
    lng��������ID = Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2)
    strDept = Sys.RowValue("���ű�", lng��������ID, "����")
    Call InitObjLis(pסԺҽ��վ)
    If gobjLIS Is Nothing Then Exit Sub
    Call CreatePlugInOK(pסԺҽ���´�, mint����)
    
    On Error GoTo errH
     
    '������ѡ��ļ�����Ŀ��ʽ����: �������ID1,ִ�п���ID1,����ʱ��1,������Ŀ����1,�걾1,����ҽ��1,�ɼ���ʽ������ĿID 1;�������ID2,ִ�п���ID2,����ʱ��2,������Ŀ����2,�걾2,����ҽ��2,�ɼ���ʽ������ĿID 2;.....
    strResult = gobjLIS.ShowLisApplicationForm(mfrmParent, lng�������, mlng����ID, mbytӤ��, mlng��ҳID, rsPati!����, "" & rsPati!�Ա�, "" & rsPati!����, IIF(mlng�������� = 1, 1, 2), _
        Val("" & rsPati!�����), Val("" & rsPati!סԺ��), Val("" & rsPati!������), strDiag, UserInfo.����, UserInfo.����ID, UserInfo.������, lng��������ID, strDept, blnCancel, strErr)
     
    If strErr <> "" Then
        MsgBox "����ӿ��ڲ�����" & strErr, vbInformation, gstrSysName
    ElseIf blnCancel Then
        Exit Sub    'ȡ�����˳�
    Else
        arrSQL = Array()
        '�޸����뵥ʱ����ɾ���ɵ�ҽ��
        If lng������� <> 0 Then
            strҽ��IDs = GetAdivceBy�������(lng�������)
            For i = 0 To UBound(Split(strҽ��IDs, ","))
                '����ɾ��ǰ��ҽӿ�
                On Error Resume Next
                If Not gobjPlugIn Is Nothing Then
                    If gobjPlugIn.AdviceDeletBefor(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(Split(strҽ��IDs, ",")(i)), mint����) = False Then
                        If err.Number = 0 Then Exit Sub
                    End If
                    Call zlPlugInErrH(err, "AdviceDeletBefor")
                End If
                If err.Number <> 0 Then err.Clear
                On Error GoTo errH
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Delete(" & Split(strҽ��IDs, ",")(i) & ",1)"
                strDelIDs = strDelIDs & "," & Split(strҽ��IDs, ",")(i)
            Next
            strDelIDs = Mid(strDelIDs, 2)
        End If
        
        If strResult <> "" Then
            '��ϼ��
            If InStr(mstr�����Ժ���, "C") > 0 Then
                bln��ҽ = Sys.DeptHaveProperty(mlng����ID, "��ҽ��")
                str���� = IIF(bln��ҽ, "2,12", "2")
                If Not ExistsDiagNoses(mlng����ID, mlng��ҳID, str����) Then
                    strMsg = "���˵���Ժ��ϻ�û�����룬�������벡�˵���Ժ������´����ҽ����"
                End If
                If strMsg <> "" Then
                    If InStr(";" & mMainPrivs & ";", ";��ҳ����;") > 0 Then
                        vsAdvice.Refresh
                        MsgBox strMsg & vbCrLf & vbCrLf & "�밴 [ȷ��] �������������档", vbInformation, gstrSysName
                        blnSucceed = True
                        RaiseEvent EditDiagnose(Me, mlng����ID, mlng��ҳID, mlng����ID, str����, blnSucceed)
                        vsAdvice.Refresh
                        If Not blnSucceed Then Exit Sub
                    Else
                        vsAdvice.Refresh
                        MsgBox strMsg, vbInformation, gstrSysName
                        vsAdvice.Refresh: Exit Sub
                    End If
                End If
            End If
            
            bln���Ѷ��� = True
            
            If mint���� <> 0 Then
                If gclsInsure.GetCapability(supportʵʱ���, mlng����ID, mint����) Then
                    blnCheckItem = True
                End If
            End If

            If strDiag <> "" Then
                str��� = GetDiag�������(strDiag)
                If str��� <> "" Then
                    str��� = "���뵥���<Split2>0<Split2><Split2>" & str���
                End If
            End If
             
            dat��ǰʱ�� = zlDatabase.Currentdate()
            strCurDate = "To_Date('" & Format(dat��ǰʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
            lng��� = GetMaxAdviceNO(mlng����ID, mlng��ҳID, mbytӤ��)
            lng������� = -1
            '�ڸ÷����ж�rsLISInfo, rsTmp��ֵ
            Call LisInfoTrans(strResult, rsLISInfo, rsTmp)
                        
            'ֻ��������
            For i = 1 To rsLISInfo.RecordCount
                
                If lng������� <> Val(rsLISInfo!��� & "") Then
                    lng������� = Val(rsLISInfo!��� & "")
                    lng������� = Get�������
                End If
        
                lng��ҽ��ID = lng��ҽ��ID + 1
                str���ID = "<FAKEID>" & lng��ҽ��ID & "</FAKEID>"
                lng���ID = lng��ҽ��ID
                lng�ɼ�����ID = Val(rsLISInfo!�ɼ�����ID & "")
                lngִ�п���ID = Val(rsLISInfo!ִ�п���ID & "")
                str��ʼִ��ʱ�� = rsLISInfo!��ʼִ��ʱ�� & ""
                str�걾 = rsLISInfo!�걾 & ""
                str���� = rsLISInfo!���� & ""
                str���� = rsLISInfo!���� & ""
                str���� = rsLISInfo!���� & ""
                lng�ɼ���ĿID = Val(rsLISInfo!�ɼ���ĿID & "")
                lng������ĿID = Val(rsLISInfo!������ĿID & "")
                                
                dat��ʼִ��ʱ�� = CDate(str��ʼִ��ʱ��)
                str��ʼִ��ʱ�� = "To_Date('" & Format(dat��ʼִ��ʱ��, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
                
                '�ж��Ƿ��ǲ�¼ҽ��
                If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Or (mbytӤ�� > 0 And datTurn <> CDate(0)) Then
                    str���� = "2"
                ElseIf DateDiff("n", dat��ʼִ��ʱ��, dat��ǰʱ��) > gint��¼��� Then
                    str���� = "2"
                End If
                    
                'a.�Ȳ�������ҽ�� ���뵥�������ĵļ���ҽ��ֻ��һ��������ĿID
                rsTmp.Filter = "ID=" & lng������ĿID
                str������Ŀ = rsTmp!���� & ""
                str����Ƽ����� = Val("" & rsTmp!�Ƽ�����)
                str����ִ������ = IIF("" & rsTmp!ִ�п��� = "", "NULL", "" & rsTmp!ִ�п���)
                strҽ������ = str������Ŀ & IIF("" = rsLISInfo!ʱ������ & "", "", "(" & rsLISInfo!ʱ������ & ")")
                lng��� = lng��� + 1
                strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(lng������ĿID) & "||2")
                blnCancel = CheckLISAppAdvice(2, mlng����ID, mlng��ҳID, mint����, "C", lng������ĿID, lng��������ID, UserInfo.����, lngִ�п���ID, Val(rsTmp!ִ�п��� & ""), strժҪ & "||0||0|| ||0")
                If Not blnCancel Then Exit Sub
                
                lng��ҽ��ID = lng��ҽ��ID + 1
                strҽ��ID = "<FAKEID>" & lng��ҽ��ID & "</FAKEID>"
                lngҽ��ID = lng��ҽ��ID
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & _
                    strҽ��ID & "," & str���ID & "," & lng��� & ",2," & mlng����ID & "," & _
                    mlng��ҳID & "," & mbytӤ�� & ",1,1,'C'," & _
                    lng������ĿID & ",Null,Null,Null,1," & _
                    "'" & strҽ������ & "',Null," & "'" & str�걾 & "','һ����',Null," & _
                    "Null,Null,Null," & str����Ƽ����� & "," & lngִ�п���ID & _
                    "," & str����ִ������ & "," & str���� & "," & str��ʼִ��ʱ�� & ",Null," & mlng����ID & "," & _
                    lng��������ID & ",'" & UserInfo.���� & "'," & strCurDate & ",NULL," & ZVal(mlngǰ��ID) & "," & _
                    "NULL,0,Null," & IIF(strժҪ = "", "Null", "'" & strժҪ & "'") & ",'" & UserInfo.���� & "'" & _
                    ",Null,Null,Null,Null," & lng������� & ",null,null,null,null,null,'" & rsLISInfo!ʱ��ID & "')"
                
                strItems = strItems & "," & lng������ĿID & ":" & lngִ�п���ID
                
                If blnCheckItem Then
                    strTabAdvice = _
                        "select " & lngҽ��ID & " as ID," & lng��� & " as ���," & lng���ID & " as ���ID,'C' as �������," & lng������ĿID & " as ������ĿID," & _
                        lng������ĿID & " as ������ĿID,-null as �շ�ϸĿID, 1 As ����, 0 As ����,'" & str�걾 & "' as �걾��λ,'' As ��鷽��," & _
                        "0 as ִ�б��," & Val("" & rsTmp!�Ƽ�����) & " as �Ƽ�����, 0 As ��������," & Val("" & rsTmp!ִ�п���) & " As ִ������," & lngִ�п���ID & " as ִ�п���id from dual"
                End If

                'b.�ٲ����ɼ�����ҽ��
                rsTmp.Filter = "ID=" & lng�ɼ���ĿID
                str�ɼ����� = rsTmp!���� & ""
                str�ɼ��Ƽ����� = Val("" & rsTmp!�Ƽ�����)
                str�ɼ�ִ������ = "" & rsTmp!ִ�п���
                strҽ������ = AdviceMakeText(str������Ŀ, str�ɼ�����, str�걾)
                If "" <> rsLISInfo!ʱ������ & "" Then strҽ������ = strҽ������ & "(" & rsLISInfo!ʱ������ & ")"
                lng��� = lng��� + 1
                strժҪ = ""
                strժҪ = gclsInsure.GetItemInfo(mint����, mlng����ID, 0, "", 0, "", CStr(lng�ɼ���ĿID) & "||2")
                blnCancel = CheckLISAppAdvice(2, mlng����ID, mlng��ҳID, mint����, "E", lng�ɼ���ĿID, lng��������ID, UserInfo.����, lng�ɼ�����ID, Val(rsTmp!ִ�п��� & ""), strժҪ & "||0||0|| ||0")
                If Not blnCancel Then Exit Sub
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ����¼_Insert(" & _
                    str���ID & ",Null," & lng��� & ",2," & mlng����ID & "," & _
                    mlng��ҳID & "," & mbytӤ�� & ",1,1,'E'," & _
                    lng�ɼ���ĿID & ",Null,Null,Null,1," & _
                    "'" & strҽ������ & "','" & str���� & "'," & "'" & str�걾 & "','һ����',Null," & _
                    "Null,Null,Null," & str�ɼ��Ƽ����� & "," & lng�ɼ�����ID & _
                    "," & str�ɼ�ִ������ & "," & str���� & "," & str��ʼִ��ʱ�� & ",Null," & mlng����ID & "," & _
                    lng��������ID & ",'" & UserInfo.���� & "'," & strCurDate & ",NULL," & ZVal(mlngǰ��ID) & "," & _
                    "NULL,0,Null," & IIF(strժҪ = "", "Null", "'" & strժҪ & "'") & ",'" & UserInfo.���� & "'" & _
                    ",Null,Null,Null,Null," & lng������� & ",null,null,null,null,null,'" & rsLISInfo!ʱ��ID & "')"
                
                strItems = strItems & "," & lng�ɼ���ĿID & ":" & lng�ɼ�����ID
                
                If blnCheckItem Then
                    strTabAdvice = strTabAdvice & " Union ALL " & _
                        "select " & lng���ID & " as ID," & lng��� & " as ���,-null as ���ID,'E' as �������," & lng������ĿID & " as ������ĿID," & _
                        lng�ɼ���ĿID & " as ������ĿID,-null as �շ�ϸĿID, 1 As ����, 0 As ����,'" & str�걾 & "' as �걾��λ,'' As ��鷽��," & _
                        "0 as ִ�б��," & Val("" & rsTmp!�Ƽ�����) & " as �Ƽ�����, 0 As ��������," & Val("" & rsTmp!ִ�п���) & " As ִ������," & lng�ɼ�����ID & " as ִ�п���id from dual"
                End If
                
                'ҽ��������
                If gintҽ������ = 2 Then bln���Ѷ��� = True
                strMsg = CheckAdviceInsure(mint����, bln���Ѷ���, mlng����ID, mlng��������, "", Mid(strItems, 2), Left(strҽ������, 50), mlng����ID)
                If strMsg <> "" Then
                    If gintҽ������ = 1 Then
                        vMsg = frmMsgBox.ShowMsgBox(strMsg & vbCrLf & vbCrLf & "Ҫ��������ҽ����", Me)
                        If vMsg = vbNo Or vMsg = vbCancel Then Exit Sub
                        If vMsg = vbIgnore Then bln���Ѷ��� = False
                    ElseIf gintҽ������ = 2 Then
                        MsgBox strMsg & vbCrLf & vbCrLf & "���Ⱥ������Ա��ϵ��������ҽ�����������档", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    strMsg = ""
                End If
                
                'ҽ���ܿ�ʵʱ��⣺�״�����(����)���߸���ʱ���
                If blnCheckItem Then
                    If MakePriceRecord���뵥("12", mlng����ID, mlng��ҳID, strTabAdvice, strItems, rsPati!�ѱ� & "", lng��������ID, rsPrice) Then
                        If Not gclsInsure.CheckItem(mint����, 1, 0, rsPrice) Then
                            MsgBox "ҽ�������δͨ(ִ��Insure.CheckItem�ӿ�)�������´��LIS���뵥���ܱ��档", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If str���� <> "" And str��� <> "" Then
                    str���� = str��� & "<Split1>" & str����
                ElseIf str���� = "" And str��� <> "" Then
                    str���� = str���
                End If
                
                '�������븽�������������Ȳ���ҽ��
                If str���� <> "" Then
                    arrAppend = Split(str����, "<Split1>")
                    For j = 0 To UBound(arrAppend)
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_����ҽ������_Insert(" & str���ID & "," & _
                            "'" & Split(arrAppend(j), "<Split2>")(0) & "'," & Val(Split(arrAppend(j), "<Split2>")(1)) & "," & _
                            j + 1 & "," & ZVal(Split(arrAppend(j), "<Split2>")(2)) & ",'" & Replace(Split(arrAppend(j), "<Split2>")(3), "'", "''") & "'" & _
                            IIF(j = 0, ",1", "") & ")"
                        lng������� = j + 1
                    Next
                End If
                If strDiag <> "" Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_�������ҽ��_Insert(" & str���ID & ",'" & strDiag & "')"
                End If
                rsLISInfo.MoveNext
            Next
        End If
        
        '�����в�����ʵ��ҽ��ID
        If lng��ҽ��ID > 0 Then
            For j = 1 To lng��ҽ��ID
                Y = zlDatabase.GetNextID("����ҽ����¼")
                If j = 1 Then
                    strҽ��IDs = ""
                    strҽ��IDs = Y
                Else
                    strҽ��IDs = strҽ��IDs & "," & Y
                End If
            Next
            varID = Split(strҽ��IDs, ",")
            
            For i = 0 To UBound(arrSQL)
                strTmp = arrSQL(i)
                
                If InStr(strTmp, "<FAKEID>") > 0 Then
                    j = Mid(strTmp, (InStr(strTmp, "<FAKEID>") + 8), (InStr(strTmp, "</FAKEID>") - InStr(strTmp, "<FAKEID>") - 8))
                    strTmp = Replace(strTmp, "<FAKEID>" & j & "</FAKEID>", varID(j - 1))
                    
                    If InStr(strTmp, "<FAKEID>") > 0 Then '����滻����
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
        
        '�ٴ�·���ж�
         If mlng·��״̬ = 1 And Not gobjPath Is Nothing And lng������� <> 0 Then
             Call FuncPathSet(lng�������)
         End If
         Call ZLHIS_CIS_001(mclsMipModule, mlng����ID, mstr����, mstrסԺ��, , IIF(mlng�������� = 1, 1, 2), mlng��ҳID, mlng����ID, , mlng����ID, "", , mstr����, _
               lng���ID, str����, 1, "E", "", UserInfo.����, Format(dat��ʼִ��ʱ��, "yyyy-MM-dd HH:MM:00"), lng��������ID, "", , , "")
    
         'ˢ��ҽ��
         Call RefreshData
       
        '����ɾ������ҽӿ�
        On Error Resume Next
        arrDelID = Split(strDelIDs, ",")
        For i = 0 To UBound(arrDelID)
            If Val(arrDelID(i)) <> 0 Then
                If Not gobjPlugIn Is Nothing Then
                    Call gobjPlugIn.AdviceDeleted(glngSys, pסԺҽ���´�, mlng����ID, mlng��ҳID, Val(arrDelID(i)), mint����)
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

Private Function GetAdivceBy�������(ByVal lng������� As Long) As String
'���ܣ�����������Ż�ȡ���м��ҽ��ID�����ɼ�ҽ��ID��
    Dim i As Long, strTmp As String
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_�������)) = lng������� Then
                If Val(.TextMatrix(i, COL_��������)) = 6 And .TextMatrix(i, COL_�������) = "E" Then
                    strTmp = strTmp & "," & .TextMatrix(i, COL_ID)
                End If
            End If
        Next
        GetAdivceBy������� = Mid(strTmp, 2)
    End With
End Function

Private Function AdviceMakeText(ByVal str���� As String, ByVal str�ɼ� As String, ByVal str�걾 As String) As String
'���ܣ���������ҽ����ҽ������
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
               
    'ȷ���Ƿ���
    blnDefine = Not mrsDefine Is Nothing And Not mobjVBA Is Nothing
    If blnDefine Then
        mrsDefine.Filter = "�������='C'"
        If mrsDefine.EOF Then
            blnDefine = False
        ElseIf Trim(NVL(mrsDefine!ҽ������)) = "" Then
            blnDefine = False
        End If
    End If
    
    If Not blnDefine Then
        strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
    Else
        strText = mrsDefine!ҽ������
        If InStr(strText, "[������Ŀ]") > 0 Then
            strField = str����
            strText = Replace(strText, "[������Ŀ]", """" & strField & """")
        End If
        If InStr(strText, "[����걾]") > 0 Then
            strField = str�걾
            strText = Replace(strText, "[����걾]", """" & strField & """")
        End If
        If InStr(strText, "[�ɼ�����]") > 0 Then
            strField = str�ɼ�
            strText = Replace(strText, "[�ɼ�����]", """" & strField & """")
        End If
        
        '����ҽ������
        On Error Resume Next
        strText = mobjVBA.Eval(strText)
        If mobjVBA.Error.Number <> 0 Then
            strText = str���� & IIF(str�걾 <> "", "(" & str�걾 & ")", "")
        End If
        err.Clear: On Error GoTo 0
    End If
        
    AdviceMakeText = strText
End Function

Private Sub GetAdvicesSameSend(ByVal lng���ͺ� As Long, ByRef strLIS As String, ByRef strALL As String, Optional ByVal str������� As String = "C")
'���ܣ����ݷ��ͺŻ�ȡһ���͵�ҽ������ID
'������strLIS ���Σ�����ҽ��ID����strALL����ҽ��ID��
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIDs As String, i As Long
    Dim str����IDs As String
    
    strSQL = "Select b.�������, Nvl(b.���id,b.Id) As id" & vbNewLine & _
        "From ����ҽ������ A, ����ҽ����¼ B" & vbNewLine & _
        "Where a.ҽ��id = b.Id And a.���ͺ� =[1] And b.����id =[2] And b.��ҳid =[3]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���ͺ�, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        If InStr("," & strIDs & ",", "," & rsTmp!ID & ",") = 0 Then
            strIDs = strIDs & "," & rsTmp!ID
            If rsTmp!������� & "" = str������� Then
                str����IDs = str����IDs & "," & rsTmp!ID
            End If
        End If
        rsTmp.MoveNext
    Next
    
    strALL = Mid(strIDs, 2)
    strLIS = Mid(str����IDs, 2)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub PrintLisReport(ByVal lngPatiDeptID As Long, objFrm As Object)
    'LIS�������鱨���ӡ
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
    '��Ѫִ�е���ӡ
    If InitObjBlood(True) = True Then
        Call gobjPublicBlood.ShowBloodInstantRptPrint(objFrm, lngAdviceID)
    End If
End Sub

Private Function CheckPatiIsAduit() As Boolean
'���ܣ���鲡���Ƿ�ʼ���
    Dim rsTmp As Recordset, strSQL As String
    Dim int��˱�־ As Integer
    
    
    
    If mblnBatch Then CheckPatiIsAduit = True: Exit Function
    strSQL = "Select a.��˱�־ From ������ҳ a" & _
                " Where a.����ID=[1] And a.��ҳID=[2]"
    On err GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������˼��", mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        If Val("" & rsTmp!��˱�־) >= 1 And gbyt������˷�ʽ = 1 Then
            MsgBox "�ò��˵ķ���������˻��Ѿ���ˣ����������ҽ���ͷ��á�", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPatiIsAduit = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub SetFontSize(ByVal bytSize As Byte)
'����:����ҽ���嵥�������С
'���:bytSize��0-С(ȱʡ)��1-��
    mlngFontSize = IIF(bytSize = 0, 9, 12)
    
    '����vsFlexGrid�ؼ���ʹ�ø��Ի�����ʱ��Ӵ��п�����ڴ�����μ����ǲ���������,��Ҫ��getForm��������
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
    
    If mvarCond.��ʾģʽ = 0 Then
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_����)
    Else
        Call Grid.SetFontSize(vsAdvice, mlngFontSize, col_ҽ������)
    End If
    
    Call Grid.SetFontSize(vsAppend, mlngFontSize)
    If tbcAppend.Selected.Tag = "����" Then '��һ�е�����Ҫ���⴦��
        vsAppend.Cell(flexcpFontSize, 0, 0, 0, vsAppend.Cols - 1) = mlngFontSize
    End If
    
    Call Grid.SetFontSize(vsExec, mlngFontSize)
    Call Grid.SetFontSize(vsfAdivceDetail, mlngFontSize)
    
    If Not mfrmCompoundMedicine Is Nothing Then
        Call Grid.SetFontSize(mfrmCompoundMedicine.vsSend, mlngFontSize)
        Call Grid.SetFontSize(mfrmCompoundMedicine.vsExec, mlngFontSize)
    End If
    
    'ѪҺִ�к�ѪҺ��ϸ����
    If Not mobjFrmBloodList Is Nothing Then
        If mobjFrmBloodList.Visible = True Then Call mobjFrmBloodList.SetFontSize(mlngFontSize)
    End If
    
    If Not mobjFrmBlood Is Nothing Then
        If mobjFrmBlood.Visible = True Then Call mobjFrmBlood.SetFontSize(mlngFontSize)
    End If
    
    Call SetRTFFont(0)
End Sub

Private Function CheckBabyEdit(ByVal lngBaby As Long) As Integer
'���ܣ����ĸӤ�����Ƿ�����༭
'���أ�0������༭��1=Ӥ�����Ҳ�����༭����ҽ����2=���˿��Ҳ�����༭Ӥ��ҽ��
'������lngBabyӤ�����
    CheckBabyEdit = 0
    If mlngӤ������ID <> 0 And mstrӤ�� <> "" Then
        If (mlngӤ������ID = mlngҽ������ID Or mlngӤ������ID = mlngҽ������ID) And lngBaby = 0 Then
            CheckBabyEdit = 1
        ElseIf (mlng����ID = mlngҽ������ID Or mlng����ID = mlngҽ������ID) And lngBaby > 0 Then
            CheckBabyEdit = 2
        End If
    End If
End Function

Private Function CheckDelAdivceOfPathItem(ByVal lngҽ��ID As Long) As Boolean
'���ܣ����ҽ����Ӧ��·����Ŀ�Ƿ�����ɾ��������Ǳ���ִ�е���Ŀ����Ӧ��ҽ��������Ҫ����ԭ��ѡ�񲢸��±���ԭ��
'       ��ӹ�����ԭ��Ĳ������
'���أ�True-����ɾ����ҽ����false-����ɾ��
'����:lngҽ��ID
    Dim blnCancel As Boolean, blnMust As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset, rsAdvice As ADODB.Recordset
    Dim strReason As String
    Dim vPoint As PointAPI
    Dim strTemp As String
    Dim arrTmp As Variant
    Dim arrSQL As Variant
    Dim i As Long

    '1.���·����Ŀ
    strSQL = "Select  c.Id as ִ��Id, c.����,c.����ԭ��,d.ִ�з�ʽ,c.����,c.�׶�ID,c.·����¼ID,c.��ĿID " & _
             " From ����·��ҽ�� B, ����·��ִ�� C, �ٴ�·����Ŀ D" & vbNewLine & _
             "Where b.����ҽ��Id=[1] And b.·��ִ��id = c.Id And d.Id = c.��Ŀid And d.ִ�з�ʽ in (1,2,4)"

    On Error GoTo errH

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "���·��ҽ��", lngҽ��ID)

    If rsTmp.RecordCount < 1 Then
        CheckDelAdivceOfPathItem = True
        Exit Function    '�� �������ɵ�·��ҽ��
    End If
    '2.���ҽ���ܷ�ɾ��
    '��·����Ŀ������У�Ե�δ���ϵ�����ҽ������ʾ����ֹɾ��    ҽ��״̬ ��3-��У��
    strSQL = "Select a.����ҽ��ID,b.ҽ��״̬ " & vbNewLine & _
             "From ����·��ҽ�� A, ����ҽ����¼ B" & vbNewLine & _
             "Where a.·��ִ��id = [1] And a.����ҽ��id = b.Id  And b.ҽ��״̬>1 and b.ҽ��״̬<>4"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "���·��ҽ��", rsTmp!ִ��Id)

    If rsAdvice.RecordCount > 0 Then
        MsgBox "ɾ��ҽ�����ڵ�·����Ŀ�д�����У�Ե�δ���ϵ�ҽ�����������ϸ�ҽ������ִ�д˲�����", vbInformation, gstrSysName
        CheckDelAdivceOfPathItem = False
        Exit Function
    End If
    

    If mint���� = 1 Then
        '�����Ѿ�����˵�ҽ�����������޸�ɾ����
        strSQL = "Select b.����ҽ��ID From ����·��ҽ�� B, ����ҽ����¼ C Where b.·��ִ��id = [1] And b.����ҽ��id = c.Id And c.����ҽ�� Like '%/%'"
        Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "���·��ҽ��", rsTmp!ִ��Id)
        If rsAdvice.RecordCount > 0 Then
            MsgBox "ɾ����ҽ���д���ҽ������ҽ������ȡ��������ִ�д˲�����", vbInformation, gstrSysName
            CheckDelAdivceOfPathItem = False
            Exit Function
        End If
    End If
    
    '����ִ�з�ʽ �����Ƿ��б�Ҫ��ӱ���ԭ��
    blnMust = CheckPathItemIsMust(Val(rsTmp!ִ�з�ʽ & ""), Val("" & rsTmp!����), Val("" & rsTmp!·����¼id), Val("" & rsTmp!�׶�id), Val("" & rsTmp!��ĿID))
    If Not blnMust Then CheckDelAdivceOfPathItem = True: Exit Function
    
    '----------------------------
    '3.�������ɵ���Ŀ��д����ԭ��
    For i = 1 To rsTmp.RecordCount
        If rsTmp!����ԭ�� & "" = "" Then
            strTemp = strTemp & rsTmp!ִ��Id & "," & rsTmp!���� & ";"
        End If
        rsTmp.MoveNext
    Next
    
    If strTemp = "" Then
        CheckDelAdivceOfPathItem = True
        Exit Function
    Else
        strTemp = Mid(strTemp, 1, Len(strTemp) - 1)
    End If

    strSQL = "Select b.���� as ����,a.���� as ID,a.����,a.����,a.���� From ���쳣��ԭ�� a,���쳣��ԭ�� b" & _
             " Where a.����=1 And a.ĩ��=1 And a.�ϼ�=b.���� And b.ĩ��=0 " & _
             " Order by ����,a.����"
    vPoint = zlControl.GetCoordPos(vsAdvice.hwnd, vsAdvice.CellLeft, vsAdvice.CellTop)

    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "���쳣��ԭ��", True, , , True, True, True, _
                                      vPoint.X, vPoint.Y, vsAdvice.RowHeight(vsAdvice.Row), blnCancel, False, True)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "ϵͳû�г�ʼ���쳣��ԭ������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
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
            arrSQL(UBound(arrSQL)) = "Zl_����·������_Update(" & arrTmp(0) & ",'" & arrTmp(1) & "',Null ,Null,Null,Null,Null,'" & strReason & "')"
        Next
        '�����������������ԭ�����ʧ�ܣ�ҽ������ɾ�����ٴ�ɾ��ʱ����������ӱ���ԭ���ſ�ɾ����
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
'���ܣ�����ҽ�����͡���ť����ʽ--�˵���/������
'˵����
'      ��ܴ���Ĳ˵�(mcbsMain)�����޸ģ��������Ƿ��޸ĸ��� frmDockInAdvice ��ʲôģʽ,��ʲôģʽ�� mblnInsideTools ���֣�
'      mblnInsideTools  =True �޸�cbsSub�Ĺ�������=False �޸�mcbsMain�Ĺ�������
    Dim objControl As CommandBarControl
    Dim objCtlTmp As CommandBarControl
    Dim objMenuBar As CommandBarPopup
    Dim strPrivs As String
    Dim strPara As String
    Dim objCbs As Object
    Dim i As Long
    
    If mcbsMain Is Nothing Then Exit Sub
    
    On Error GoTo errH
    If gstr��Һ�������� <> "" Then
        strPrivs = GetInsidePrivs(pסԺҽ������)
        If InStr(";" & strPrivs & ";", ";����ҩ������;") = 0 Or InStr(";" & strPrivs & ";", ";����ҩ�Ƴ���;") = 0 Then
            strPrivs = ""
        End If
    End If
    
    '�˵����
    Set objMenuBar = mcbsMain.ActiveMenuBar.Controls(IIF(mblnInsideTools, 2, 3))
    For i = objMenuBar.CommandBar.Controls.Count To 1 Step -1
        If objMenuBar.CommandBar.Controls(i).ID = conMenu_Edit_Send Then
            objMenuBar.CommandBar.Controls(i).Delete: Exit For
        End If
    Next i
    strPara = zlDatabase.GetPara("��Դ����", glngSys, p��Һ��������, "")
    With objMenuBar.CommandBar.Controls
        Set objControl = .Find(, conMenu_Edit_Audit)
        If Not objControl Is Nothing Then
            If strPrivs <> "" Then
                Set objMenuBar = .Add(xtpControlButtonPopup, conMenu_Edit_Send, "ҽ������(&G)", objControl.Index + 1)
                Set objCtlTmp = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "��������ҽ��(&G)")
                If InStr("," & strPara & ",", "," & mlng����ID & ",") > 0 Or strPara = "" Then
                    objCtlTmp.Caption = "����ҽ��(������Һ)(&G)"
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "����ҽ��(����Һ)(&I)")
                Else
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "���;���Ӫ��ҩƷ(&I)")
                End If
                objControl.IconId = conMenu_Edit_Send
            Else
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "ҽ������(&G)", objControl.Index + 1): objControl.ToolTipText = ""
            End If
        End If
    End With
    
    '���������
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
                Set objMenuBar = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "����", objControl.Index + 1): objMenuBar.Style = xtpButtonIconAndCaption
                Set objCtlTmp = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Send, "��������ҽ��")
                
                If InStr("," & strPara & ",", "," & mlng����ID & ",") > 0 Or strPara = "" Then
                    objCtlTmp.Caption = "����ҽ��(������Һ)"
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "����ҽ��(����Һ)")
                Else
                    Set objControl = objMenuBar.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_SendInfusion, "���;���Ӫ��ҩƷ")
                End If
                
                objControl.IconId = conMenu_Edit_Send
            Else
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", objControl.Index + 1): objControl.Style = xtpButtonIconAndCaption: objControl.ToolTipText = "ҽ������"
            End If
        End If
    End With
    
    '�ȼ�
    With objCbs.KeyBindings
        .Add FCONTROL, vbKeyG, conMenu_Edit_Send   '��������ҽ��
        .Add 0, vbKeyF2, conMenu_Edit_SendInfusion '������ҺҩƷҽ��
    End With
    
    If mblnInsideTools Then objCbs.RecalcLayout
    
    mcbsMain.RecalcLayout
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub AddToolBarInDoctor()
'���ܣ����ù�������ť����Ӧ��ҽ���˵�����Ĺ������İ�ť���Ƚ���ɾ�������
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
    '���������
    If mblnInsideTools Then
        Set objCbs = cbsSub
        cbsSub(2).Visible = Not mblnHideFilter
    Else
        Set objCbs = mcbsMain
    End If

    '�ҵ�Ҫ��ӵ�λ��
    lngIdx = 0
    For Each objControl In objCbs(2).Controls '�����ǰ������һ��Control
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
    
    'ɾ����������ť
    For i = objCbs(2).Controls.Count To 1 Step -1
        If InStr(strTmp, "," & objCbs(2).Controls(i).ID & ",") > 0 Then
            objCbs(2).Controls(i).Delete
        Else
            If mblnInsideTools Then objCbs(2).Controls(i).Delete
        End If
    Next i

    With objCbs(2).Controls
        If mvarCond.����ģʽ <> 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_NewItem, "�¿�", lngIdx + 1): objPopup.BeginGroup = True
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_NewItem * 10# + 1, "�¿�")
                    objControl.IconId = conMenu_Edit_NewItem
                .Add xtpControlButton, conMenu_Edit_Modify, "�޸�"
                .Add xtpControlButton, conMenu_Edit_Delete, "ɾ��"
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
        End If
        
        If mint���� = 0 Then 'ֻ��סԺҽ������վ����ʱ�����⼸����ť
            strTmp = ""
            intTmp = Val(Mid(gstrInUseApp, 1, 1))
            If intTmp = 1 Then strTmp = strTmp & ",�������:" & conMenu_Edit_PacsApply
            intTmp = Val(Mid(gstrInUseApp, 2, 1))
            If intTmp = 1 And Not gobjLIS Is Nothing Then strTmp = strTmp & ",��������:" & conMenu_Edit_LISApply
            intTmp = Val(Mid(gstrInUseApp, 3, 1))
            If intTmp = 1 Then strTmp = strTmp & ",��Ѫ����:" & conMenu_Edit_BloodApply
            intTmp = Val(Mid(gstrInUseApp, 4, 1))
            If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_OperationApply
            intTmp = Val(Mid(gstrInUseApp, 5, 1))
            If intTmp = 1 Then strTmp = strTmp & ",��������:" & conMenu_Edit_ConsultationApply
            Get�Զ������뵥 2, mstr�Զ������뵥IDs
            If mstr�Զ������뵥IDs <> "" Then
                For i = 0 To UBound(Split(mstr�Զ������뵥IDs, "|"))
                    strTmp = strTmp & "," & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(1) & ":" & (conMenu_Edit_ApplyCustom * 100# + i) & ":" & Split(Split(mstr�Զ������뵥IDs, "|")(i), ",")(0)
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
        
        If mvarCond.����ģʽ = 3 And mint���� = 0 Then 'ֻ��סԺҽ������վ����ʱ�����⼸����ť
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyModi, "�޸�", lngIdx + 1)
                objControl.IconId = 3002
                objControl.ToolTipText = "�޸�����"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyView, "�鿴", objControl.Index + 1)
                objControl.IconId = 102
                objControl.ToolTipText = "�鿴����"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyDel, "ȡ��", objControl.Index + 1)
                objControl.IconId = 3004
                objControl.ToolTipText = "ȡ������"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        End If
        
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "����", lngIdx + 1)
            objControl.BeginGroup = True
            objControl.Style = xtpButtonIconAndCaption
        lngIdx = objControl.Index
        
        If mint���� = 0 Then
            If mvarCond.����ģʽ <> 3 Then  'ֻ��סԺҽ������վ����ʱ�����⼸����ť
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Stop, "ֹͣ", objControl.Index + 1)
                    objControl.Style = xtpButtonIconAndCaption
                If gblnѪ��ϵͳ Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_TraReaction, "��Ѫ��Ӧ", objControl.Index + 1)
                        objControl.IconId = 4113
                        objControl.Style = xtpButtonIconAndCaption
                End If
            End If
            
            If InStr(GetInsidePrivs(pסԺҽ���´�), "�����������") = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "����(&G)", objControl.Index + 1)
                    objControl.IconId = conMenu_Edit_Send
                    objControl.BeginGroup = True
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
            Else
                Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "����", objControl.Index + 1)
                With objPopup.CommandBar.Controls
                    .Add xtpControlButton, conMenu_Edit_SendBilling, "סԺ����"
                    .Add xtpControlButton, conMenu_Edit_SendCharge, "�����շ�"
                End With
                objPopup.Style = xtpButtonIconAndCaption
                lngIdx = objPopup.Index
            End If
        End If
        
        If mint���� = 1 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Send, "����", lngIdx + 1)
            objControl.Style = xtpButtonIconAndCaption
            If Val(zlDatabase.GetPara("����ǰ�Զ�У��", glngSys, pסԺҽ������, 0)) = 1 Then
                objControl.BeginGroup = True
            End If
            lngIdx = objControl.Index
        End If
        
        If mint���� = 2 Then
            If InStr(GetInsidePrivs(pסԺҽ���´�), "�����������") = 0 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_SendBilling, "��������(&G)", objControl.Index + 1): objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_Send
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            Else
                Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Send, "��������", objControl.Index + 1): objPopup.BeginGroup = True
                With objPopup.CommandBar.Controls
                    .Add xtpControlButton, conMenu_Edit_SendBilling, "סԺ����"
                    .Add xtpControlButton, conMenu_Edit_SendCharge, "�����շ�"
                End With
                objPopup.Style = xtpButtonIconAndCaption
                lngIdx = objPopup.Index
            End If
        End If
        If mint���� = 1 And mvarCond.����ģʽ <> 3 Then
            Set objControl = .Add(xtpControlButton, conMenu_Edit_Audit, "У��", lngIdx + 1)
            objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        End If
        
        Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Untread, "����", lngIdx + 1)
            objPopup.Style = xtpButtonIconAndCaption
        lngIdx = objPopup.Index
        
        If mvarCond.����ģʽ = 3 Then
            Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend, "����", lngIdx + 1): objPopup.BeginGroup = True
                objPopup.IconId = conMenu_Manage_Report
                objPopup.ToolTipText = "���ı���"
                
            With objPopup.CommandBar.Controls
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 1, "������ʽ(&B)"): objControl.IconId = 102
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 6, "�����ʽ(&P)"): objControl.IconId = 102
                If gobjExchange Is Nothing And mint���� <> 1 Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 4, "���Ѳ���(&R)")
                        objControl.BeginGroup = True
                    .Add xtpControlButton, conMenu_Edit_Compend * 10# + 5, "�Զ����(&A)"
                End If
            End With
            objPopup.Style = xtpButtonIconAndCaption
            lngIdx = objPopup.Index
            If gobjExchange Is Nothing Then
                If mint���� <> 1 Then
                    Set objPopup = .Add(xtpControlSplitButtonPopup, conMenu_Edit_Compend * 10# + 2, "��ӡ����", lngIdx + 1)
                        objPopup.IconId = 103
                        objPopup.Style = xtpButtonIconAndCaption
                        With objPopup.CommandBar.Controls
                            Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������"): objControl.IconId = 102
                            objControl.Style = xtpButtonIconAndCaption
                        End With
                    lngIdx = objPopup.Index
                Else
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_Compend * 10# + 3, "Ԥ������(&V)", lngIdx + 1)
                    objControl.IconId = 102
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
                End If
            End If
    
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "��Ƭ", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "��Ƭ����"
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            Set objControl = .Add(xtpControlButton, conMenu_Edit_MarkKeyMap, "�ؼ�ͼ��", lngIdx + 1)
                objControl.BeginGroup = True
                objControl.IconId = conMenu_Edit_MarkMap
                objControl.ToolTipText = "�ؼ�ͼ��"
                objControl.Style = xtpButtonIconAndCaption
            Set objControl = .Add(xtpControlButton, conMenu_Manage_ReportLisView, "���", objControl.Index + 1): objControl.IconId = conMenu_Manage_ReportLisView
                objControl.ToolTipText = "���������"
                objControl.Style = xtpButtonIconAndCaption
            lngIdx = objControl.Index
        Else
            If mint���� = 1 Then '
                Set objControl = .Add(xtpControlButton, conMenu_Edit_Price, "�Ƽ�", lngIdx + 1)
                objControl.Style = xtpButtonIconAndCaption
                Set objControl = .Add(xtpControlButton, conMenu_Edit_ReStop, "ȷ��ֹͣ", objControl.Index + 1)
                objControl.Style = xtpButtonIconAndCaption
                If Not mblnInsideTools Then
                    Set objControl = .Add(xtpControlButton, conMenu_Report_Reports, "ִ�е�", objControl.Index + 1): objControl.IconId = 3205
                    objControl.Style = xtpButtonIconAndCaption
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "ִ�еǼ�", objControl.Index + 1): objControl.IconId = 3587
                    objControl.Style = xtpButtonIconAndCaption
                End If
                Set objControl = .Add(xtpControlButton, conMenu_Manage_ThingAudit, "�˶�", objControl.Index + 1)
                objControl.Style = xtpButtonIconAndCaption
                lngIdx = objControl.Index
            End If
    
            If mint���� = 0 Then
                If Not mblnInsideTools Then
                    Set objControl = .Add(xtpControlButton, conMenu_Edit_BatExecute, "ִ�еǼ�", lngIdx + 1): objControl.IconId = 3587
                        objControl.Style = xtpButtonIconAndCaption
                        lngIdx = objControl.Index
                End If
            End If
            
            If mint���� = 1 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_AdvicePrice, "��ҽ������", lngIdx + 1)
                    objControl.IconId = conMenu_Edit_Price
                    objControl.Style = xtpButtonIconAndCaption
                    lngIdx = objControl.Index
            End If
            If mint���� <> 2 Then
                Set objControl = .Add(xtpControlButton, conMenu_Edit_ChargeDelApply, "����", lngIdx + 1)
                    objControl.IconId = conMenu_Edit_ChargeOff
                    objControl.Style = xtpButtonIconAndCaption
            End If
                
            If mblnPass Then  '������ҩ�˵�
                Call gobjPass.zlPassCommandBarAdd(mobjPassMap, objCbs(2).Controls, conMenu_Edit_MediAudit, conMenu_Edit_FeeAudit, objControl.Index + 1)
            End If
            
            Set objControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "ǩ��", objControl.Index + 1): objControl.BeginGroup = True
                objControl.IconId = conMenu_Tool_Sign
                objControl.Style = xtpButtonIconAndCaption
        End If
    End With
    
    If mblnInsideTools Then objCbs.RecalcLayout
    
    mcbsMain.RecalcLayout
    
    If mint���� = 1 And mvarCond.����ģʽ <> 3 Then
        Call SetSendCommandBar
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetRollAdviceIDs(ByVal lngҽ��ID As Long, ByVal bytMode As Byte, Optional ByVal lng�������� As Integer, Optional ByVal dat����ʱ�� As Date, Optional ByVal blnѪ����� As Boolean) As String
'���ܣ���ȡ��Ҫ���˵�ҽ����¼��
'������lngҽ��ID һ��ҽ������ID
'      bytMode-1.����һ��ҽ����¼��2.������������ҽ����¼��(����)
'      lng��������-��������ʱ����
'      dat����ʱ��-��������ʱ����
'      blnѪ����� ����Ѫ������ʱ���ô��룬��Ӧ�Ĳ���Ϊ �����������ϲ���

    Dim rsTmp       As ADODB.Recordset
    Dim strSQL      As String
    Dim strTmp As String
    
    On Error GoTo errH

    If bytMode = 1 Then
        strSQL = "Select f_List2str(Cast(Collect(ID || '') As t_Strlist)) As ҽ��ids  From ����ҽ����¼ Where ID =[1] Or ���id =[1]"
    Else
        strSQL = "Select f_List2str(Cast(Collect(ID || '') As t_Strlist)) As ҽ��ids From ����ҽ����¼ Where" & _
                 " Not (ҽ��״̬ = 8 And ҽ����Ч = 1) And" & vbNewLine & _
                 "      ID In (Select ҽ��id" & vbNewLine & _
                 "             From ����ҽ��״̬" & vbNewLine & _
                 "             Where (��������, ����ʱ��, ������Ա) In (Select ��������, ����ʱ��, ������Ա" & vbNewLine & _
                 "                                          From ����ҽ��״̬" & vbNewLine & _
                 "                                          Where ҽ��id = [1] And ����ʱ�� = [2] And �������� = [3]))"
                 
        If blnѪ����� Then
            strSQL = "Select f_List2str(Cast(Collect(ID || '') As t_Strlist)) As ҽ��ids From ����ҽ����¼ Where" & _
                 " ������� = 'K' And ���id Is Null And" & vbNewLine & _
                 "      ID In (Select ҽ��id" & vbNewLine & _
                 "             From ����ҽ��״̬" & vbNewLine & _
                 "             Where (��������, ����ʱ��, ������Ա) In (Select ��������, ����ʱ��, ������Ա" & vbNewLine & _
                 "                                          From ����ҽ��״̬" & vbNewLine & _
                 "                                          Where ҽ��id = [1] And ����ʱ�� = [2] And �������� = [3]))"
        End If

    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, dat����ʱ��, lng��������)
    If rsTmp.RecordCount = 1 Then GetRollAdviceIDs = rsTmp!ҽ��ids & ""
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckAdviceAddModi(Optional ByVal intType As Integer, Optional ByRef lngҽ��ID As Long, Optional ByRef datTurn As Date) As Boolean
'���ܣ��¿����޸�ʱ����Ƿ������޸Ļ�����
'������intType=0-������1-�޸�
    Dim lngBabyEdit As Long
    Dim blnReturn As Boolean
    
    If mlng����ID = 0 Then Exit Function
    If CheckDataMoved Then Exit Function
    '��鲡���Ƿ��������
    If Not CheckPatiIsAduit Then Exit Function
    With vsAdvice
        If intType = 1 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
            If lngҽ��ID = 0 Then Exit Function
        
            lngBabyEdit = CheckBabyEdit(Val(.TextMatrix(.Row, COL_Ӥ��ID)))
            If lngBabyEdit = 1 Then
                MsgBox "��ǰ���˲��ڱ����ң�������༭����ҽ����", vbInformation, gstrSysName
                Exit Function
            ElseIf lngBabyEdit = 2 Then
                MsgBox "��ǰ���˵�Ӥ�����ڱ����ң�������༭Ӥ��ҽ����", vbInformation, gstrSysName
                Exit Function
            End If
            
            'ҽ���´��ҽ��
            If Val(.TextMatrix(.Row, COL_ǰ��ID)) <> mlngǰ��ID Then
                MsgBox "�����޸ĸ�ҽ��,��ҽ���Ǹ���������ҽ�������ġ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            'ת�Ʋ���
            If CheckOtherDeptPatiOpt = False Then Exit Function
            
            '��У�Ի��ѷ�ֹ
            If InStr(",4,8,9,", .TextMatrix(.Row, COL_ҽ��״̬)) > 0 Then
                MsgBox "��ǰѡ���ҽ���Ѿ����ϻ�ֹͣ�������޸ġ�", vbInformation, gstrSysName
                Exit Function
            ElseIf InStr(",1,2,", .TextMatrix(.Row, COL_ҽ��״̬)) = 0 Then
                MsgBox "��ǰѡ���ҽ���Ѿ���У�ԣ������޸ġ�", vbInformation, gstrSysName
                Exit Function
            End If
            
            '��ǩ����ҽ�������޸�
            If Val(.TextMatrix(.Row, COL_ǩ����)) = 1 Then
                MsgBox "��ǰѡ���ҽ���Ѿ�ǩ���������޸ġ�����ȡ��ǩ����", vbInformation, gstrSysName
                Exit Function
            End If
            
            If mint���� = 1 Then '��ʿվ����
                '��ʿ�����Ѿ�����˵�ҽ�����������޸��޸�
                If .TextMatrix(.Row, COL_����ҽ��) Like "*/*" Then
                    MsgBox "��ǰѡ���ҽ���Ѿ���ҽ����ˣ������޸ġ�", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                '��ִҵ�ʸ��ҽ��ֻ��ɾ���޸�δ��˵�ҽ����
                If Not mblnHaveAuditPriv Then
                    If HaveAuditPriv(GetAuditName(.TextMatrix(.Row, COL_����ҽ��))) Then
                        MsgBox "��û���ʸ��޸ĵ�ǰѡ���ҽ�������ߵ�ǰѡ���ҽ���Ѿ�����ˣ������޸ġ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End If
        End If
    End With
    
    If mintPState = ps���ת�� Or mintPState = psԤ�� Or mintPState = ps��Ժ Then
        If CheckPatiTurnLimit(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, datTurn, mintPState) = False Then Exit Function
    End If
    CheckAdviceAddModi = True
End Function

Private Sub FuncApplyBlood(ByVal intType As Long)
'���ܣ���Ѫ���뵥
'������intType=0 ������=1�޸ģ�=2�鿴 ,=4 �˶�
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long
    Dim lngNo As Long
    Dim bln��Ѫ As Boolean
    Dim blnApply As Boolean
    
    If intType <> 2 Then
        If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
        If intType = 0 Then
            If Not FuncPathAdd() Then Exit Sub
        End If
        '����Ƿ������м�����רҵ����ְ��
        If gbln��Ѫ�����м����� Then
            If UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "����ҽʦ" And UserInfo.רҵ����ְ�� <> "������ҽʦ" Then
                MsgBox "��������Ѫ�ּ��������Ѫҽ��ֻ���м�������רҵ����ְ��ҽʦ�����´", vbInformation, "��Ѫ���뵥"
                Exit Sub
            End If
        End If
        '�޸�ʱ����Ƿ����
        If intType = 1 Then
            If Not CanEditBloodAdvice(Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)), Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��־)) = 1, Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��鷽��)) = 1) Then Exit Sub
        End If
    
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         lngRow = vsAdvice.Row
         bln��Ѫ = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_��鷽��)) = 1
    End If
    
    If gblnѪ��ϵͳ = True Then
        blnApply = frmApplyBloodNew.ShowMe(Me, mlng����ID, mlng��ҳID, IIF(mlng�������� = 1, 1, 0), intType, lngUpdateAdvice, mlng����ID, mlng����ID, Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2), mintPState, datTurn, mrsDefine, mclsMipModule, , , , , mbytӤ��, , mlngǰ��ID, IIF(bln��Ѫ = True, 1, 0))
    Else
        blnApply = frmApplyBlood.ShowMe(Me, mlng����ID, mlng��ҳID, IIF(mlng�������� = 1, 1, 0), intType, lngUpdateAdvice, mlng����ID, mlng����ID, Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2), mintPState, datTurn, mrsDefine, mclsMipModule, , , , , mbytӤ��, , mlngǰ��ID)
    End If
    
    If blnApply = True Then
    
        If mlng·��״̬ = 1 And Not gobjPath Is Nothing And (intType = 0 Or intType = 1) And lngUpdateAdvice <> 0 Then
            '��ȡ��Ѫ�������
            lngNo = Sys.RowValue("����ҽ����¼", lngUpdateAdvice, "�������", "ID")
            If lngNo <> 0 Then Call FuncPathSet(lngNo)
        End If
        'ˢ��ҽ��
        Call RefreshData
        'ѡ�����һ��ҽ��
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_ҽ������
    End If
End Sub

Private Sub FuncApplyOperation(ByVal intType As Long)
'���ܣ��������뵥
'������intType=0 ������=1�޸ģ�=2�鿴
    Dim lngUpdateAdvice As Long
    Dim datTurn As Date
    Dim lngRow As Long, strDefine As String
    Dim lng��������ID As Long
    Dim lngNo As Long
    
    If intType <> 2 Then
        If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
        '�޸�ʱ����Ƿ����
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)) = 2 Then
                MsgBox "���뵥�Ѿ���ˣ����������޸ġ�", vbInformation, "�������뵥"
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
    
    lng��������ID = Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2)
    If Not mrsDefine Is Nothing Then
        mrsDefine.Filter = "�������='F'"
        If Not mrsDefine.EOF Then strDefine = Trim(NVL(mrsDefine!ҽ������))
    End If

    If frmApplyOperation.ShowMe(Me, 0, intType, mlng����ID, mlng��ҳID, IIF(mlng�������� = 1, 1, 0), lngUpdateAdvice, mlng����ID, lng��������ID, strDefine, mlng����ID, mintPState, datTurn, 0, mclsMipModule, , , mlngǰ��ID, mbytӤ��) Then
        'ˢ��ҽ��
        Call RefreshData
        'ѡ�����һ��ҽ��
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_ҽ������
        
        If mlng·��״̬ = 1 And Not gobjPath Is Nothing And lngUpdateAdvice <> 0 Then
            lngNo = Sys.RowValue("����ҽ����¼", lngUpdateAdvice, "�������", "ID")
            If lngNo <> 0 Then Call FuncPathSet(lngNo)
        End If
    End If
End Sub

Private Sub FuncApplyConsultation(ByVal intType As Long)
'���ܣ��������뵥
'������intType=0 ������=1�޸ģ�=2�鿴
    Dim datTurn As Date
    Dim lngUpdateAdvice As Long
    Dim lngRow As Long, lngNo As Long
    Dim lng��������ID As Long

    If Not CheckWindow Then Exit Sub
    
    If intType <> 2 Then
        If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
    End If
    
    If intType <> 0 Then
         lngUpdateAdvice = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
         lngRow = vsAdvice.Row
    End If
    
    lng��������ID = Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2)
    Set mfrmEac = frmApplyConsultation
    If frmApplyConsultation.ShowMe(mfrmParent, lngUpdateAdvice, lngNo, intType, 0, mlng����ID, mlng��ҳID, mlng����ID, lng��������ID, mlng����ID, mintPState, datTurn, mclsMipModule, , , mlngǰ��ID, mbytӤ��) Then
        'ˢ��ҽ��
        Call RefreshData
        'ѡ�����һ��ҽ��
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_ҽ������
    End If
    
End Sub

Private Sub zlPASSMap()
'����:����Pass VsAdvie����ӳ��
'ע��:ɾ�����޸�������������ʱ�����������ҩ�����еĹ�������
    Dim blnTmp As Boolean
    
    If mobjPassMap Is Nothing Then
        Set mobjPassMap = DynamicCreate("zlPassInterface.clsPassMap", "������ҩ���", True)
    End If
    
    If gobjPass Is Nothing Then
        blnTmp = False
    Else
        blnTmp = gobjPass.PassType <> UNPASS
    End If
    
    mblnPass = blnTmp And Not mobjPassMap Is Nothing
    
    If mblnPass Then
        With mobjPassMap
            .lngModel = PM_סԺҽ���嵥
            .int���� = mint����
            Set .frmMain = Me
            Set .vsAdvice = vsAdvice
            Set .VSCOL = .GetVSCOL(COL_ID, COL_���ID, COL_�������, _
                COL_������ĿID, COL_�շ�ϸĿID, col_ҽ������, COL_��Ч, COL_����, COL_������λ, COL_�÷�, COL_����, , COL_����ʱ��, COL_����ҽ��, _
                COL_��ʼʱ��, COL_��������ID, COL_��ֹʱ��, COL_Ƶ��, , , , COL_��ʾ, COL_���, COL_ҽ��״̬, , , , , COL_ִ������, COL_�걾��λ, _
                , , , , , COL_����, , COL_ҽ������, COL_��ҩĿ��, COL_��������)
            Set .PassPati = .GetPatient()
            mblnPass = gobjPass.zlPassCheck(mobjPassMap)
        End With
    End If
End Sub

Private Sub zlPASSPati()
'����:���ò�����Ϣ
    
    With mobjPassMap.PassPati
        .lng����ID = mlng����ID
        .lng��ҳID = mlng��ҳID
    End With
End Sub

Public Sub LocatedAdviceRow(ByVal lngҽ��ID As Long)
'���ܣ���λҽ����
    Dim blnExist As Boolean
    Dim i As Long
    
    i = vsAdvice.FindRow(CStr(lngҽ��ID), , COL_ID)
    If i = -1 Then Exit Sub
    If vsAdvice.RowHidden(vsAdvice.Row) Then Exit Sub    '��λ���������еĴ���
    vsAdvice.Row = i
    Call vsAdvice.ShowCell(i, vsAdvice.FixedCols)
    Call ShowAdvicePlan(i, blnExist) '����ǰ�����ҽ����λ���������ҳǩ
    If blnExist Then
        For i = 0 To tbcAppend.ItemCount - 1
            If tbcAppend(i).Tag = "����" Then
                tbcAppend(i).Selected = True
                Exit For
            End If
        Next
    End If
End Sub

Private Sub SetCISMsg(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngҽ��ID As Long, ByVal lng���� As Long)
'���ܣ�����ҽ����ͣ��Ϣ����ҽ������ȷ��ֹͣʱ����
'������lng���� 1 ��ʾ����ҽ�� 0 �ǽ���ҽ��
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "select 1 From ҵ����Ϣ�嵥 A Where a.����id=[1] And a.����id=[2] And a.���ͱ��� ='ZLHIS_CIS_002' And a.���ȳ̶�=[3] And a.�Ƿ�����=0 And Rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID, IIF(lng���� = 1, 2, 1))
    If rsTmp.EOF Then
        strSQL = "Select a.�������� As ����,a.��Ժ����id As ����id, a.��ǰ����id As ����id From ������ҳ A Where a.����id =[1] And a.��ҳid =[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        strSQL = "Zl_ҵ����Ϣ�嵥_Insert(" & lng����ID & "," & lng��ҳID & "," & rsTmp!����ID & "," & rsTmp!����ID & "," & IIF(rsTmp!���� = 1, 1, 2) & ",'����ֹͣҽ����','0010','ZLHIS_CIS_002'," & _
            lngҽ��ID & "," & IIF(lng���� = 1, 2, 1) & ",0,null," & rsTmp!����ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Name)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetTagһ����ҩ(Optional ByVal lngRow As Long)
'���ܣ���һ����ҩ��ҽ��ǰ�ӱ�־
    Dim i As Long
    Dim lngBg As Long, lngEd As Long
    Dim j As Long
    Dim lngStart As Long, lngEnd As Long

    If mvarCond.����ģʽ = 3 Then Exit Sub

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
             If RowInһ����ҩ(i, lngBg, lngEd) Then
                For j = lngBg To lngEd
                    If j = lngBg Then
                        .TextMatrix(j, COL_��) = "��"
                    ElseIf j = lngEd Then
                        .TextMatrix(j, COL_��) = "��"
                    Else
                        .TextMatrix(j, COL_��) = "��"
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
'���ܣ��°滤ʿվ����ʱ���ڲ��������ĸ�
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
    '��������ť
    rsBar.Sort = "��� desc"
    Set objBar = cbsSub(2)
    lngTmp = -1
    With objBar.Controls
        If Not rsBar.EOF Then
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!������, lngTmp)
                    objControl.IconId = rsBar!ͼ��ID
                    objControl.Parameter = rsBar!������
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
        rsBar.Sort = "���"
        Set objPopup = objBar.Controls.Add(xtpControlButtonPopup, conMenu_Tool_PlugIn, "��չ����", , False)
            objPopup.BeginGroup = True
            objPopup.Style = xtpButtonIconAndCaption
        With objPopup.CommandBar.Controls
            For i = 1 To rsBar.RecordCount
                Set objControl = .Add(xtpControlButton, rsBar!����ID, rsBar!�˵���)
                objControl.IconId = rsBar!ͼ��ID
                objControl.Parameter = rsBar!������
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
'���ܣ���ʾָ���е�ԤԼ��Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lngIdx As Long
    Dim i As Long
    
    blnExist = False
    rtfSche.Text = "": rtfSche.SelStart = 0
    
    On Error GoTo errH
    
    If Val(vsAdvice.TextMatrix(lngRow, COL_RISԤԼID)) = 0 Then Exit Function
    
    strSQL = "select ����豸����,To_Char(ԤԼ����,'YYYY-MM-DD') as ԤԼ����," & vbNewLine & _
        "To_Char(ԤԼ��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') as ԤԼ��ʼʱ��," & vbNewLine & _
        "To_Char(ԤԼ����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ԤԼ����ʱ��," & vbNewLine & _
        "To_Char(ԤԼ��ʼʱ���,'YYYY-MM-DD HH24:MI:SS') as ԤԼ��ʼʱ���," & vbNewLine & _
        "To_Char(ԤԼ����ʱ���,'YYYY-MM-DD HH24:MI:SS') as ԤԼ����ʱ���,DECODE(�Ƿ����,1,'�Ѿ�ԤԼ����','�Ѿ�ԤԼ') as ԤԼ״̬" & vbNewLine & _
        "from RIS���ԤԼ Where ҽ��ID=[1]"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsAdvice.TextMatrix(lngRow, COL_ID)))
    If Not rsTmp.EOF Then
        With rtfSche
            For i = 0 To rsTmp.Fields.Count - 1
                .SelBold = False
                .SelText = IIF(.Text = "", "", vbCrLf) & rsTmp.Fields(i).Name & "��" & NVL(rsTmp.Fields(i).value)
                lngIdx = .Find(rsTmp.Fields(i).Name & "��", , , rtfNoHighlight Or rtfMatchCase)
                If lngIdx <> -1 Then
                    .SelStart = lngIdx
                    .SelLength = Len(rsTmp.Fields(i).Name & "��")
                    .SelBold = True
                    .SelIndent = 100
                End If
                .SelStart = Len(.Text)
            Next
            '��궨λ�ڵ�һ��
            lngIdx = .Find(rsTmp.Fields(0).Name & "��", 0, , rtfNoHighlight Or rtfMatchCase)
            If lngIdx <> -1 Then .SelStart = lngIdx + Len(rsTmp.Fields(0).Name & "��")
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
'���ܣ�RISҽ��ԤԼ
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    lngResult = -1
    If HaveRIS Then
        With vsAdvice
            If InStr(",1,3,8,", "," & .TextMatrix(.Row, COL_ҽ��״̬) & ",") > 0 Then
                lngResult = gobjRis.HISScheduling(2, Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_������ĿID)))
                If lngResult = 0 Then
                    '�ɹ�ԤԼ�����״̬
                    strSQL = "select min(ԤԼID) as ID from RIS���ԤԼ where ҽ��id=[1]"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                    .TextMatrix(.Row, COL_RISԤԼID) = rsTmp!ID & ""
                End If
            Else
                MsgBox "ҽ��״̬Ϊ�¿���У�ԡ��ѷ���ʱ������ԤԼ��", vbInformation, gstrSysName
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
'���ܣ�RISҽ��ȡ��ԤԼ
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngResult As Long
    Dim blnDo As Boolean
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_RISԤԼID)) <> 0 Then
            If Val(.TextMatrix(.Row, COL_ҽ��״̬)) = 8 Then
                strSQL = "Select Max(b.ִ��״̬) As ��� From ����ҽ����¼ A, ����ҽ������ B Where a.Id = b.ҽ��id And (a.Id =[1] Or a.���id=[1])"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(.Row, COL_ID)))
                If Not rsTmp.EOF Then
                    If Val(rsTmp!��� & "") = 0 Then
                        blnDo = True
                    Else
                        MsgBox "��ҽ���Ѿ���ִ�л��߲���ִ�в���ȡ��ԤԼ��", vbInformation, gstrSysName
                    End If
                End If
            Else
                blnDo = True
            End If
        End If
        If blnDo Then
            If HaveRIS Then
                lngResult = gobjRis.HISSchedulingEx(Val(.TextMatrix(.Row, COL_ID)), Val(.TextMatrix(.Row, COL_RISԤԼID)))
                If lngResult = 0 Then
                    '�ɹ���ȡ������״̬
                    .TextMatrix(.Row, COL_RISԤԼID) = ""
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
'���ܣ����ݵ�ǰ�е���������ҽ�������е�ͼ���ʶ
'˵����ע���ǵ������ã�����һ������
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_ID)) <> 0 Then
            If Val(.TextMatrix(lngRow, COL_����ID)) <> 0 Or _
                .TextMatrix(lngRow, COL_��鱨��ID) <> "" Or _
                Val(.TextMatrix(lngRow, COL_RIS����ID)) <> 0 Or _
                Val(.TextMatrix(lngRow, COL_LIS����ID)) <> 0 Then
                
                
                If Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 0 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("����").Picture
                ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 1 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("��������").Picture
                ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 2 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("���沿����").Picture
                End If
            Else
                If Val(.TextMatrix(lngRow, COL_RISԤԼID)) <> 0 Then
                    Set .Cell(flexcpPicture, lngRow, COL_F����) = frmIcons.imgFlag.ListImages("ԤԼ").Picture
                End If
            End If
        End If
    End With
End Sub

Private Sub FuncPathSet(ByVal lng������� As Long)
'����:����Ƿ�����·������Ŀ
'True-·������Ŀ;False-��·������Ŀ
    Dim byt����ʱ������ As Byte
    Dim lng·����ĿID As Long, lng�׶�Id As Long, lng���� As Long
    Dim i As Long, k As Long, lngִ��ID As Long
    Dim strSQL As String
    Dim str������ĿIDs As String, str·����Ŀ���� As String
    Dim strAdvices As String, str��ID As String, strList As String, strAdvicesOut As String
    Dim str���� As String
    Dim str��Ч As String
    Dim str��ʼ���� As String, strAddDate As String
    Dim dat���� As Date, DatAddDate As Date
    Dim rsTmp As ADODB.Recordset
    Dim rsAdvice As ADODB.Recordset
    Dim rsPath As ADODB.Recordset
    Dim rsStep As ADODB.Recordset
    Dim arrSQL As Variant
    Dim arrTmp As Variant
    Dim blnTrans As Boolean
    Dim blnPathOut As Boolean
    
    On Error GoTo errH:
    
    If Not (mint���� = 0 Or mint���� = 2) Then Exit Sub
    
    strSQL = "Select a.Id As ҽ��id,a.���ID, Nvl(a.���id, a.Id) As ��ID, a.������Ŀid,a.�������, b.��������, a.��ʼִ��ʱ��, a.ҽ����Ч" & vbNewLine & _
            "From ����ҽ����¼ A, ������ĿĿ¼ B" & vbNewLine & _
            "Where a.������Ŀid = b.Id And a.������� = [1]" & vbNewLine & _
            "Order By a.���"

    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�������)
    If rsAdvice.RecordCount = 0 Then Exit Sub

    For i = 1 To rsAdvice.RecordCount
    'ҩƷ������ҩ;�����÷����巨����Ѫ����;��,���鲻���ɼ���ʽ��������������������������鲻����λ����
        If i = 1 Then str��ID = rsAdvice!��ID & ""
        If str��ID <> rsAdvice!��ID & "" Then
            str��ID = rsAdvice!��ID & ""
            strAdvices = Mid(strAdvices, 2)
            str������ĿIDs = Mid(str������ĿIDs, 2)
            strList = strList & "&" & strAdvices & "|" & str������ĿIDs & "|" & str��Ч
            strAdvices = ""
            str������ĿIDs = ""
            str��Ч = ""
        End If
        strAdvices = strAdvices & "," & rsAdvice!ҽ��ID
        If Not (rsAdvice!������� & "" = "E" And InStr(",2,3,4,6,8,", "," & rsAdvice!�������� & ",") > 0) And Not (InStr(",G,F,D,", "," & rsAdvice!������� & ",") > 0 And NVL(rsAdvice!���ID, 0) <> 0) Then
            str������ĿIDs = str������ĿIDs & "," & rsAdvice!������ĿID
            If str��ʼ���� = "" Then str��ʼ���� = Format(rsAdvice!��ʼִ��ʱ��, "YYYY-MM-DD")
            If str��Ч = "" Then str��Ч = rsAdvice!ҽ����Ч
        End If
        rsAdvice.MoveNext
    Next
    strAdvices = Mid(strAdvices, 2)
    str������ĿIDs = Mid(str������ĿIDs, 2)
    If InStr(strList, strAdvices & "|" & str������ĿIDs & "|" & str��Ч) = 0 Then strList = strList & "&" & strAdvices & "|" & str������ĿIDs & "|" & str��Ч
    strList = Mid(strList, 2)
    arrTmp = Split(strList, "&")
    '��ȡ·����ǰ�׶�;��ǰ����
    Set rsPath = GetPatiPathInfo(mlng����ID, mlng��ҳID, str·����Ŀ����)
    If rsPath.RecordCount = 0 Then Exit Sub
    If rsPath!���� = CDate(str��ʼ����) Then
        strSQL = "Select ��ǰ�׶�id From �����ٴ�·�� Where ����ID = [1] And ��ҳID=[2]" & vbNewLine & _
                        "Union All" & vbNewLine & _
                        "Select ��ǰ�׶�id From ���˺ϲ�·�� Where ����ID = [1] And ��ҳID=[2]"
    Else
        '����ҽ���Ŀ�ʼ���ڻ�ȡ��Ӧ��·���׶�
        strSQL = "Select a.�׶�id as ��ǰ�׶�ID " & vbNewLine & _
                "From ����·��ִ�� A, �����ٴ�·�� B" & vbNewLine & _
                "Where b.Id = a.·����¼id And b.����id = [1] And b.��ҳid = [2] And a.���� = [3]" & vbNewLine & _
                "Union" & vbNewLine & _
                "Select a.�׶�id as ��ǰ�׶�ID  " & vbNewLine & _
                "From ����·��ִ�� A, ���˺ϲ�·�� B" & vbNewLine & _
                "Where b.Id = a.�ϲ�·����¼id And b.����id = [1] And b.��ҳid = [2] And a.���� = [3]"
    End If
    Set rsStep = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID, CDate(str��ʼ����))
    DatAddDate = zlDatabase.Currentdate
    strAddDate = "To_Date('" & Format(DatAddDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    arrSQL = Array()
    
    '������һ��������Ŷ�Ӧ����ҽ��
    For k = LBound(arrTmp) To UBound(arrTmp)
        strAdvices = Split(arrTmp(k), "|")(0)
        str������ĿIDs = Split(arrTmp(k), "|")(1)
        str��Ч = Split(arrTmp(k), "|")(2)
        If rsStep.RecordCount > 0 Then rsStep.MoveFirst    '��ҽ����ʼ���ڴ���·���ĵ�ǰ�׶εĵ�ǰ����ʱ,rsSetp���ؼ�¼��Ϊ0,Ĭ��Ϊ·������Ŀ
        Do While Not rsStep.EOF
            If rsStep!��ǰ�׶�ID & "" <> "" Then
                lng·����ĿID = CheckPathInItem(mlng����ID, mlng��ҳID, str������ĿIDs, str����, Val(rsStep!��ǰ�׶�ID & ""), False, CByte(str��Ч))
            End If
            If lng·����ĿID <> 0 Then Exit Do
            rsStep.MoveNext
        Loop
        '·������������
        If lng·����ĿID = 0 Then
            blnPathOut = True
            strAdvicesOut = strAdvicesOut & "," & strAdvices
        Else
            '·������Ŀ
            If rsPath!���� > CDate(str��ʼ����) Then
                Set rsTmp = GetPatiPathAppend(rsPath!·����¼id, CDate(str��ʼ����))
                If rsTmp.RecordCount > 0 Then
                    lng�׶�Id = rsTmp!�׶�id
                    lng���� = rsTmp!����
                    dat���� = CDate(str��ʼ����)
                End If
                byt����ʱ������ = 1 '��¼
            Else
                lng�׶�Id = rsPath!��ǰ�׶�ID
                lng���� = rsPath!��ǰ����
                dat���� = rsPath!����
                If rsPath!���� = CDate(str��ʼ����) Then
                    byt����ʱ������ = 0
                Else
                    byt����ʱ������ = 2 '�ݴ�
                End If
            End If
        
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����·������_Insert(0," & mlng����ID & "," & mlng��ҳID & ",NULL,0," & rsPath!·����¼id & "," & lng�׶�Id & _
                                     ",To_Date('" & Format(dat����, "yyyy-MM-dd") & "','YYYY-MM-DD')," & lng���� & _
                                     ",'" & str���� & "'," & lng·����ĿID & ",'" & strAdvices & "',Null,Null" & _
                                     ",'" & UserInfo.���� & "'," & strAddDate & ",NULL,1,Null,Null,Null,NUlL," & IIF(byt����ʱ������ = 1, 1, 0) & ",Null,Null,NUlL,Null,Null,Null,NUlL," & IIF(byt����ʱ������ = 0, "NULL", byt����ʱ������) & ")"
            
        End If
    Next
    
    If blnPathOut Then
        '·������Ŀ
        strAdvicesOut = Mid(strAdvicesOut, 2)
        If strAdvicesOut <> "" Then
            If rsPath!���� > CDate(str��ʼ����) Then
                Set rsTmp = GetPatiPathAppend(rsPath!·����¼id, CDate(str��ʼ����))
                If rsTmp.RecordCount > 0 Then
                    lng�׶�Id = rsTmp!�׶�id
                    lng���� = rsTmp!����
                    dat���� = CDate(str��ʼ����)
                End If
                byt����ʱ������ = 1 '��¼
            Else
                lng�׶�Id = rsPath!��ǰ�׶�ID
                lng���� = rsPath!��ǰ����
                dat���� = rsPath!����
                If rsPath!���� = CDate(str��ʼ����) Then
                    byt����ʱ������ = 0
                Else
                    byt����ʱ������ = 2 '�ݴ�
                End If
            End If
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_����·������_Insert(0," & mlng����ID & "," & mlng��ҳID & ",Null,0," & _
                                      rsPath!·����¼id & "," & lng�׶�Id & ",To_Date('" & Format(dat����, "yyyy-MM-dd") & "','YYYY-MM-DD')," & lng���� & _
                                      ",'" & str·����Ŀ���� & "',Null" & ",'" & strAdvicesOut & "',Null,Null,'" & UserInfo.���� & "'," & strAddDate & ",'·������Ŀ'" & _
                                      ",1,Null,Null,Null,NUlL," & IIF(byt����ʱ������ = 1, 1, 0) & ",Null,Null,NUlL,Null,Null,Null,NUlL," & IIF(byt����ʱ������ = 0, "NULL", byt����ʱ������) & ")"
        End If
    End If
    '�����ύ
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    '���·������Ŀһ������
    If blnPathOut Then
        lngִ��ID = GetPathOutItemID(Val(rsPath!·����¼id), DatAddDate)
        'ǿ��ˢ�¶�ȡ·��������Ϣ����Ϊ��������л�����
        Call gobjPath.zlRefresh(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, mintPState, False, True)
        Call gobjPath.zlExePathAppendItem(str·����Ŀ����, strAdvicesOut, lngִ��ID, dat����)
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
    Dim str��ǰ���� As String
    Dim i As Long
    Dim lng����ID As Long, lng���ID As Long
    Dim bln��ҽ As Boolean
    Dim blnDo As Boolean, blnIsCancel As Boolean
    Dim blnIsSend As Boolean, blnYes As Boolean
    Dim rsTmp As ADODB.Recordset, rsPath As ADODB.Recordset
    Dim objDiagEdit As zlMedRecPage.clsDiagEdit
    
    '��鲡���Ƿ��´��˳�Ժҽ��
    If mstrӤ�� = "" And mlng·��״̬ = 2 Then
        If CheckOutAdvice(mlng����ID, mlng��ҳID) Then
            MsgBox "�ò����Ѿ�����������·�����´��˳�Ժҽ�����������¿�ҽ����", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    
     '��ǰ���²��ˣ�δ�¹�ҽ���ģ���ǰ�����п��õ�·����ʱ�����һ�δ��д��Ժ��������ϵģ���ʾ����д��Ժ��ϡ�
    If mlng·��״̬ = -1 And mlng�������� <> 1 Then
        If InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") > 0 Then
            On Error GoTo errH
            strSQL = "select 1 From ����ҽ����¼ Where ����ID=[1] and ��ҳID=[2] and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
            If rsTmp.RecordCount = 0 Then
                If HavePath(mlng����ID) Then
                    strSQL = "select 1 From ������ϼ�¼ Where ������� In (1, 2, 11, 12) And ��¼��Դ = 3 And ����ID=[1] and ��ҳID=[2] and rownum<2"
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
                    If rsTmp.RecordCount = 0 Then
                        If MsgBox("�������п��õ��ٴ�·������Ϊ�˼�ʱ�����ٴ�·���������Ƿ���д��Ժ��ϣ�", vbQuestion + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                            If objDiagEdit Is Nothing Then
                                Set objDiagEdit = New zlMedRecPage.clsDiagEdit
                                Call objDiagEdit.InitDiagEdit(gcnOracle, glngSys, IIF(mlng�������� = 1, 1260, 1261), mclsMipModule)
                            End If
                            If objDiagEdit.ShowDiagEdit(Me, 0, mlng����ID, mlng��ҳID, IIF(mlng�������� = 1, 1, 2), mlng����ID, "", "", 0) Then
                                Set rsTmp = Get����ID(mlng����ID, mlng��ҳID, mlng����ID, bln��ҽ)
                                If bln��ҽ Then
                                    rsTmp.Filter = "������� =12 OR ������� = 2 "
                                    For i = 1 To rsTmp.RecordCount
                                        lng����ID = Val("" & rsTmp!����id)
                                        lng���ID = Val("" & rsTmp!���id)
                                        Set rsPath = GetPathTable(lng����ID, lng���ID, mlng����ID)
                                        If rsPath.RecordCount > 0 Then Exit For
                                        rsTmp.MoveNext
                                    Next
                                Else
                                    If rsTmp.RecordCount > 0 Then
                                        lng����ID = Val("" & rsTmp!����id)
                                        lng���ID = Val("" & rsTmp!���id)
                                    End If
                                    Set rsPath = GetPathTable(lng����ID, lng���ID, mlng����ID)
                                End If
                                If Not rsPath Is Nothing Then
                                    If rsPath.RecordCount > 0 Then
                                        Call gobjPath.zlRefresh(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, mintPState, mblnMoved, True)
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
    
    '·���еĲ��ˣ�����û������·����Ŀ�����ȵ�������
    If mlng·��״̬ = 1 And mvarCond.Ӥ�� <= 0 Then
        blnDo = True
        If mint���� = 2 Then
            blnDo = zlDatabase.GetPara("ҽ��ҽ����·������", glngSys, p�ٴ�·��Ӧ��, 0) = 0
        End If
        'δ����ʱ�������ҽ��������
        mblnNotEvaluete = Val(zlDatabase.GetPara("δ����ʱ�������ҽ��������", glngSys, p�ٴ�·��Ӧ��, 1)) = 1
        
        If blnDo Then
            If CheckPathNotEvaluete(mlng����ID, mlng��ҳID, blnIsSend, str��ǰ����) = False Then
                If gobjPath Is Nothing Then
                    MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ������¿�ҽ����", vbInformation, gstrSysName
                ElseIf InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") = 0 Then
                    MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ���û������·����Ȩ�ޣ������¿�ҽ����", vbInformation, gstrSysName
                Else
                    '֮ǰ����û�н���·��ҳ�棬��Ҫ��ͨ��ˢ�½ӿڶ�ȡ��ʼ����
                    Call gobjPath.zlRefresh(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, mintPState, mblnMoved, True)
                    Call gobjPath.zlExecPathSend(blnIsCancel)
                    Call LoadAdvice
                End If
                If Not blnIsCancel Then Exit Function
             Else
                If Not blnIsSend Then
                    If gobjPath Is Nothing Then
                        MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ������¿�ҽ����", vbInformation, gstrSysName
                        Exit Function
                    ElseIf InStr(GetInsidePrivs(p�ٴ�·��Ӧ��), ";����·��;") = 0 Then
                        MsgBox "�ò��˵��쵱ǰ�׶ε�·����Ŀδ���ɣ���û������·����Ȩ�ޣ������¿�ҽ����", vbInformation, gstrSysName
                        Exit Function
                    Else
                        '��������˲�����δ����ʱ�������ҽ�������죬����ʾ������ֱ�ӽ����������ɲ���
                        If mblnNotEvaluete Then
                            blnYes = MsgBox("��Ҫ���·������Ŀ��''" & str��ǰ���� & "'?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
                        End If
                        '���ѡ���������������ɲ�����ѡ�����������¿�·������Ŀ�� ��ǰ����
                        If blnYes = False Then
                            '֮ǰ����û�н���·��ҳ�棬��Ҫ��ͨ��ˢ�½ӿڶ�ȡ��ʼ����
                            Call gobjPath.zlRefresh(mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, mintPState, mblnMoved, True)
                            'û�����ɣ��򷵻�false��ֹ�¿�����
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
'���ܣ�ҽ������м�����
    Dim i As Long
    With vsAdvice
        .Rows = rsData.RecordCount + 1
        For i = 1 To rsData.RecordCount
            .TextMatrix(i, COL_ID) = NVL(rsData!ID)
            .TextMatrix(i, COL_���ID) = NVL(rsData!���ID)
            .TextMatrix(i, COL_���) = NVL(rsData!���)
            .TextMatrix(i, COL_Ӥ��ID) = NVL(rsData!Ӥ��ID)
            .TextMatrix(i, COL_ҽ��״̬) = NVL(rsData!ҽ��״̬)
            .TextMatrix(i, COL_�������) = NVL(rsData!�������)
            .TextMatrix(i, COL_��������) = NVL(rsData!��������)
            .TextMatrix(i, COL_�������) = NVL(rsData!�������)
            .TextMatrix(i, COL_��־) = NVL(rsData!��־)
            .TextMatrix(i, COL_��ʾ) = NVL(rsData!��ʾ)
            .TextMatrix(i, COL_��Ч) = NVL(rsData!��Ч)
            .TextMatrix(i, COL_��ʼʱ��) = NVL(rsData!��ʼʱ��)
            .TextMatrix(i, COL_��) = NVL(rsData!��)
            .TextMatrix(i, col_ҽ������) = NVL(rsData!ҽ������)
            .TextMatrix(i, col_����) = NVL(rsData!����)
            .TextMatrix(i, COL_Ƥ��) = NVL(rsData!Ƥ��)
            .TextMatrix(i, COL_����) = NVL(rsData!����)
            .TextMatrix(i, COL_����) = NVL(rsData!����)
            .TextMatrix(i, COL_����) = NVL(rsData!����)
            .TextMatrix(i, COL_Ƶ��) = NVL(rsData!Ƶ��)
            .TextMatrix(i, COL_�÷�) = NVL(rsData!�÷�)
            .TextMatrix(i, COL_ҽ������) = NVL(rsData!ҽ������)
            .TextMatrix(i, COL_ִ��ʱ��) = NVL(rsData!ִ��ʱ��)
            .TextMatrix(i, COL_��ֹʱ��) = NVL(rsData!��ֹʱ��)
            .TextMatrix(i, COL_ִ�п���) = NVL(rsData!ִ�п���)
            .TextMatrix(i, COL_ִ������) = NVL(rsData!ִ������)
            .TextMatrix(i, COL_�ϴ�ִ��) = NVL(rsData!�ϴ�ִ��)
            .TextMatrix(i, COL_״̬) = NVL(rsData!״̬)
            .TextMatrix(i, COL_����ҽ��) = NVL(rsData!����ҽ��)
            .TextMatrix(i, COL_����ʱ��) = NVL(rsData!����ʱ��)
            .TextMatrix(i, COL_У�Ի�ʿ) = NVL(rsData!У�Ի�ʿ)
            .TextMatrix(i, COL_У��ʱ��) = NVL(rsData!У��ʱ��)
            .TextMatrix(i, COL_ͣ��ҽ��) = NVL(rsData!ͣ��ҽ��)
            .TextMatrix(i, COL_ͣ��ʱ��) = NVL(rsData!ͣ��ʱ��)
            .TextMatrix(i, COL_ͣ����ʿ) = NVL(rsData!ͣ����ʿ)
            .TextMatrix(i, COL_ȷ��ͣ��ʱ��) = NVL(rsData!ȷ��ͣ��ʱ��)
            .TextMatrix(i, COL_����ҩ��) = NVL(rsData!����ҩ��)
            .TextMatrix(i, COL_����״̬) = NVL(rsData!����״̬)
            .TextMatrix(i, COL_�걾״̬) = NVL(rsData!�걾״̬)
            .TextMatrix(i, COL_������ĿID) = NVL(rsData!������ĿID)
            .TextMatrix(i, COL_�Թܱ���) = NVL(rsData!�Թܱ���)
            .TextMatrix(i, COL_ִ�б��) = NVL(rsData!ִ�б��)
            .TextMatrix(i, COL_���δ�ӡ) = NVL(rsData!���δ�ӡ)
            .TextMatrix(i, COL_ǰ��ID) = NVL(rsData!ǰ��ID)
            .TextMatrix(i, COL_ǩ����) = NVL(rsData!ǩ����)
            .TextMatrix(i, COL_�ļ�ID) = NVL(rsData!�ļ�ID)
            .TextMatrix(i, COL_������) = NVL(rsData!������)
            .TextMatrix(i, COL_����ID) = NVL(rsData!����ID)
            .TextMatrix(i, COL_�շ�ϸĿID) = NVL(rsData!�շ�ϸĿID)
            .TextMatrix(i, COL_������λ) = NVL(rsData!������λ)
            .TextMatrix(i, COL_��������ID) = NVL(rsData!��������id)
            .TextMatrix(i, COL_���״̬) = NVL(rsData!���״̬)
            .TextMatrix(i, COL_�������) = NVL(rsData!�������)
            .TextMatrix(i, COL_��˱��) = NVL(rsData!��˱��)
            .TextMatrix(i, COL_��ΣҩƷ) = NVL(rsData!��ΣҩƷ)
            .TextMatrix(i, COL_�걾��λ) = NVL(rsData!�걾��λ)
            .TextMatrix(i, COL_��ҩĿ��) = NVL(rsData!��ҩĿ��)
            .TextMatrix(i, COL_��鱨��ID) = NVL(rsData!��鱨��ID)
            .TextMatrix(i, COL_�������״̬) = NVL(rsData!�������״̬)
            .TextMatrix(i, COL_���������) = NVL(rsData!���������)
            .TextMatrix(i, COL_RISԤԼID) = NVL(rsData!RISԤԼID)
            .TextMatrix(i, COL_RIS����ID) = NVL(rsData!RIS����ID)
            .TextMatrix(i, COL_LIS����ID) = NVL(rsData!LIS����ID)
            .TextMatrix(i, COL_RISԤԼ״̬) = NVL(rsData!RISԤԼ״̬)
            .TextMatrix(i, col_������Ŀ����) = NVL(rsData!������Ŀ����)
            .TextMatrix(i, COL_��鷽��) = NVL(rsData!��鷽��)
            .TextMatrix(i, COL_Σ��ֵID) = NVL(rsData!Σ��ֵID)
            .TextMatrix(i, COL_�׵���) = Val(rsData!�Ƿ��������� & "")
            rsData.MoveNext
        Next
    End With
End Sub

Private Sub FuncAdviceRISPrintSch(ByVal lngFunID As Long)
'���ܣ�RISҽ��ԤԼ����ӡ
'������lngFunID ����ID�� lngFunID ��conMenu_Tool_RisPrint����ӡ����ԤԼ����lngFunID ��conMenu_Tool_RisPrintBat��������ӡ
    Dim lngResult As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strName As String
    
    On Error GoTo errH
    
    lngResult = -1
    If HaveRIS Then
        If lngFunID = conMenu_Tool_RisPrint Then
            With vsAdvice
                If Not .TextMatrix(.Row, COL_�������) = "D" Then
                    MsgBox "��ǰҽ������Ӱ������Ŀ��", vbInformation, gstrSysName
                    Exit Sub
                End If
                If .TextMatrix(.Row, COL_RISԤԼID) = 0 Then
                    MsgBox "��ǰӰ����ҽ��û�б�ԤԼ�����ܴ�ӡ��", vbInformation, gstrSysName
                    Exit Sub
                End If
                lngResult = gobjRis.HISPrintOneRisScheduleRpt(Val(.TextMatrix(.Row, COL_ID)))
            End With
        Else
            Call frmAdviceRisReport.ShowMe(Me, mlng����ID)
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
'���ܣ��ж�ѡ�е�ҽ���Ƿ�����Ѫҽ��
    Dim i As Long
    Dim blnTrue As Boolean
    
    With vsAppend
        For i = .FixedRows To .Rows - 1
           If .TextMatrix(i, COLSend("�������")) = "K" Then
                blnTrue = (.TextMatrix(i, COLSend("��Ѫ����")) = "1")
                Exit For
           End If
        Next
    End With
    IsUseBloodAdvice = blnTrue
End Function

Private Function HaveItemToRis(ByVal lng���ͺ� As Long, ByRef lngҽ��ID As Long) As Boolean
'���ܣ������ͺŹ��˱��η��͵�ҽ�����Ƿ��з���RISȥ��ҽ��û��
'˵������RIS��أ����������˷��Ͳ���ʱ���ã������η�������>=2��ҽ��ʱ���ֹ�����أ����뵥�����ˡ���ΪRIS�Ǳ�һ��ֻ�ܴ���һ��ҽ����
    Dim strSQL As String
    Dim strIDs As String, i As Long
    Dim strTmp As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    strSQL = "Select a.id,a.������Ŀid,0 as RISItem" & vbNewLine & _
        "From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C" & vbNewLine & _
        "Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And b.���ͺ� =[1] And" & vbNewLine & _
        "      (a.������� In ('F', 'D') Or a.������� = 'E' And Nvl(c.��������,'0') in ('5','0')) And a.���id Is Null And a.ҽ����Ч = 1"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng���ͺ�)
    Set rsTmp = zlDatabase.CopyNewRec(rsTmp)
    
    For i = 1 To rsTmp.RecordCount
        If InStr("," & strIDs & ",", "," & rsTmp!������ĿID & ",") = 0 Then
            strIDs = strIDs & "," & rsTmp!������ĿID
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
            If InStr("," & strTmp & ",", "," & rsTmp!������ĿID & ",") > 0 Then
                rsTmp!RISItem = 1
            End If
            rsTmp.MoveNext
        Next
        rsTmp.Filter = "RISItem=1"
        If rsTmp.RecordCount = 1 Then
            lngҽ��ID = rsTmp!ID
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
'���ܣ���ȡ���������ʾ��
    Dim strTmp As String
    With vsAdvice
        If Val(.TextMatrix(lngRow, COL_RIS����ID)) <> 0 Then
            strTmp = "(RIS����)"
        ElseIf Val(.TextMatrix(lngRow, COL_����ID)) <> 0 Then
            strTmp = "(HIS����)"
        ElseIf .TextMatrix(lngRow, COL_��鱨��ID) <> "" Then
            strTmp = "(רҵ��PACS����)"
        ElseIf Val(.TextMatrix(lngRow, COL_LIS����ID)) <> 0 Then
            strTmp = "(����LIS����)"
        Else
            If Val(.TextMatrix(lngRow, COL_RISԤԼID)) <> 0 Then
                If Val(.TextMatrix(lngRow, COL_RISԤԼ״̬)) = 0 Then
                    strTmp = "�Ѿ�ԤԼ"
                Else
                    strTmp = "�Ѿ�ԤԼ����"
                End If
            End If
        End If
        If strTmp <> "" And Val(.TextMatrix(lngRow, COL_RISԤԼID)) = 0 Then
            If Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 0 Then
                strTmp = "����δ��" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 1 Then
                strTmp = "��������" & strTmp
            ElseIf Val(.Cell(flexcpData, lngRow, COL_����״̬)) = 2 Then
                strTmp = "���沿������" & strTmp
            End If
        End If
    End With
    GetAdviceReportTip = strTmp
End Function

Private Sub FuncApplyCustom(ByVal intType As Long, ByVal lng�ļ�ID As Long)
'���ܣ��Զ������뵥
'������intType=0 ������=1�޸ģ�=2�鿴
    Dim lng������� As Long
    Dim datTurn As Date
    Dim lngRow As Long
    Dim lng��������ID As Long
    Dim lngNo As Long
    Dim objApplyCustom As New frmApplyCustom
    
    If intType <> 2 Then
        If mint���� <> 2 Then If CheckAdviceAddModi(intType, 0, datTurn) = False Then Exit Sub
        '�޸�ʱ����Ƿ����
        If intType = 1 Then
            If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_���״̬)) = 2 Then
                MsgBox "���뵥�Ѿ���ˣ����������޸ġ�", vbInformation, "���뵥"
                intType = 2
            End If
        End If
        If intType = 0 Then
            If Not FuncPathAdd() Then Exit Sub
        End If
    End If
    
    If intType <> 0 Then
         lng������� = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_�������))
         lngRow = vsAdvice.Row
    End If
    
    lng��������ID = Get��������ID(UserInfo.ID, mlng�������ID, mlng����ID, 2)
    If objApplyCustom.ShowMe(mfrmParent, 0, intType, mlng����ID, mlng��ҳID, IIF(mlng�������� = 1, 1, 0), lng�ļ�ID, lng�������, mlng����ID, lng��������ID, mlng����ID, mrsDefine, mintPState, datTurn, 0, mclsMipModule, mlngǰ��ID, mbytӤ��, mint����) Then
        If mlng·��״̬ = 1 And Not gobjPath Is Nothing And lng������� <> 0 Then
            lngNo = lng�������
            If lngNo <> 0 Then Call FuncPathSet(lngNo)
        End If
        'ˢ��ҽ��
        Call RefreshData
        'ѡ�����һ��ҽ��
        vsAdvice.Row = vsAdvice.Rows - 1
        vsAdvice.ShowCell vsAdvice.Rows - 1, col_ҽ������
    End If
End Sub

Private Sub FuncAdviceRISModi()
'���ܣ�����RISԤԼ
    Dim lngҽ��ID As Long
    Dim lngԤԼID As Long
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        lngԤԼID = Val(.TextMatrix(.Row, COL_RISԤԼID))
    End With
    
    strSQL = "select 1 from ����ҽ������ a where a.ҽ��id=[1] and nvl(a.ִ��״̬,0) in (0,3) and nvl(a.ִ�й���,0)<=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        If HaveRIS(False) Then
            Call gobjRis.HISReSchedule(lngҽ��ID, lngԤԼID)
        End If
    Else
        MsgBox "����Ŀ�Ѿ�ִ�У�����������������", vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function MakeBillCharge(ByVal lngҽ��ID As Long) As Long
'���ܣ����ҽ������ʱ�Զ�������������ҩƷ����
'�������Ƿ����������˷���
'���أ�0-�������˲�����1-��ֹ���˲���
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strҽ��IDs As String
    Dim i As Long
    Dim strNO As String
    Dim datCur As Date
    Dim blnTran As Boolean
    Dim arrSQL As Variant
    Dim strMsg As String
    
    On Error GoTo errH
    
    '��ȡ����ҽ��IDƴ��
    strSQL = "Select a.id,b.no,a.���ID,a.ҽ������ From ����ҽ����¼ A,����ҽ������ B Where a.id=b.ҽ��id and (a.Id = [1] Or a.���id = [1]) and b.��¼����=2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    For i = 1 To rsTmp.RecordCount
        strҽ��IDs = strҽ��IDs & "," & rsTmp!ID
        strNO = rsTmp!NO
        If IsNull(rsTmp!���ID) Then
            strMsg = rsTmp!ҽ������ & ""
        End If
        rsTmp.MoveNext
    Next
    strҽ��IDs = Mid(strҽ��IDs, 2)
    If strҽ��IDs = "" Then Exit Function
    
    '����Ƿ����δ��˵���������
    strSQL = "Select 1 From ����ҽ����¼ A, ����ҽ������ B, סԺ���ü�¼ C, ���˷������� D" & vbNewLine & _
        " Where (a.Id = [1] Or a.���id = [1]) And a.Id = b.ҽ��id And b.ҽ��id = c.ҽ����� And c.Id = d.����id And" & vbNewLine & _
        " c.��¼״̬ In (0, 1, 3) And d.״̬ = 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    If Not rsTmp.EOF Then
        strMsg = "ҽ��""" & strMsg & """��Ŀ����δ��˵��������룬��ȡ�����������������ٻ��ˡ�"
        MsgBox strMsg, vbInformation, gstrSysName
        MakeBillCharge = 1
        Exit Function
    End If
    
    '�ж��Ƿ��������
    strSQL = "Select Min(ID) As ����id, Max(a.���˲���id) As �������id,a.ҽ�����,a.�շ�ϸĿid, Sum(a.����) As ����" & vbNewLine & _
        "From סԺ���ü�¼ A" & vbNewLine & _
        "Where a.No = [1] And a.��¼���� = 2 And a.�շ���� In ('5', '6')" & vbNewLine & _
        "And instr(','||[2]||',', ','||a.ҽ�����||',')>0" & vbNewLine & _
        "and nvl(a.ִ��״̬,0)<>0" & vbNewLine & _
        "Group By a.�շ�ϸĿid,a.ҽ�����" & vbNewLine & _
        "having Sum(a.����)>0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, strҽ��IDs)
    
    If Not rsTmp.EOF Then
        strMsg = "ҽ��""" & strMsg & """��Ŀ������ҩƷ�ѷ�ҩ����ֹ���ˡ�" & vbCrLf & _
            "�Ƿ��Զ���������ҩƷ���ã���������������룬���������ǰ������"
            
        If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        
            arrSQL = Array()
            datCur = zlDatabase.Currentdate
            For i = 1 To rsTmp.RecordCount
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "zl_���˷�������_insert(" & rsTmp!����ID & "," & rsTmp!�շ�ϸĿID & "," & rsTmp!�������id & "," & rsTmp!���� & ",'" & UserInfo.���� & "'," & _
                    "To_Date('" & Format(datCur, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1,0,'���˼��ҽ���Զ�����')"
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
'���ܣ����ã�����/ȡ����Σֵҽ������
'������strPar-��ʽ��Σ��ֵID,ҽ��ID(��ҽ��ID)
'      blnCheck-true ȡ����ϵ��false ���ù�ϵ
    Dim lngΣ��ֵID As Long
    Dim lngҽ��ID As Long
    Dim lng���� As Long
    Dim strSQL As String
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim i As Long
    Dim lngOtherΣ��ֵID As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    lng���� = IIF(blnCheck, 2, 1)
    lngΣ��ֵID = Split(strPar, ",")(0)
    lngҽ��ID = Split(strPar, ",")(1)
    strSQL = "Zl_����Σ��ֵҽ��_Update(" & lng���� & "," & lngΣ��ֵID & "," & lngҽ��ID & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    If blnCheck Then
        'ͬһ��ҽ���ɹ������Σ��ֵ��ȡ��ʱҪ��һ���ж��Ƿ��й���
        strSQL = "select a.Σ��ֵID,a.ҽ��ID from ����Σ��ֵҽ�� a where a.ҽ��ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
        If Not rsTmp.EOF Then
            lngOtherΣ��ֵID = rsTmp!Σ��ֵID & ""
        End If
    End If
    
    
    If RowInһ����ҩ(vsAdvice.Row, lngBegin, lngEnd) Then
        For i = lngBegin To lngEnd
            Set vsAdvice.Cell(flexcpPicture, i, col_ҽ������) = Nothing
            Set vsAdvice.Cell(flexcpPicture, i, col_����) = Nothing
            If blnCheck Then
                vsAdvice.TextMatrix(i, COL_Σ��ֵID) = lngOtherΣ��ֵID
            Else
                vsAdvice.TextMatrix(i, COL_Σ��ֵID) = lngΣ��ֵID
            End If
            Call SetAdviceIcon(i)
        Next
    Else
        '���½�����ͼ��
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_ҽ������) = Nothing
        Set vsAdvice.Cell(flexcpPicture, vsAdvice.Row, col_����) = Nothing
        If blnCheck Then
            vsAdvice.TextMatrix(vsAdvice.Row, COL_Σ��ֵID) = lngOtherΣ��ֵID
        Else
            vsAdvice.TextMatrix(vsAdvice.Row, COL_Σ��ֵID) = lngΣ��ֵID
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

Private Function GetCriticalAdvice(ByRef lngҽ��ID As Long) As ADODB.Recordset
'���ܣ����ݵ�ǰѡ���е�ҽ����ѯ����֮������Σ��ֵ��¼
'���������� lngҽ��ID ����ǰ������ѡ��ҽ������ҽ��ID
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    With vsAdvice
        If Val(.TextMatrix(.Row, COL_���ID)) <> 0 Then
            lngҽ��ID = Val(.TextMatrix(.Row, COL_���ID))
        Else
            lngҽ��ID = Val(.TextMatrix(.Row, COL_ID))
        End If
    End With
    
    strSQL = "select a.Σ��ֵID,a.ҽ��ID from ����Σ��ֵҽ�� a where a.ҽ��ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
    
    Set GetCriticalAdvice = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Getȷ�ϻ���(lngҽ��ID As Long) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "SELECT A.����ʱ�� FROM ����ҽ������ A where ҽ��ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(lngҽ��ID))
    If Not rsTmp.EOF Then
        Getȷ�ϻ��� = IIF(rsTmp!����ʱ�� & "" = "", False, True)
    Else
        Getȷ�ϻ��� = False
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Executeȷ�ϻ���(blnCancel As Boolean) As Boolean
'���ܣ��Ƿ�ȷ��ҽ���μӻ���
    Dim strSQL As String, lng����ID As Long
    
    If mlng����ID = 0 Then Exit Function
    If MsgBox("�Ƿ�" & IIF(blnCancel, "ȡ��ȷ��", "ȷ��") & "ҽ���μ��˵Ļ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    If Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) <> 0 Then
        strSQL = "Zl_����ҽ������_���ﴦ��(" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & ",0," & IIF(blnCancel = True, "3", "2") & ",'" & UserInfo.���� & "')"
        On Error GoTo errH
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Call vsAdvice_AfterRowColChange(-1, -1, vsAdvice.Row, vsAdvice.Col) 'Ҫ����ִ��״̬'����Ҫ����ִ��״̬
    End If
    Executeȷ�ϻ��� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetCriticalData()
'���ܣ���ȡΣ��ֵ��¼
    Dim strSQL As String
    On Error GoTo errH
    If mblnΣ��ֵ Then
        strSQL = "select a.id,a.Σ��ֵ���� from ����Σ��ֵ��¼ a where a.����ID=[1] and a.��ҳID=[2] order by a.����ʱ�� desc"
        Set mrsΣ��ֵ = zlDatabase.OpenSQLRecord(strSQL, "zlRefresh", mlng����ID, mlng��ҳID)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub FuncViewLisRpt()
'���ܣ�������鱨��
'˵����������ģʽ�����жϱ��ξ����Ƿ���PDF����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    If mblnMoved Then
        strSQL = "select 1 from H����ҽ����¼ a,H����ҽ������ b,Hҽ���������� c where a.id=b.ҽ��id and b.����id=c.id and c.����  in (0,2) and a.����id=[1] and a.��ҳid=[2]"
    Else
        strSQL = "select 1 from ����ҽ����¼ a,����ҽ������ b,ҽ���������� c where a.id=b.ҽ��id and b.����id=c.id and c.����  in (0,2) and a.����id=[1] and a.��ҳid=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    If Not rsTmp.EOF Then
        '����ҳǩ��ʾ
        Call frmLisALL.ShowMe(mfrmParent, mlng����ID, mlng��ҳID, mlng����ID, mlng����ID, pסԺҽ���´�, mMainPrivs)
    Else
        '��ǰ����ģʽ
        Call InitObjLis(pסԺҽ��վ)
        If Not gobjLIS Is Nothing Then
            gobjLIS.PatientSampleBrowse mfrmParent, mlng����ID, mMainPrivs, mlng����ID, mlng����ID, 2, mlng��ҳID
        Else
            frmLisView.ShowMe mlng����ID, pסԺҽ���´�, mfrmParent
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
'���ܣ�������д�ܾ�������ɴ��ڵ��ú�����ҩ�����ӿ�
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strAdviceIDs As String
    Dim strErr As String
    
    On Error GoTo errH
    
    strSQL = "select 1 from ����ҽ����¼ a where a.����id=[1] and a.��ҳid=[2] and a.ҽ��״̬=1 and a.������� in ('5','6') and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        '���¿���ҩƷҽ��
        Call gobjPass.ZLPharmReviewResultIn(mfrmParent, mlng����ID, mlng��ҳID, strErr)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function GetFormOperation() As String
'���ܣ���ȡ�������ѡ�񣬸ýӿڻ��ڴ���ж��ǰ���ã��°滤ʿվ �������񴰿�
'���أ���¼��ǰ�����пؼ�ѡ��״̬
'˵������������ǰҽ����ѡ��XML�ṹ��ʽ
    

    'Private Type FilterCond
    '    Ӥ�� As Integer
    '    ���� As Boolean
    '    ���� As Boolean
    '    δ���� As Boolean
    '    ���� As Integer     '0-ȫ����1����飬2�����飬3������
    '    δ������ As Boolean
    '    �ѳ����� As Boolean
    '    ��ʾģʽ As Integer '0-��࣬1����ϸ
    '    ҽ����ʾ As Integer '0-����ҽ����1������ҽ��
    '    ����ģʽ As Integer '0-����������1��������2��������3������
    '    ��ʼʱ�� As Date
    '    ����ʱ�� As Date
    '    �Ǳ���ҽ�� As Boolean
    '    �Ǳ���ҽ�� As Boolean
    '    δ����ֹʱ�� As Boolean '�Ƿ���ʾδ��(ִ����ֹʱ��)��ҽ��
    'End Type
    'mvarCond
    'cboTime.ListIndex
    'lngҽ��ID = Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID))
    
    Dim strXML As String

    strXML = "<root>"
    strXML = strXML & "<ye>" & mvarCond.Ӥ�� & "</ye>"     'Ӥ��
    strXML = strXML & "<cz>" & IIF(mvarCond.����, 1, 0) & "</cz>" '������Boolean 0/1
    strXML = strXML & "<kn>" & IIF(mvarCond.����, 1, 0) & "</kn>"  '���ڣ�Boolean 0/1
    strXML = strXML & "<wjz>" & IIF(mvarCond.δ����, 1, 0) & "</wjz>"  'δ���ʣ� Boolean 0/1
    strXML = strXML & "<bg>" & mvarCond.���� & "</bg>"  '����
    strXML = strXML & "<wcbg>" & IIF(mvarCond.δ������, 1, 0) & "</wcbg>"  'δ������ Boolean
    strXML = strXML & "<ycbg>" & IIF(mvarCond.�ѳ�����, 1, 0) & "</ycbg>"  '�ѳ����� Boolean
    strXML = strXML & "<xsms>" & mvarCond.��ʾģʽ & "</xsms>"  '��ʾģʽ
    strXML = strXML & "<yzxs>" & mvarCond.ҽ����ʾ & "</yzxs>"  'ҽ����ʾ
    strXML = strXML & "<glms>" & mvarCond.����ģʽ & "</glms>"  '����ģʽ
    strXML = strXML & "<kssj>" & Format(mvarCond.��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "</kssj>"   '    ��ʼʱ�� As Date
    strXML = strXML & "<jssj>" & Format(mvarCond.����ʱ��, "yyyy-MM-dd HH:mm:ss") & "</jssj>"   '     ����ʱ�� As Date
    strXML = strXML & "<sbgyz>" & IIF(mvarCond.�Ǳ���ҽ��, 1, 0) & "</sbgyz>" '    �Ǳ���ҽ�� As Boolean
    strXML = strXML & "<fbgyz>" & IIF(mvarCond.�Ǳ���ҽ��, 1, 0) & "</fbgyz>" '    �Ǳ���ҽ�� As Boolean
    strXML = strXML & "<wdzzsj>" & IIF(mvarCond.δ����ֹʱ��, 1, 0) & "</wdzzsj>" '    δ����ֹʱ�� As Boolean '�Ƿ���ʾδ��(ִ����ֹʱ��)��ҽ��
    strXML = strXML & "<cbotime>" & cboTime.ListIndex & "</cbotime>" 'ʱ�䷶Χ����������ֵ
    strXML = strXML & "<yzid>" & Val(vsAdvice.TextMatrix(vsAdvice.Row, COL_ID)) & "</yzid>" '����ѡ���ҽ��ID
    strXML = strXML & "</root>"
    

    GetFormOperation = strXML
End Function

Public Function RestoreFormOperation(ByVal strValue As String)
'���ܣ��ָ��������ѡ��
'������strValue ǰ�����пؼ�ѡ��״̬
'Public Sub LocatedAdviceRow(ByVal lngҽ��ID As Long)
    Dim objXML As New zl9ComLib.clsXML
    Dim strTmp As String
    
    On Error Resume Next
    
    Call objXML.OpenXMLDocument(strValue)
    
    Call objXML.GetSingleNodeValue("ye", strTmp) 'Ӥ��
    mvarCond.Ӥ�� = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("cz", strTmp) '����
    mvarCond.���� = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("kn", strTmp) '����
    mvarCond.���� = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("wjz", strTmp) 'δ����
    mvarCond.δ���� = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("bg", strTmp) '����
    mvarCond.���� = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("wcbg", strTmp) 'δ������
    mvarCond.δ������ = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("ycbg", strTmp) '�ѳ�����
    mvarCond.�ѳ����� = 1 = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("xsms", strTmp) '��ʾģʽ
    mvarCond.��ʾģʽ = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("yzxs", strTmp) 'ҽ����ʾ
    mvarCond.ҽ����ʾ = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("glms", strTmp) '����ģʽ
    mvarCond.����ģʽ = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("kssj", strTmp) '��ʼʱ��
    mvarCond.��ʼʱ�� = CDate(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("jssj", strTmp) '����ʱ��
    mvarCond.����ʱ�� = CDate(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("sbgyz", strTmp) '�Ǳ���ҽ��
    mvarCond.�Ǳ���ҽ�� = 1 = Val(strTmp): strTmp = ""
    
    
    Call objXML.GetSingleNodeValue("fbgyz", strTmp) '�Ǳ���ҽ��
    mvarCond.�Ǳ���ҽ�� = 1 = Val(strTmp): strTmp = ""
    
    
    Call objXML.GetSingleNodeValue("wdzzsj", strTmp) 'δ����ֹʱ��
    mvarCond.δ����ֹʱ�� = 1 = Val(strTmp): strTmp = ""
        
    
    Call objXML.GetSingleNodeValue("cbotime", strTmp) 'ʱ�䷶Χ����������ֵ
    cboTime.ListIndex = Val(strTmp): strTmp = ""
    
    Call objXML.GetSingleNodeValue("yzid", strTmp) '������ʱ��
    mvarCond.ҽ��ID = Val(strTmp): strTmp = ""
    
End Function

Private Sub Set�걾״̬()
'���ܣ��Լ���ҽ�����ñ걾״̬�У������LIS�����з���
    Dim i As Long, strҽ��IDs As String, strMsg As String
    Dim rsAdvice As ADODB.Recordset
    Dim strIDAndRow As String, strTmp As String
    Dim lngRow As Long
    
    On Error GoTo errH
    
    If mvarCond.����ģʽ <> 3 Then Exit Sub
    Call InitObjLis(pסԺҽ��վ)
    If gobjLIS Is Nothing Then Exit Sub
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, COL_�������) = "E" And .TextMatrix(i, COL_��������) = "6" And Val(.TextMatrix(i, COL_���ID)) = 0 And Val(.TextMatrix(i, COL_ҽ��״̬)) = 8 Then
                strҽ��IDs = strҽ��IDs & "," & Val(.TextMatrix(i, COL_ID))
                strIDAndRow = strIDAndRow & "," & Val(.TextMatrix(i, COL_ID)) & ";" & i & "<Tab>"
            End If
        Next
        If strҽ��IDs <> "" Then
            Set rsAdvice = gobjLIS.GetSampleType(Mid(strҽ��IDs, 2), strMsg)
            If strMsg <> "" Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
            If Not rsAdvice Is Nothing Then
                rsAdvice.Filter = 0
                For i = 1 To rsAdvice.RecordCount
                    If InStr(strIDAndRow, "," & rsAdvice!ҽ��ID & ";") > 0 Then
                        strTmp = Split(strIDAndRow, "," & rsAdvice!ҽ��ID & ";")(1)
                        lngRow = Val(Split(strTmp, "<Tab>")(0))
                        .TextMatrix(lngRow, COL_�걾״̬) = rsAdvice!ҽ��״̬ & ""
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
'���ܣ�������鱨��
'˵����δ�����Ķ����
    Dim blnAutoRead As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngҽ��ID As Long
    
    On Error GoTo errH
    Call CreateObjectPacs(mobjPublicPACS)
    If Not mobjPublicPACS Is Nothing Then
        
        strSQL = "select max(b.id) as ҽ��ID  from ����ҽ������ a,����ҽ����¼ b " & _
                " Where a.��鱨��ID Is Not Null And a.ҽ��ID = b.ID And b.����id=[1] and b.��ҳid=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
        lngҽ��ID = Val(rsTmp!ҽ��ID & "")
        
        Call mobjPublicPACS.zlDocShowReport(lngҽ��ID, , blnAutoRead, mfrmParent)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
