VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmQCHistory 
   Caption         =   "��ʷ�ʿز�ѯ"
   ClientHeight    =   8565
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   Icon            =   "frmQCHistory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11400
   StartUpPosition =   3  '����ȱʡ
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
   Begin VB.ComboBox cbo���� 
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
   Begin VB.ComboBox cbo���� 
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
      Begin VB.CommandButton cmdˢ�� 
         Height          =   600
         Left            =   2085
         Picture         =   "frmQCHistory.frx":0EBE
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   90
         Width           =   330
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Index           =   0
         Left            =   435
         TabIndex        =   5
         Top             =   75
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   64356355
         CurrentDate     =   39110
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Index           =   1
         Left            =   435
         TabIndex        =   6
         Top             =   390
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd��"
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
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   420
         Width           =   180
      End
      Begin VB.Label lbl���� 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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

Private Enum mColG  '�ʿ�Ʒ����
    ID = 0: ѡ��: ����: �ʿ�Ʒ: ��ֵ: SD: ˮƽ: ��ʼ����: ��������: �����ֵ: ����SD: ����cv
End Enum
Private Enum mColL  '�ʿ����ݱ���
    ��� = 0: ID: ����: ����
End Enum
Private Enum mColR  '�ʿر������
    ID = 0: ������Ŀid: ���: ����: �걾��: ��Ŀ: ���: �ʿ�Ʒ: ˮƽ
End Enum
Private Enum mTab   '�ʿ�ͼ����:����Ϊͳ������(Ƶ��ͼ)��Levey_Jenningsͼ��Z_����ͼ��Youdenͼ���ۻ���ͼ��Monicaͼ��Grubbs���Grubbsͼ��
    LJ = 0: FQ: ZS: YD: MN: CS: Grubbs: GS
End Enum

Private Enum mColC  '�ʿ�Ʒ����
    ID = 0: Ԥ���ֵ: Ԥ��SD: Ԥ��CV: ���¾�ֵ: ����sd: ����CV: �ۼƾ�ֵ: �ۼ�sd: �ۼ�CV
End Enum


Const conPane_Record = 201
Const conPane_Charts = 202
Const conPane_Report = 203
Const conPane_Data = 204
Const conPane_Calc = 205
'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mstrPrivs As String     '��ǰʹ����Ȩ�޴�
Private mlngListWidth As Long   '�б������ƿ��
Private mblnCusum As Boolean    '��ǰ�����Ƿ�Ӧ���ۻ��͹��򣬾����Ƿ��ṩ�ۻ���ͼ��

Private mfrmRptTxt As frmQCTodayReport  '���������Ӵ���
Private mfrmChartFQ As frmQCChartFQ '����ͳ�ƴ���
Private mfrmChartLJ As frmQCChartLJ     'LJ����ͼ����
Private mfrmChartZS As frmQCChartZS     'Z-����ͼ����
Private mfrmChartYD As frmQCChartYD     'Youdenͼ����
Private mfrmChartCS As frmQCChartCS     '�ۻ���ͼ����
Private mfrmChartMN As frmQCChartMN     'Monicaͼ����

Private mfrmGrubbs As frmQCGrubbs      'Grubbs ���
Private mfrmChartGS As frmQCChartGS    'Grubbs ͼ��
'-----------------------------------------------------
'��ʱ����
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
Private mblnShowAll As Boolean              '��ʾ����ʧ�ر���
Private mstr�ڼ�  As String                 '���ڼ��������
Private mEditMode As Integer                '�༭ģʽ 0=�Ǳ༭ 1=���ڱ༭
Private mstrPigeonhole As String            '�鵵��

Private mLastStartDate As Date, mLastEndDate As Date
Private mLastCell As String '�����뿪ǰ�ĵ�Ԫ��������������Ź���
Private mint��ʾʧЧ��¼ As Integer '0-����ʾ��1-��ʾ
Private mintLJͼ��λ��ʾ     As Integer '0-����λ, 1-��λ (Ĭ��)
Private Const ID_MENU_MOUSE = 90                                    '�Ҽ��˵�
Private mlngItemID As Long                                          '��ǰѡ�е���ĿID
Private mLastItemID As Long                                         '�ϴ���ʾ����ĿID�������ظ�ˢ��
'-----------------------------------------------------
'����Ϊ�ڲ���������
'-----------------------------------------------------
Private Function zlRefRecord() As Long
    '���ܣ�ˢ���ʿؽ����¼
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim date��ʼ As Date, date���� As Date
    
    Err = 0: On Error GoTo ErrHand
    If mlngItemID = 0 Then Exit Function
    '|| '-' || Decode(Nvl(R.���ý��, 0), 0, 999, R.���ý��)
    gstrSql = "Select R.id,Q.����ʱ�� As ����,Q.ʱ��, To_Char(Q.���Դ���, '000')  As ����," & vbNewLine & _
            "       Q.�ʿ�Ʒid, Zl_lis_ToNumber(Q.�ʿ�Ʒid,R.������Ŀid,R.������,R.id) As ���," & vbNewLine & _
            "       Nvl(T.���, 0) As ���, Q.������,R.���ý��" & vbNewLine & _
            "From �����ʿؼ�¼ Q, ������ͨ��� R, �����ʿر��� T" & vbNewLine & _
            "Where Q.�걾id = R.����걾id And R.ID = T.���id(+) /* And Nvl(R.�Ƿ����, 0) = 1*/ And Q.����id + 0 = [1] And" & vbNewLine & _
            "      R.������Ŀid + 0 = [2] And" & vbNewLine & _
            IIf(mint��ʾʧЧ��¼ = 1, "", "Nvl(R.���ý��, 0)=0 And ") & _
            "      (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By Q.����ʱ��,  Q.���Դ���, Q.�ʿ�Ʒid"
            'Nvl(���ý��, 0) * -1 +
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, _
                CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)), mlngItemID, _
                Format(Me.dtp����(0).Value, "yyyy-MM-dd"), Format(Me.dtp����(1).Value, "yyyy-MM-dd"))
    With Me.vfgRecord
        .Redraw = flexRDNone
        .Clear
        .FixedCols = 3
        .Cols = .FixedCols
        .ExtendLastCol = False '���Զ���չ���һ�еĿ��
        .Rows = 6 + Me.vfgReagent.Rows - 1
        
        .ColWidth(0) = 1200
        .TextMatrix(mColL.ID, 0) = "": .RowHidden(mColL.ID) = True
        .TextMatrix(mColL.���, 0) = ""
        .TextMatrix(mColL.���, 1) = "��ֵ": .ColWidth(1) = 500
        .TextMatrix(mColL.���, 2) = "SD": .ColWidth(2) = 500
        
        .TextMatrix(mColL.����, 0) = "����" & vbNewLine & "ʱ��"  ': .ColWidth(mColL.����) = 1050
        .RowHeight(mColL.����) = 600
        
        .TextMatrix(mColL.����, 0) = "����" ': .ColWidth(mColL.����) = 600 ': .ColHidden(mColL.����) = True
        .TextMatrix(.Rows - 2, 0) = "ʵ������": .RowHidden(.Rows - 2) = True
        .TextMatrix(.Rows - 1, 0) = "������": .RowHidden(.Rows - 1) = True '.ColWidth(.Cols - 1) = 800
        For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
            strTemp = Split(Me.vfgReagent.TextMatrix(lngCount, mColG.����), ", ")(0)
            .TextMatrix(mColL.ID, 0) = .TextMatrix(mColL.ID, 0) & "|" & strTemp & "=" & Me.vfgReagent.TextMatrix(lngCount, mColG.ID)
            .TextMatrix(lngCount + mColL.����, 0) = strTemp
            .TextMatrix(lngCount + mColL.����, 1) = Me.vfgReagent.TextMatrix(lngCount, mColG.��ֵ)
            .TextMatrix(lngCount + mColL.����, 2) = Me.vfgReagent.TextMatrix(lngCount, mColG.SD)
            
            If Me.vfgReagent.Cell(flexcpChecked, lngCount, mColG.ѡ��) = flexChecked Then
                '.ColWidth(lngCount + mColL.����) = 900
                .RowHidden(lngCount + mColL.����) = False
            Else
                '.ColWidth(lngCount + mColL.����) = 0
                .RowHidden(lngCount + mColL.����) = True
            End If
        Next
        .ColAlignment(0) = flexAlignLeftCenter
'        For lngCount = 0 To .Rows - 1
'            .FixedAlignment(lngCount) = flexAlignCenterCenter
'        Next
        Do While Not rsTemp.EOF
            lngRow = 0
            For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
                If rsTemp!�ʿ�Ʒid = Val(Me.vfgReagent.TextMatrix(lngCount, mColG.ID)) Then
                    lngRow = lngCount + mColL.����
                    date��ʼ = CDate(Me.vfgReagent.TextMatrix(lngCount, mColG.��ʼ����))
                    date���� = CDate(Me.vfgReagent.TextMatrix(lngCount, mColG.��������))
                    Exit For
                End If
            Next
            If lngRow <> 0 Then
                lngCol = 0
                '���ʿ�Ʒ�����ڷ�Χ ��ʾ����
                If CDate(Format(rsTemp!����, "yyyy-MM-dd")) >= date��ʼ And _
                   CDate(Format(rsTemp!����, "yyyy-MM-dd")) <= date���� Then
                    For lngCount = .FixedCols To .Cols - 1
                        If .TextMatrix(.Rows - 2, lngCount) = Format(rsTemp!����, "yyyy-MM-dd") And _
                            .TextMatrix(mColL.����, lngCount) = "" & rsTemp!���� Then
                            lngCol = lngCount: Exit For
                        End If
                    Next
                    If lngCol = 0 Then
                        .Cols = .Cols + 1
                        lngCol = .Cols - 1
                        .ColWidth(lngCol) = 500
                        
                        .TextMatrix(mColL.���, lngCol) = .Cols - .FixedCols
                        .TextMatrix(mColL.����, lngCol) = Format(rsTemp!����, "yy-MM-dd") & vbNewLine & Trim("" & rsTemp!ʱ��)
                        
                        .TextMatrix(mColL.����, lngCol) = "" & rsTemp!����
                        .TextMatrix(.Rows - 2, lngCol) = Format(rsTemp!����, "yyyy-MM-dd")
                        .TextMatrix(.Rows - 1, lngCol) = "" & rsTemp!������
                    Else
                        If InStr(1, .TextMatrix(.Rows - 1, lngCol), rsTemp!������) = 0 Then
                            .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 1, lngCol) & "," & rsTemp!������
                        End If
                    End If
                    .TextMatrix(mColL.ID, lngCol) = .TextMatrix(mColL.ID, lngCol) & "|" & Val("" & rsTemp!�ʿ�Ʒid) & "=" & Val("" & rsTemp!ID)
                    .TextMatrix(lngRow, lngCol) = Trim("" & rsTemp!���)
                    If Left(.TextMatrix(lngRow, lngCol), 1) = "." Then .TextMatrix(lngRow, lngCol) = "0" & .TextMatrix(lngRow, lngCol)
                    
                    Select Case Val("" & rsTemp!���)
                    Case 1
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HC0FFFF
                        .Cell(flexcpFontBold, lngRow, lngCol) = True
                    Case 2
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HC0C0FF
                        .Cell(flexcpFontBold, lngRow, lngCol) = True
                    End Select
                    
                    '�������ý�����Ϊ��ɫ
                    If Val("" & rsTemp!���ý��) = 1 Then
                        .Cell(flexcpBackColor, lngRow, lngCol) = &HC0C0C0
                        .Cell(flexcpFontBold, lngRow, lngCol) = True
                    End If
                End If '-- ���ʿ�Ʒָ�����ڷ�Χ������ʾ
            End If
            rsTemp.MoveNext
        Loop
        If .Cols > .FixedCols Then
            .Cell(flexcpAlignment, mColL.���, .FixedCols, mColL.����, .Cols - 1) = flexAlignCenterCenter
            .AutoSize 0, .Cols - 1
        End If
        .Redraw = flexRDDirect
        If .Cols > .FixedCols Then .Col = .FixedCols: .Row = mColL.���� + 1
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
    '���ܣ�������ʾ���ԣ�ˢ�³��ʿؼ�¼��ͼ�κͱ���
    Dim strLists As String, intLists As Integer
    Dim lngItemID As Long, strFromDate As String, strToDate As String
    Dim strѡ�����ʿ�Ʒ As String '���û�ͼ����Ҫ�õ��ʿ�Ʒ���ڷ�Χ����
    
    If mlngItemID = 0 Then Exit Sub
    
    If mlngItemID = mLastItemID Then Exit Sub
    If mlngItemID = -1 Then mlngItemID = mLastItemID
    mLastItemID = mlngItemID
    With Me.vfgReagent
        strLists = "": intLists = 0: strѡ�����ʿ�Ʒ = ""
        For lngCount = 0 To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mColG.ѡ��) = flexChecked Then
                strLists = strLists & "," & .TextMatrix(lngCount, mColG.ID)
                strѡ�����ʿ�Ʒ = strѡ�����ʿ�Ʒ & ";" & .TextMatrix(lngCount, mColG.ID) & "=" & Format(CDate("" & .TextMatrix(lngCount, mColG.��ʼ����)), "yyyy-MM-dd") & "," & Format(CDate("" & .TextMatrix(lngCount, mColG.��������)), "yyyy-MM-dd")
                intLists = intLists + 1
            End If
        Next
        If strLists <> "" Then strLists = Mid(strLists, 2)
    End With
    lngItemID = mlngItemID
    strFromDate = Format(Me.dtp����(0).Value, "yyyy-MM-dd")
    strToDate = Format(Me.dtp����(1).Value, "yyyy-MM-dd")
    If strѡ�����ʿ�Ʒ <> "" Then strѡ�����ʿ�Ʒ = Mid(strѡ�����ʿ�Ʒ, 2)
    'ˢ���ʿر���
    If Me.dkpMan.FindPane(conPane_Report).Closed = False Then
        Call zlRefReport(strLists, lngItemID, strFromDate, strToDate)
    End If
        
    '��õ�ǰѡ��Ŀ���ͼ�������ʿ�Ʒ�仯����������ʾ�Ŀ���ͼ�Σ���ˢ������
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
    If Me.tbcCharts.Item(mTab.FQ).Selected Then Call mfrmChartFQ.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ)
    If Me.tbcCharts.Item(mTab.LJ).Selected Then Call mfrmChartLJ.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ, mintLJͼ��λ��ʾ)
    If Me.tbcCharts.Item(mTab.ZS).Selected Then Call mfrmChartZS.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ)
    If Me.tbcCharts.Item(mTab.YD).Selected Then Call mfrmChartYD.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ)
    If Me.tbcCharts.Item(mTab.CS).Selected Then Call mfrmChartCS.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ)
    If Me.tbcCharts.Item(mTab.MN).Selected Then Call mfrmChartMN.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ)
    If Me.tbcCharts.Item(mTab.Grubbs).Selected Then Call mfrmGrubbs.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ)
    If Me.tbcCharts.Item(mTab.GS).Selected Then Call mfrmChartGS.zlRefresh(strLists, lngItemID, strFromDate, strToDate, strѡ�����ʿ�Ʒ)
    
End Sub

Private Sub zlShowQCReport()
    '���ܣ�����δʧ�ص��ʿ�����
    Dim strLists As String, intLists As Integer
    Dim lngItemID As Long, strFromDate As String, strToDate As String
    Dim strѡ�����ʿ�Ʒ As String '���û�ͼ����Ҫ�õ��ʿ�Ʒ���ڷ�Χ����
    
    If mlngItemID = 0 Then Exit Sub
    
'    If mlngItemID = mLastItemID Then Exit Sub
    If mlngItemID = -1 Then mlngItemID = mLastItemID
    mLastItemID = mlngItemID
    With Me.vfgReagent
        strLists = "": intLists = 0: strѡ�����ʿ�Ʒ = ""
        For lngCount = 0 To .Rows - 1
            If .Cell(flexcpChecked, lngCount, mColG.ѡ��) = flexChecked Then
                strLists = strLists & "," & .TextMatrix(lngCount, mColG.ID)
                strѡ�����ʿ�Ʒ = strѡ�����ʿ�Ʒ & ";" & .TextMatrix(lngCount, mColG.ID) & "=" & Format(CDate("" & .TextMatrix(lngCount, mColG.��ʼ����)), "yyyy-MM-dd") & "," & Format(CDate("" & .TextMatrix(lngCount, mColG.��������)), "yyyy-MM-dd")
                intLists = intLists + 1
            End If
        Next
        If strLists <> "" Then strLists = Mid(strLists, 2)
    End With
    lngItemID = mlngItemID
    strFromDate = Format(Me.dtp����(0).Value, "yyyy-MM-dd")
    strToDate = Format(Me.dtp����(1).Value, "yyyy-MM-dd")
    If strѡ�����ʿ�Ʒ <> "" Then strѡ�����ʿ�Ʒ = Mid(strѡ�����ʿ�Ʒ, 2)

    Call frmQCReport.ShowME(strLists, lngItemID, strFromDate, strToDate, Me)
        
   
End Sub


Public Sub zlRefReport(strResList As String, lngItemID, strFromDate As String, strToDate As String)
    '���ܣ�ˢ���ʿر���
    '������ strResList  ��ǰѡ����ʿ�Ʒid�����Զ��ŷָ�
    '       lngItemId   ��ǰ��Ŀid
    '       strFromDate ��ʼ����
    '       strToDate   ��������
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long
    
    Err = 0: On Error GoTo ErrHand
    '��ȡʧ�ر���
    gstrSql = "Select R.ID,R.������Ŀid, Nvl(T.���, 0) As ���, Q.����ʱ�� As ����, Q.�걾��� As �걾��,D.������ ||'/'||Ӣ���� as ��Ŀ, Zl_lis_ToNumber(Q.�ʿ�Ʒid,R.������Ŀid,R.������,R.id) As ���," & vbNewLine & _
            "       M.���� || ', ' || M.���� As �ʿ�Ʒ, M.ˮƽ, Q.������" & vbNewLine & _
            "From �����ʿؼ�¼ Q, �����ʿ�Ʒ M, ������ͨ��� R, �����ʿر��� T,����������Ŀ D" & vbNewLine & _
            "Where Q.�ʿ�Ʒid = M.ID And Q.�걾id = R.����걾id And R.ID = T.���id And Nvl(R.���ý��,0)=0 And /*Nvl(R.�Ƿ����, 0) = 1 And*/ " & vbNewLine & _
            "      Instr(',' || [1] || ',', ',' || Q.�ʿ�Ʒid || ',') > 0 And R.������Ŀid + 0 = D.ID" & IIf(mblnShowAll, "", " And R.������Ŀid + 0 = [2]") & " And" & vbNewLine & _
            "      (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
            "Order By Q.����ʱ��, R.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strResList, lngItemID, strFromDate, strToDate)
    With Me.vfgReport
        .Redraw = flexRDNone
        
        .Clear
        
        If mblnShowAll Then
            .ToolTipText = "˫���б��е���Ŀ������ʾ��Ŀ���ʿ����ݡ�"
        Else
            .ToolTipText = "����[�鿴]�˵���ѡ����ʾ����ʧ�ر��桱"
        End If
        Set .DataSource = rsTemp
        Call .AutoSize(mColR.���, .Cols - 1)
        .ColWidth(mColR.ID) = 0: .ColHidden(mColR.ID) = True
        .ColWidth(mColR.������Ŀid) = 0: .ColHidden(mColR.ID) = True
        .ColWidth(mColR.���) = 280: .TextMatrix(0, mColR.���) = ""
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        For lngCount = .FixedRows To .Rows - 1
            Select Case .TextMatrix(lngCount, mColR.���)
                Case 1: Set .Cell(flexcpPicture, lngCount, mColR.���) = Me.imgList.ListImages(1).Picture
                Case 2: Set .Cell(flexcpPicture, lngCount, mColR.���) = Me.imgList.ListImages(2).Picture
            End Select
            .TextMatrix(lngCount, mColR.���) = ""
            If Left(.TextMatrix(lngCount, mColR.���), 1) = "." Then .TextMatrix(lngCount, mColR.���) = "0" & .TextMatrix(lngCount, mColR.���)
        Next
        .Redraw = flexRDDirect
        
        
        gstrSql = "select ������, �鵵�� from �����ʿر��� where ���id = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)))
        If rsTemp.EOF = False Then
            mstrPigeonhole = Trim(Nvl(rsTemp("�鵵��")))
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
    '���ܣ�ˢ���ʿؽ����¼
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lngCol As Long, lngRow As Long
    Dim date��ʼ As Date, date���� As Date
    
    
    Dim intFixeWidth As Integer
    
    Err = 0: On Error GoTo ErrHand
    If mlngItemID = 0 Then Exit Sub
    
    
    
    With Me.vfgCalc
        intFixeWidth = 1200
        .Redraw = flexRDNone
        .Cols = 10
        .Rows = 3
        .FixedRows = 2
        .ExtendLastCol = False '���Զ���չ���һ�еĿ��
        .MergeCells = flexMergeFree
        
        .ColHidden(mColC.ID) = True
        
        .TextMatrix(0, mColC.Ԥ���ֵ) = "Ԥ��": .ColWidth(mColC.Ԥ���ֵ) = intFixeWidth: .ColAlignment(mColC.Ԥ���ֵ) = flexAlignCenterCenter
        .TextMatrix(0, mColC.Ԥ��SD) = "Ԥ��": .ColWidth(mColC.Ԥ��SD) = intFixeWidth: .ColAlignment(mColC.Ԥ��SD) = flexAlignCenterCenter
        .TextMatrix(0, mColC.Ԥ��CV) = "Ԥ��": .ColWidth(mColC.Ԥ��CV) = intFixeWidth: .ColAlignment(mColC.Ԥ��CV) = flexAlignCenterCenter
        
        .TextMatrix(0, mColC.���¾�ֵ) = "����": .ColWidth(mColC.���¾�ֵ) = intFixeWidth: .ColAlignment(mColC.���¾�ֵ) = flexAlignCenterCenter
        .TextMatrix(0, mColC.����sd) = "����": .ColWidth(mColC.����sd) = intFixeWidth: .ColAlignment(mColC.����sd) = flexAlignCenterCenter
        .TextMatrix(0, mColC.����CV) = "����": .ColWidth(mColC.����CV) = intFixeWidth: .ColAlignment(mColC.����CV) = flexAlignCenterCenter
        
        .TextMatrix(0, mColC.�ۼƾ�ֵ) = "�������ۼ�": .ColWidth(mColC.�ۼƾ�ֵ) = intFixeWidth: .ColAlignment(mColC.�ۼƾ�ֵ) = flexAlignCenterCenter
        .TextMatrix(0, mColC.�ۼ�sd) = "�������ۼ�": .ColWidth(mColC.�ۼ�sd) = intFixeWidth: .ColAlignment(mColC.�ۼ�sd) = flexAlignCenterCenter
        .TextMatrix(0, mColC.�ۼ�CV) = "�������ۼ�": .ColWidth(mColC.�ۼ�CV) = intFixeWidth: .ColAlignment(mColC.�ۼ�CV) = flexAlignCenterCenter
        
        
        .TextMatrix(1, mColC.Ԥ���ֵ) = "��ֵ": .ColAlignment(mColC.Ԥ���ֵ) = flexAlignCenterCenter
        .TextMatrix(1, mColC.Ԥ��SD) = "SD": .ColAlignment(mColC.Ԥ��SD) = flexAlignCenterCenter
        .TextMatrix(1, mColC.Ԥ��CV) = "CV": .ColAlignment(mColC.Ԥ��CV) = flexAlignCenterCenter
        
        .TextMatrix(1, mColC.���¾�ֵ) = "��ֵ": .ColAlignment(mColC.���¾�ֵ) = flexAlignCenterCenter
        .TextMatrix(1, mColC.����sd) = "SD": .ColAlignment(mColC.����sd) = flexAlignCenterCenter
        .TextMatrix(1, mColC.����CV) = "CV": .ColAlignment(mColC.����CV) = flexAlignCenterCenter
        
        .TextMatrix(1, mColC.�ۼƾ�ֵ) = "��ֵ": .ColAlignment(mColC.�ۼƾ�ֵ) = flexAlignCenterCenter
        .TextMatrix(1, mColC.�ۼ�sd) = "SD": .ColAlignment(mColC.�ۼ�sd) = flexAlignCenterCenter
        .TextMatrix(1, mColC.�ۼ�CV) = "CV": .ColAlignment(mColC.�ۼ�CV) = flexAlignCenterCenter
        
        .Rows = Me.vfgReagent.Rows + 1
        For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
            .TextMatrix(lngCount + 1, mColC.Ԥ���ֵ) = Format(Me.vfgReagent.TextMatrix(lngCount, mColG.�����ֵ), "##0.00##")
            .TextMatrix(lngCount + 1, mColC.Ԥ��SD) = Format(Me.vfgReagent.TextMatrix(lngCount, mColG.����SD), "##0.00##")
            .TextMatrix(lngCount + 1, mColC.Ԥ��CV) = Format(Round(Val(Me.vfgReagent.TextMatrix(lngCount, mColG.����cv)) * 100, 4), "##0.00##")
            
            gstrSql = "Select Round(Avg(���), 4) As ��ֵ, Round(Stddev(���), 4) As Sd, Count(*) As ����" & vbNewLine & _
                "From (Select Trunc(Q.����ʱ��) As ����," & vbNewLine & _
                "              Avg(zl_Lis_toNumber(Q.�ʿ�ƷID,R.������ĿID,R.������,R.ID)) As ���" & vbNewLine & _
                "       From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿر��� T" & vbNewLine & _
                "       Where Q.�걾id = R.����걾id And Q.�ʿ�Ʒid = [1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
                "             Nvl(R.���ý��,0)=0 And R.ID=T.���ID(+) And Q.����ʱ�� Between   [3] and [4]  And Nvl(T.���, 0) <> 2" & vbNewLine & _
                "       Group By Trunc(Q.����ʱ��))"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReagent.TextMatrix(lngCount, mColG.ID)), mlngItemID, _
                            CDate(CStr(Format(Me.dtp����(0), "yyyy-MM-dd"))), CDate(CStr(Format(Me.dtp����(1), "yyyy-MM-dd"))))
                            
            .TextMatrix(lngCount + 1, mColC.���¾�ֵ) = IIf(Val(rsTemp("��ֵ") & "") = 0, "", Format(Val(rsTemp("��ֵ") & ""), "##0.00##"))
            .TextMatrix(lngCount + 1, mColC.����sd) = IIf(Val(rsTemp("sd") & "") = 0, "", Format(Val(rsTemp("sd") & ""), "##0.00##"))
            If Val(rsTemp("Sd") & "") <> 0 And Val(rsTemp("��ֵ") & "") <> 0 Then
                .TextMatrix(lngCount + 1, mColC.����CV) = Format(Round(Val(rsTemp("Sd") & "") / Val(rsTemp("��ֵ") & "") * 100, 2), "##0.00##")
            Else
                .TextMatrix(lngCount + 1, mColC.����CV) = ""
            End If
            
            gstrSql = "Select Round(Avg(���), 4) As ��ֵ, Round(Stddev(���), 4) As Sd, Count(*) As ����" & vbNewLine & _
                "From (Select Trunc(Q.����ʱ��) As ����," & vbNewLine & _
                "              Avg(zl_Lis_toNumber(Q.�ʿ�ƷID,R.������ĿID,R.������,R.ID)) As ���" & vbNewLine & _
                "       From �����ʿؼ�¼ Q, ������ͨ��� R,�����ʿر��� T" & vbNewLine & _
                "       Where Q.�걾id = R.����걾id And Q.�ʿ�Ʒid = [1] And R.������Ŀid + 0 = [2] And" & vbNewLine & _
                "             Nvl(R.���ý��,0)=0 And R.ID=T.���ID(+) And Q.����ʱ�� < [3] And Nvl(T.���, 0) <> 2" & vbNewLine & _
                "       Group By Trunc(Q.����ʱ��))"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReagent.TextMatrix(lngCount, mColG.ID)), mlngItemID, _
                            CDate(CStr(Format(Me.dtp����(0), "yyyy-MM-dd"))), CDate(CStr(Format(Me.dtp����(1), "yyyy-MM-dd"))))
                            
            .TextMatrix(lngCount + 1, mColC.�ۼƾ�ֵ) = IIf(Val(rsTemp("��ֵ") & "") = 0, "", Format(Val(rsTemp("��ֵ") & ""), "##0.00##"))
            .TextMatrix(lngCount + 1, mColC.�ۼ�sd) = IIf(Val(rsTemp("sd") & "") = 0, "", Format(Val(rsTemp("sd") & ""), "##0.00##"))
            If Val(rsTemp("Sd") & "") <> 0 And Val(rsTemp("��ֵ") & "") <> 0 Then
                .TextMatrix(lngCount + 1, mColC.�ۼ�CV) = Format(Round(Val(rsTemp("Sd") & "") / Val(rsTemp("��ֵ") & "") * 100, 2), "##0.00##")
            Else
                .TextMatrix(lngCount + 1, mColC.�ۼ�CV) = ""
            End If
        Next
        .MergeRow(0) = True
        .Redraw = flexRDDirect
    End With
    
    
'    '|| '-' || Decode(Nvl(R.���ý��, 0), 0, 999, R.���ý��)
'    gstrSql = "Select R.id,Q.����ʱ�� As ����,Q.ʱ��, To_Char(Q.���Դ���, '000')  As ����," & vbNewLine & _
'            "       Q.�ʿ�Ʒid, Zl_lis_ToNumber(Q.�ʿ�Ʒid,R.������Ŀid,R.������,R.id) As ���," & vbNewLine & _
'            "       Nvl(T.���, 0) As ���, Q.������,R.���ý��" & vbNewLine & _
'            "From �����ʿؼ�¼ Q, ������ͨ��� R, �����ʿر��� T" & vbNewLine & _
'            "Where Q.�걾id = R.����걾id And R.ID = T.���id(+) /* And Nvl(R.�Ƿ����, 0) = 1*/ And Q.����id + 0 = [1] And" & vbNewLine & _
'            "      R.������Ŀid + 0 = [2] And" & vbNewLine & _
'            IIf(mint��ʾʧЧ��¼ = 1, "", "Nvl(R.���ý��, 0)=0 And ") & _
'            "      (Q.����ʱ�� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd'))" & vbNewLine & _
'            "Order By Q.����ʱ��,  Q.���Դ���, Q.�ʿ�Ʒid"
'            'Nvl(���ý��, 0) * -1 +
'    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, _
'                CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)), mlngItemID, _
'                Format(Me.dtp����(0).Value, "yyyy-MM-dd"), Format(Me.dtp����(1).Value, "yyyy-MM-dd"))
'    With Me.vfgCalc
'        .Redraw = flexRDNone
'        .Clear
'        .FixedCols = 3
'        .Cols = .FixedCols
'        .ExtendLastCol = False '���Զ���չ���һ�еĿ��
'        .Rows = 6 + Me.vfgReagent.Rows - 1
'
'        .ColWidth(0) = 1200
'        .TextMatrix(mColL.ID, 0) = "": .RowHidden(mColL.ID) = True
'        .TextMatrix(mColL.���, 0) = ""
'        .TextMatrix(mColL.���, 1) = "��ֵ": .ColWidth(1) = 700
'        .TextMatrix(mColL.���, 2) = "SD": .ColWidth(2) = 700
'
'        .TextMatrix(mColL.����, 0) = "����" & vbNewLine & "ʱ��"  ': .ColWidth(mColL.����) = 1050
'        .RowHeight(mColL.����) = 600
'
'        .TextMatrix(mColL.����, 0) = "����" ': .ColWidth(mColL.����) = 600 ': .ColHidden(mColL.����) = True
'        .TextMatrix(.Rows - 2, 0) = "ʵ������": .RowHidden(.Rows - 2) = True
'        .TextMatrix(.Rows - 1, 0) = "������": .RowHidden(.Rows - 1) = True '.ColWidth(.Cols - 1) = 800
'        For lngCount = Me.vfgReagent.FixedRows To Me.vfgReagent.Rows - 1
'            strTemp = Split(Me.vfgReagent.TextMatrix(lngCount, mColG.����), ", ")(0)
'            .TextMatrix(mColL.ID, 0) = .TextMatrix(mColL.ID, 0) & "|" & strTemp & "=" & Me.vfgReagent.TextMatrix(lngCount, mColG.ID)
'            .TextMatrix(lngCount + mColL.����, 0) = strTemp
'            .TextMatrix(lngCount + mColL.����, 1) = Me.vfgReagent.TextMatrix(lngCount, mColG.��ֵ)
'            .TextMatrix(lngCount + mColL.����, 2) = Me.vfgReagent.TextMatrix(lngCount, mColG.SD)
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
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Me.vfgRecord.Cols <= Me.vfgRecord.FixedCols Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    Set objPrint.Body = Me.vfgRecord
    objPrint.Title.Text = Mid(Me.cbo����.Text, InStr(1, Me.cbo����.Text, ",") + 1) & "�ʿؽ���嵥"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cbo����_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim lngMachineID As Long                '����ID
    
    On Error GoTo errH
    
    lngMachineID = Val(zlDatabase.GetPara("����", glngSys, 1209, 0))
    
    If Me.cbo����.ListCount <= 0 Then Exit Sub
    
    If InStr(1, mstrPrivs, "���п���") > 0 Then
        gstrSql = "Select Distinct D.ID, D.����, D.����, D.�ʿ�ˮƽ��" & vbNewLine & _
                "From �������� D, �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ Q" & vbNewLine & _
                "Where D.ID = M.����id And M.ID = Q.�ʿ�Ʒid And Nvl(D.΢����, 0) <> 1 and d.ʹ��С��id = [1] " & vbNewLine
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)))
    Else
'        gstrSql = "Select Distinct D.ID, D.����, D.����, D.�ʿ�ˮƽ��" & vbNewLine & _
                "From �������� D, �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ Q" & vbNewLine & _
                "Where D.ID = M.����id And M.ID = Q.�ʿ�Ʒid And Nvl(D.΢����, 0) <> 1 And" & vbNewLine & _
                "      D.ʹ��С��id In (Select ����id From ������Ա Where ��Աid = [1]) and d.ʹ��С��id = [2] "
        gstrSql = "Select Distinct D.ID, D.����, D.����, D.�ʿ�ˮƽ��" & vbNewLine & _
                " From �������� D, �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ Q" & vbNewLine & _
                " Where D.ID = M.����id And M.ID = Q.�ʿ�Ʒid And Nvl(D.΢����, 0) <> 1 And D.ʹ��С��id = [2] And" & vbNewLine & _
                "      D.ID In (Select Distinct D.ID" & vbNewLine & _
                "               From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
                "               Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [1] And C.����id = D.ID)"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(UserInfo.ID), CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)))
    End If
    
    With rsTemp
        Me.cbo����.Clear
        
        Do While Not .EOF
            Me.cbo����.AddItem !���� & Space(200) & !�ʿ�ˮƽ��
            Me.cbo����.ItemData(Me.cbo����.NewIndex) = !ID
            If !ID = lngMachineID Then
                Me.cbo����.ListIndex = Me.cbo����.NewIndex
            End If
            .MoveNext
        Loop
'        If Me.cbo����.ListCount = 0 Then MsgBox "��δ���������ص��ʿ����ã�", vbInformation, gstrSysName: Unload Me: Exit Sub
        If Me.cbo����.ListCount > 0 And cbo����.ListIndex = -1 Then
            Me.cbo����.ListIndex = 0
'            If Me.cbo����.ListCount = 1 Then Me.cbo����.Enabled = False
        End If
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'-----------------------------------------------------
'����Ϊ�ؼ��¼�����
'-----------------------------------------------------

Private Sub cbo����_Click()
    Dim lngItemID As Long   '��ĿID

    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = Val(zlDatabase.GetPara("��Ŀ", glngSys, 1209, 0))
    
    mblnCusum = False
    If Me.cbo����.ListIndex = -1 Then Exit Sub
    Me.cbo����.Tag = Right(Me.cbo����.Text, 1)
    
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select Nvl(Count(*), 0) As �ۻ���" & vbNewLine & _
            "From ������������ A, �����ʿع��� R" & vbNewLine & _
            "Where A.����id = R.ID And A.���� = '1' And R.���� = 3 And A.����id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)))
    If rsTemp.RecordCount > 0 Then
        If rsTemp.Fields(0).Value > 0 Then mblnCusum = True
    End If
    
    gstrSql = "Select Distinct I.ID, I.����, I.Ӣ����, I.������" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ Q, ����������Ŀ I" & vbNewLine & _
            "Where M.ID = Q.�ʿ�Ʒid And Q.��Ŀid = I.ID And M.����id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)))
    
    If rsTemp.RecordCount <= 0 Then MsgBox "��δ��������ʿ�Ʒ���ã�", vbInformation, gstrSysName: Unload Me: Exit Sub
    
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
            cbrControl.Caption = "���Ϊ"
            Me.cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_Untread                                   'ȡ��
        mfrmRptTxt.zlEditCancel
        mEditMode = 0
        Set cbrControl = cbsThis.FindControl(, conMenu_Edit_Save, True, True)
        cbrControl.Caption = "���Ϊ"
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
    Case conMenu_Edit_Adjust                                        '��дʧ�ر���
        Set panThis = Me.dkpMan.FindPane(conPane_Report)
        panThis.Select
        Call mfrmRptTxt.ZlEditStart(Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)))
        mEditMode = 1
        Set cbrControl = cbsThis.FindControl(, conMenu_Edit_Save, True, True)
        cbrControl.Caption = "���汨��"
        Me.cbsThis.RecalcLayout
    Case conMenu_Edit_QCReport                                      '�ʿر���
        zlShowQCReport
    Case conMenu_Edit_Archive                                       '�鵵
        gstrSql = "select �鵵�� from �����ʿر��� where ���id = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)))
        If rsTmp.EOF = False Then
            If Nvl(rsTmp("�鵵��")) = "" Then
                If MsgBox("���Ҫ����ǰʧ�ر���鵵��", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                gstrSql = "Zl_�����ʿر���_Archive(" & Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)) & ",0)"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
                mstrPigeonhole = gstrDBUser
            Else
                If MsgBox("��ʧ�ر����Ѿ��鵵�����ȡ���鵵��", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbNo Then Exit Sub
                gstrSql = "Zl_�����ʿر���_Archive(" & Val(Me.vfgReport.TextMatrix(Me.vfgReport.Row, mColR.ID)) & ",1)"
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
        mlngItemID = -1 'ǿ��ˢ��
        Call zlRefOthers
    Case conMenu_View_Append '�Ƿ�λ��ʾLJͼ
        If mintLJͼ��λ��ʾ = 0 Then
            mintLJͼ��λ��ʾ = 1
        Else
            mintLJͼ��λ��ʾ = 0
        End If
        mlngItemID = -1 'ǿ��ˢ��
        Call zlRefOthers
    Case conMenu_View_Option

        Set panThis = Me.dkpMan.FindPane(conPane_Report)
        If panThis.Closed = False Then
            panThis.Close
        Else
            panThis.Select
            mlngItemID = -1 'ǿ��ˢ��
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
                Me.cbo����.ItemData(Me.cbo����.ListIndex), mlngItemID, _
                DateBegin, _
                 CLng(Val("" & Me.vfgReagent.TextMatrix(Me.vfgReagent.Row, mColG.ID))), dateEnd) Then
            Call zlRefRecord
            mlngItemID = -1 'ǿ��ˢ��
            Call zlRefOthers
        End If
    Case conMenu_Tool_Define
        If mlngItemID <= 0 Then Exit Sub
        If Me.vfgReagent.Rows - 1 > Me.vfgReagent.FixedRows Then
            If frmQCRedefine.ShowME(Me, _
                    Me.cbo����.ItemData(Me.cbo����.ListIndex), mlngItemID, _
                    CDate(Format(Me.dtp����(1).Value, "yyyy-MM-dd")), CLng(Val("" & Me.vfgReagent.TextMatrix(Me.vfgReagent.Row, mColG.ID)))) Then
                Call vfgItem_RowColChange
            End If
        Else
            If frmQCRedefine.ShowME(Me, Me.cbo����.ItemData(Me.cbo����.ListIndex), mlngItemID, CDate(Format(Me.dtp����(1).Value, "yyyy-MM-dd"))) Then
                Call vfgItem_RowColChange
            End If
        End If
        
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    
    Case conMenu_Tool_SignNew
        '����������
        If Me.cbo����.ListCount > 0 Then
            Call zlDatabase.SetPara("����", Me.cbo����.ItemData(Me.cbo����.ListIndex), glngSys, 1209)
            frmQCAddData.ShowME mstrPrivs, Me
        End If

    Case conMenu_Tool_SignVerify
        '����������
        If Me.cbo����.ListCount > 0 Then
            Call zlDatabase.SetPara("����", Me.cbo����.ItemData(Me.cbo����.ListIndex), glngSys, 1209)
            frmQCAddData1.ShowME mstrPrivs, Me
        End If
        
    Case conMenu_Manage_Discard '����
        Call Discard_OR_Recall(mLastCell, 1)
    Case conMenu_Manage_Recall  '����
        Call Discard_OR_Recall(mLastCell, 0)
    Case conMenu_Manage_Reset   '�鿴����ԭ��
        Call Discard_OR_Recall(mLastCell, 2)
    Case conMenu_View_Jump      '��ʾʧЧ����
        If mint��ʾʧЧ��¼ = 0 Then
            mint��ʾʧЧ��¼ = 1
            Call zlRefRecord
        Else
            mint��ʾʧЧ��¼ = 0
            Call zlRefRecord
        End If
    Case conMenu_Tool_Reference_1
        '��
        Call ItemMoveUpDown(1)
    Case conMenu_Tool_Reference_2
        '��
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
        Control.Enabled = (InStr(1, mstrPrivs, "����") > 0 And Me.vfgRecord.Cols > Me.vfgRecord.FixedCols)
    Case conMenu_Tool_Define
        Control.Enabled = (InStr(1, mstrPrivs, "��ֵ") > 0) And mlngItemID > 0
    Case conMenu_Tool_SignNew, conMenu_Tool_SignVerify
        Control.Enabled = (InStr(1, mstrPrivs, "�ʿؼ�¼����") > 0)
    Case conMenu_View_ShowAll
        Control.Checked = mblnShowAll
    Case conMenu_Edit_Save
        If mEditMode = 1 Then
            Control.Caption = "���汨��"
        Else
            Control.Caption = "���Ϊ"
        End If
    Case conMenu_Edit_Untread                           'ȡ��
        Control.Enabled = (mEditMode = 1)
        Control.Visible = (mEditMode = 1)
    Case conMenu_Edit_Adjust                            '��дʧ�ر���
        Control.Enabled = (mEditMode = 0 And mstrPigeonhole = "")
    Case conMenu_Edit_QCReport
        
    Case conMenu_Edit_Archive                           '�鵵
'        Control.Enabled = (mstrPigeonhole <> "")
    Case conMenu_Manage_Discard
        Control.Enabled = GetCellStat(mLastCell) = 1
    Case conMenu_Manage_Recall, conMenu_Manage_Reset
        Control.Enabled = GetCellStat(mLastCell) = 2
    Case conMenu_View_Jump
        Control.Checked = mint��ʾʧЧ��¼ = 1
    Case conMenu_View_Append
        Control.Checked = mintLJͼ��λ��ʾ = 1
    End Select
    
End Sub

Private Sub cmdˢ��_Click()
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
'        If Me.dtp����(0).Value < DateAdd("m", -3, Me.dtp����(1).Value) Then
'            '�����������������������û�и��ݣ����������µĲ�ѯ����ܶ���ʱ����
''            MsgBox "������ڿ�Ȳ��ܳ��������£�", vbInformation, gstrSysName
''            Me.dtp����(0).Value = DateAdd("m", -3, Me.dtp����(1).Value)
'        End If
'    Else
'        If Me.dtp����(0).Value > Me.dtp����(1).Value Then Me.dtp����(0).Value = Me.dtp����(1).Value
'        If Me.dtp����(1).Value > DateAdd("m", 3, Me.dtp����(0).Value) Then
'            '�����������������������û�и��ݣ����������µĲ�ѯ����ܶ���ʱ����
''            MsgBox "������ڿ�Ȳ��ܳ��������£�", vbInformation, gstrSysName
''            Me.dtp����(1).Value = DateAdd("m", 3, Me.dtp����(0).Value)
'        End If
'    End If
    Err = 0: On Error GoTo ErrHand
    
    If CDate(Format(Me.dtp����(1).Value, "yyyy-MM-dd")) < CDate(Format(Me.dtp����(0).Value, "yyyy-MM-dd")) Then
        MsgBox "�������ڲ��ܴ��ڿ�ʼ���ڣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSql = "Select M.ID, '' As ѡ��, M.���� , M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ, M.ˮƽ, to_Char(X.��ʼ����,'yy-MM-dd') as ��ʼ����,to_char(Nvl(X.��������, M.��������),'yy-MM-dd')  as ��������,�����ֵ,����SD,����cv," & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I, �����ʿؾ�ֵ X" & vbNewLine & _
            "Where M.ID = I.�ʿ�Ʒid And I.�ʿ�Ʒid = X.�ʿ�Ʒid And I.��Ŀid = X.��Ŀid And M.����id = [1] And I.��Ŀid = [2] And" & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.��ʼ���� And Nvl(X.��������, M.��������))" & vbNewLine & _
            "Order By M.��ʼ����, M.ˮƽ"
            
    gstrSql = "Select Id,ѡ��,����,�ʿ�Ʒ,Decode(substr(��ֵ,1,1),'.','0'||��ֵ,��ֵ) As ��ֵ,Decode(substr(SD,1,1),'.','0'||SD,SD) As SD,ˮƽ,min(��ʼ����) As ��ʼ����,Min(��������) As ��������,�����ֵ,����SD,����cv " & vbNewLine & _
            "From (" & vbNewLine & _
            "Select M.ID, '' As ѡ��, M.���� , M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ,I.��ֵ,I.SD, M.ˮƽ, to_Char(X.��ʼ����,'yy-MM-dd') as ��ʼ����,to_char(Nvl(X.��������, M.��������),'yy-MM-dd')  as ��������,x.��ֵ �����ֵ,x.sd ����SD,x.cv ����cv" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I, �����ʿؾ�ֵ X" & vbNewLine & _
            "Where M.ID = I.�ʿ�Ʒid And I.�ʿ�Ʒid = X.�ʿ�Ʒid And I.��Ŀid = X.��Ŀid And M.����id = [1] And I.��Ŀid = [2] And" & vbNewLine & _
            "      (To_Date([3], 'yyyy-MM-dd') Between X.��ʼ���� And Nvl(X.��������, M.��������))" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select M.ID, '' As ѡ��, M.���� , M.���� || ', ˮƽ:' || M.ˮƽ As �ʿ�Ʒ,I.��ֵ,I.SD, M.ˮƽ, to_Char(X.��ʼ����,'yy-MM-dd') as ��ʼ����,to_char(Nvl(X.��������, M.��������),'yy-MM-dd')  as ��������,x.��ֵ �����ֵ,x.sd ����SD,x.cv ����cv" & vbNewLine & _
            "From �����ʿ�Ʒ M, �����ʿ�Ʒ��Ŀ I, �����ʿؾ�ֵ X" & vbNewLine & _
            "Where M.ID = I.�ʿ�Ʒid And I.�ʿ�Ʒid = X.�ʿ�Ʒid And I.��Ŀid = X.��Ŀid And M.����id = [1] And I.��Ŀid = [2] And" & vbNewLine & _
            "        ( ( X.��ʼ���� Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd') )" & vbNewLine & _
            "         Or" & vbNewLine & _
            "          (nvl(X.��������,Sysdate) Between To_Date([3], 'yyyy-MM-dd') And To_Date([4], 'yyyy-MM-dd')+1-1/24*60*60)" & vbNewLine & _
            "         )" & vbNewLine & _
            "       )" & vbNewLine & _
            "Group By      Id,ѡ��,����,�ʿ�Ʒ,��ֵ,SD,ˮƽ,�����ֵ,����SD,����cv " & vbNewLine & _
            "Order By �ʿ�Ʒ,ˮƽ"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, _
                CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex)), mlngItemID, _
                CStr(Format(Me.dtp����(0), "yyyy-MM-dd")), CStr(Format(Me.dtp����(1), "yyyy-MM-dd")))
    
    With Me.vfgReagent
        .FixedRows = 1
        Set .DataSource = rsTemp
        .ColWidth(mColG.ѡ��) = 500
        .ColWidth(mColG.����) = 900
        .ColWidth(mColG.ID) = 0: .ColWidth(mColG.ˮƽ) = 0
        .ColHidden(mColG.ID) = True: .ColHidden(mColG.ˮƽ) = True
        .ColHidden(mColG.��ֵ) = True: .ColHidden(mColG.SD) = True
        .ColHidden(mColG.�����ֵ) = True
        .ColHidden(mColG.����SD) = True
        .ColHidden(mColG.����cv) = True
        For lngCount = .FixedRows To .Rows - 1
'            If lngCount <= Val(Me.cbo����.Tag) Then
                .Cell(flexcpChecked, lngCount, mColG.ѡ��) = flexChecked
'            Else
'                .Cell(flexcpChecked, lngCount, mColG.ѡ��) = flexUnchecked
'            End If
        Next
        If .Rows <= .FixedRows Then
            '�Զ������ֵ��SD��д�����ݿ�,��Ϊ0������ʾ
            MsgBox "��δ��ֵ,������ͼ,�����¶�ֵ��", vbInformation, Me.Caption

        End If
    End With
    
    mLastStartDate = CDate(Format(dtp����(0).Value, "yyyy-MM-dd"))
    mLastEndDate = CDate(Format(dtp����(1).Value, "yyyy-MM-dd"))


    'ˢ�½������
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
    Dim lngDeptID As Long  '����ID
    '-----------------------------------------------------
    'Ȩ�����ƴ����ƣ�����ͬʱ��������ģ�������gstrPrivs�仯�����¿�����Ч
'
    '�����м�վҪֱ�����������������һ�½ű�
    gstrPrivs = GetPrivFunc(100, 1209)
    mstrPrivs = gstrPrivs
    mlngListWidth = Me.picRecord.Width
    Me.picReport.BackColor = Me.BackColor
    Call zlCommFun.SetWindowsInTaskBar(Me.hWnd, False)
    
    If Val(zlDatabase.GetPara("�����ʿ�Ʒ", glngSys, 1209, 1)) = 1 Then
        Me.vfgReagent.Visible = False
    Else
        Me.vfgReagent.Visible = True
    End If
    
    lngDeptID = Val(zlDatabase.GetPara("����", glngSys, 1209, 0))
    
    mblnShowAll = Val(zlDatabase.GetPara("��ʾ����ʧ����Ŀ", glngSys, 1209, 0)) = 1
    mint��ʾʧЧ��¼ = Val(zlDatabase.GetPara("��ʾʧЧ����", glngSys, 1209, 1))
    mintLJͼ��λ��ʾ = Val(zlDatabase.GetPara("LJͼ��λ��ʾ", glngSys, 1209, 1))
    
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
'    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ������ͼ")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ����ͼ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_BatPrint, "��ӡ�ʿؽ��(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "������ͼ(&S)...")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��(&C)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "���ƿ���ͼ(&C)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "��ʾʧЧ����(&V)"): cbrControl.BeginGroup = True
        'Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "�����ʿر���(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Hide, "�����ʿ�Ʒѡ��(&H)"): cbrControl.Style = xtpButtonCaption
        Set cbrControl = .Add(xtpControlButton, conMenu_View_ShowAll, "��ʾ����ʧ�ر���(&H)"):
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Append, "LJͼ��λ��ʾ(&L)"):
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    cbrMenuBar.ID = xtpControlPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "���ý��(Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "���Ž��(R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "�鿴(�޸�)����ԭ��(V)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignNew, "�����������ʿؼ�¼(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_SignVerify, "�����������ʿؼ�¼(&C)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "ʧ�ؼ���(&Y)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "���¶�ֵ(&N)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "ʧ�ر���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCReport, "�ʿر���"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵")
        
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "����")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "����")
    cbrCustom.Handle = Me.cbo����.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "����")
    cbrControl.Flags = xtpFlagRightAlign
    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "����")
    cbrCustom.Handle = Me.cbo����.hWnd: cbrCustom.Flags = xtpFlagRightAlign
'    Set cbrControl = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlLabel, 0, "��Ŀ")
'    cbrControl.Flags = xtpFlagRightAlign
'    Set cbrCustom = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlCustom, 0, "��Ŀ")
'    cbrCustom.Handle = Me.cbo��Ŀ.hWnd: cbrCustom.Flags = xtpFlagRightAlign
    
'    '�Ҽ��˵�
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, ID_MENU_MOUSE, "�Ҽ��˵�", -1, False)
    cbrMenuBar.ID = ID_MENU_MOUSE
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Discard, "���ý��(Q)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Recall, "���Ž��(R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Manage_Reset, "�鿴(�޸�)����ԭ��(V)"): cbrControl.BeginGroup = True

    End With
    cbrMenuBar.Visible = False
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("C"), conMenu_Edit_MarkMap
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    
        .Add 0, VK_UP, conMenu_Tool_Reference_1
        .Add 0, VK_DOWN, conMenu_Tool_Reference_2
    
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_PrintSet
        .AddHiddenCommand conMenu_Edit_MarkMap
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Refresh
    End With
    
    Call zlDatabase.ShowReportMenu(Me.cbsThis, glngSys, glngModul, mstrPrivs)
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "���Ϊ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ��"): cbrControl.BeginGroup = True
        'Set cbrControl = .Add(xtpControlButton, conMenu_View_Option, "�����ʿر���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Analyse, "ʧ�ؼ���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Define, "���¶�ֵ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Adjust, "ʧ�ر���"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_QCReport, "�ʿر���"): cbrControl.BeginGroup = True
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Archive, "�鵵")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '����ͣ������
    Dim panThis As Pane, panChild As Pane, panSub As Pane, panCalc As Pane
    
    With Me.dkpMan
        Set panThis = .CreatePane(conPane_Record, 200, 400, DockLeftOf, Nothing)
        panThis.Title = "�ʿؽ����"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        Set panThis = .CreatePane(conPane_Charts, 400, 500, DockRightOf, Nothing)
        panThis.Title = "�ʿ�ͳ��ͼ"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panChild = .CreatePane(conPane_Data, 400, 200, DockBottomOf, panThis)
        panChild.Title = "������"
        panChild.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        
        Set panSub = .CreatePane(conPane_Report, 400, 200, DockBottomOf)
        panSub.Title = "�ʿر���"
        panSub.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        panSub.AttachTo panChild
        
        Set panCalc = .CreatePane(conPane_Calc, 400, 200, DockBottomOf)
        panCalc.Title = "�ʿؽ��ͳ��"
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
        panThis.Title = "��������"
        panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
        .Options.ThemedFloatingFrames = True
        .Options.HideClient = False
    End With
    
    '-----------------------------------------------------
    '���ñ�񸽼Ӵ���
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
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
    
    '-----------------------------------------------------
    'װ���������
    Dim rsTemp As New ADODB.Recordset
    
    Me.dtp����(1).Value = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd")): Me.dtp����(0).Value = CDate(Format(Me.dtp����(1).Value, "yyyy-MM") & "-01")
    Err = 0: On Error GoTo ErrHand
    
    If InStr(1, mstrPrivs, "���п���") > 0 Then
        gstrSql = " Select Distinct b.Id, b.���� , b.���� As ���� From �������� a ,���ű� b,�����ʿ�Ʒ c " & _
                  "Where a.ʹ��С��ID = b.ID and a.id = c.����id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName)
        
    Else
        gstrSql = "Select Distinct B.ID, B.����, B.���� As ����" & vbNewLine & _
                " From �������� A, ���ű� B, �����ʿ�Ʒ C" & vbNewLine & _
                " Where A.ʹ��С��id = B.ID And A.ID = C.����id And" & vbNewLine & _
                "      A.ʹ��С��id In (Select Distinct D.ʹ��С��id" & vbNewLine & _
                "                   From ����С���Ա A, ����С�� B, ����С������ C, �������� D" & vbNewLine & _
                "                   Where A.С��id = B.ID And B.ID = C.С��id��and ��Աid = [1] And C.����id = D.ID)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, gstrSysName, UserInfo.ID)
    End If
    
    Me.cbo����.Clear
    Do Until rsTemp.EOF
        Me.cbo����.AddItem rsTemp("����") & "-" & rsTemp("����")
        Me.cbo����.ItemData(Me.cbo����.NewIndex) = rsTemp("Id")
        If rsTemp("ID") = lngDeptID Then
            Me.cbo����.ListIndex = Me.cbo����.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    If Me.cbo����.ListCount = 0 Then MsgBox "��δ�������ʹ��С������ã�", vbInformation, gstrSysName: Unload Me: Exit Sub
    If cbo����.ListIndex = -1 Then
        Me.cbo����.ListIndex = 0
    End If
    If Me.cbo����.ListCount = 1 Then Me.cbo����.Enabled = False
    
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
        Call zlDatabase.SetPara("�����ʿ�Ʒ", 0, glngSys, 1209)
    Else
        Call zlDatabase.SetPara("�����ʿ�Ʒ", 1, glngSys, 1209)
    End If
    
    If Me.cbo����.ListCount > 0 Then
        Call zlDatabase.SetPara("����", Me.cbo����.ItemData(Me.cbo����.ListIndex), glngSys, 1209)
    End If
    If Me.cbo����.ListCount > 0 Then
        Call zlDatabase.SetPara("����", Me.cbo����.ItemData(Me.cbo����.ListIndex), glngSys, 1209)
    End If
    If mlngItemID > 0 Then
        Call zlDatabase.SetPara("��Ŀ", mlngItemID, glngSys, 1209)
    End If
    Call zlDatabase.SetPara("��ʾ����ʧ����Ŀ", IIf(mblnShowAll, 1, 0), glngSys, 1209)
    Call zlDatabase.SetPara("��ʾʧЧ����", mint��ʾʧЧ��¼, glngSys, 1209)
    Call zlDatabase.SetPara("LJͼ��λ��ʾ", mintLJͼ��λ��ʾ, glngSys, 1209)
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picCalc_Resize()
    Err = 0: On Error Resume Next

    '�����б�
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

    '�����б�
    With Me.vfgRecord
        .Left = Me.picData.ScaleLeft: .Width = Me.picData.ScaleWidth - .Left
        .Top = Me.picData.ScaleTop
        .Height = Me.picData.ScaleHeight - .Top
    End With
End Sub

Private Sub picRecord_Resize()
    Err = 0: On Error Resume Next
'    Me.cbo�ڼ�.Width = Me.picRecord.ScaleWidth - Me.cbo�ڼ�.Left
    Me.cmdˢ��.Left = Me.picRecord.ScaleWidth - Me.cmdˢ��.Width - 15
    Me.dtp����(1).Width = Me.picRecord.ScaleWidth - Me.cmdˢ��.Width - 15 - Me.dtp����(1).Left - 15
    Me.dtp����(0).Width = Me.dtp����(1).Width

    
    With Me.vfgReagent
        .Left = Me.picRecord.ScaleLeft: .Width = Me.picRecord.ScaleWidth - .Left
        .Height = (.Rows + 1.5) * 300
        .Top = Me.picRecord.ScaleHeight - .Height
    End With
    
    '�ʿ���Ŀ�б�
    With Me.vfgItem
        .Left = Me.picRecord.ScaleLeft: .Width = Me.picRecord.ScaleWidth - .Left
        .Height = Me.picRecord.ScaleHeight - .Top - IIf(Me.vfgReagent.Visible, Me.vfgReagent.Height - 15, 0)
    End With
    
End Sub

Private Sub tbcCharts_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    mlngItemID = -1 'ǿ��ˢ��
    If Me.Visible Then Call zlRefOthers
End Sub

Private Sub vfgItem_RowColChange()

    
    If mLastStartDate <> CDate(0) And mLastEndDate <> CDate(0) Then
        Me.dtp����(0) = CDate(Format(mLastStartDate, "yyyy-MM-dd"))
        Me.dtp����(1) = CDate(Format(mLastEndDate, "yyyy-MM-dd"))
    
    Else
        Me.dtp����(0) = CDate(Format(Now, "yyyy-MM-01"))
        Me.dtp����(1) = CDate(Format(Now, "yyyy-MM-dd"))
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
        If .Cell(flexcpChecked, .Row, mColG.ѡ��) = flexUnchecked Then
            .Cell(flexcpChecked, .Row, mColG.ѡ��) = flexChecked
            'Me.vfgRecord.ColWidth(.Row + mColL.���� + 1) = 900
            Me.vfgRecord.RowHidden(.Row + mColL.����) = False
        Else
            .Cell(flexcpChecked, .Row, mColG.ѡ��) = flexUnchecked
            'Me.vfgRecord.ColWidth(.Row + mColL.���� + 1) = 0
            Me.vfgRecord.RowHidden(.Row + mColL.����) = True
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
    If NewRow <= mColL.���� Or NewRow >= Me.vfgRecord.Rows - 2 Then
        Cancel = True
    Else
        If NewRow - mColL.���� - 1 >= 0 And NewRow - mColL.���� <= Me.vfgReagent.Rows - 1 Then
            On Error Resume Next
            Me.vfgReagent.Row = NewRow - mColL.����
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
    '��ӡ��Ԥ���ʿ�ͼ
    '����           intPrintMode =1 ��ӡ =2 Ԥ��
    '               intPrintType 0=LJ 1=FQ 2=ZS 3=YD 4=CS 5=MN
    
    Dim rsTmp As New ADODB.Recordset
    Dim strPrintType As String                  '��Ӧ�ĵ���
    Dim strQCID As String                       '�ʿ�ƷID���ܻ�����","�ָ��Ķ��ID
    Dim lngQCID As Long                         '�����ʿ�ƷID
    Dim lngItemID As String                     '��ĿID
    Dim lngMachine As Long                      '����ID
    Dim intLoop As Integer
    
    
    On Error GoTo errH
    
    strPrintType = "ZL1_INSIDE_1209_"
    strPrintType = strPrintType & intPrintType + 1
    
    gstrSql = "Select b.w, b.h " & vbNewLine & _
                " From Zlreports a, Zlrptitems b" & vbNewLine & _
                " Where a.Id = b.����id And a.��� = [1] And b.���� = '�ʿ�ͼ'"
                
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strPrintType)
    'û���ҵ�ʱ�˳�
    If rsTmp.EOF Then
        MsgBox "�ڵ��ݶ�����û�ж���<�ʿ�ͼ>,���ڵ����ж���һ����Ϊ<�ʿ�ͼ>��ͼ���!", vbQuestion, Me.Caption
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
    
    
    '�õ��ʿ�ƷID
    Select Case intPrintType
        Case mTab.LJ
            lngQCID = mfrmChartLJ.ZLGetLJ_QCID
            With Me.vfgReagent
                strQCID = ""
                For lngCount = 0 To .Rows - 1
                    If .Cell(flexcpChecked, lngCount, mColG.ѡ��) = flexChecked Then
                        strQCID = strQCID & "," & .TextMatrix(lngCount, mColG.ID)
                    End If
                Next
                If strQCID <> "" Then strQCID = Mid(strQCID, 2)
            End With
        Case mTab.FQ
            With Me.vfgReagent
                strQCID = ""
                For lngCount = 0 To .Rows - 1
                    If .Cell(flexcpChecked, lngCount, mColG.ѡ��) = flexChecked Then
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
    
    '�õ���ĿID
    If mlngItemID = 0 Then Exit Sub
    lngItemID = mlngItemID
    lngMachine = CLng(Me.cbo����.ItemData(Me.cbo����.ListIndex))
    
    If Dir(App.path & "\QC0.jpg") <> "" Then
        Call ReportOpen(gcnOracle, glngSys, strPrintType, Me, "�ʿ�ͼ=" & App.path & "\QC0.jpg", _
        "�ʿ�ƷID=" & lngQCID, "��ĿID=" & lngItemID, "��ʼ����=" & Format(dtp����(0), "yyyy-MM-dd"), "��������=" & Format(dtp����(1), "yyyy-MM-dd"), _
        "����ID=" & lngMachine, "�ʿ�Ʒ��=" & IIf(strQCID = "", "0", strQCID), _
        "�ʿ�ͼ1=" & App.path & "\QC1.jpg", "�ʿ�ͼ2=" & App.path & "\QC2.jpg", _
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
    lngItemID = Val(Me.vfgReport.TextMatrix(vfgReport.Row, mColR.������Ŀid))
    
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
    '���û����
    'intType : 0-����  1-����  2-�鿴���ý��
    Dim lngRow As Long, lngCol As Long, lngID As Long, strTmp As String
    Dim str�����ʿ�Ʒ As String, str��ǰ�ʿ�Ʒ As String, lng_S As Long, lng_E As Long
    Dim strSQL As String, strԭ�� As String
    Dim frmDiscard As New frmQCDiscardEdit
    On Error GoTo errH
    
    If InStr(strCell, ",") > 0 Then
        lngRow = Val(Split(strCell, ",")(0))
        lngCol = Val(Split(strCell, ",")(1))
        
        With vfgRecord
            If Not (lngCol >= .FixedCols And lngCol < .Cols And lngRow > 3 And lngRow <= .Rows - 2) Then
                MsgBox "��ѡ��һ�����ݵ�Ԫ�����ʹ�ô˹��ܣ�"
                Exit Function
            End If
            If Trim(.TextMatrix(lngRow, lngCol)) <> "" Then
                str�����ʿ�Ʒ = .TextMatrix(mColL.ID, 0)
                str��ǰ�ʿ�Ʒ = .TextMatrix(lngRow, 0)
                lng_S = InStr(str�����ʿ�Ʒ, "|" & str��ǰ�ʿ�Ʒ & "=")
                If lng_S > 0 Then
                    lng_E = InStr(lng_S + 1, str�����ʿ�Ʒ, "|")
                    If lng_E > lng_S Then
                        str��ǰ�ʿ�Ʒ = Mid(str�����ʿ�Ʒ, lng_S, lng_E - lng_S)
                    Else
                        str��ǰ�ʿ�Ʒ = Mid(str�����ʿ�Ʒ, lng_S)
                    End If
                    str��ǰ�ʿ�Ʒ = Split(str��ǰ�ʿ�Ʒ, "=")(1)
                End If
                
                strTmp = .TextMatrix(mColL.ID, lngCol)
                lng_S = InStr(strTmp, "|" & str��ǰ�ʿ�Ʒ & "=")
                If lng_S > 0 Then
                    lng_E = InStr(lng_S + 1, strTmp, "|")
                    
                    If lng_E > lng_S Then
                        str��ǰ�ʿ�Ʒ = Mid(strTmp, lng_S, lng_E - lng_S)
                    Else
                        str��ǰ�ʿ�Ʒ = Mid(strTmp, lng_S)
                    End If
                    lngID = Val(Split(str��ǰ�ʿ�Ʒ, "=")(1))
                End If
                
                If lngID > 0 Then
                    If intType = 1 Then
                        If frmDiscard.ShowME(lngID, strԭ��, Me) Then
                            strSQL = "zl_������ͨ���_����(" & lngID & ",1,'" & strԭ�� & "')"
                            zlDatabase.ExecuteProcedure strSQL, Me.Caption
                            .Cell(flexcpBackColor, lngRow, lngCol) = &HC0C0C0
                            Discard_OR_Recall = True
                        End If
                    ElseIf intType = 0 Then
                        strSQL = "zl_������ͨ���_����(" & lngID & ",0)"
                        zlDatabase.ExecuteProcedure strSQL, Me.Caption
                        .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor
                        Discard_OR_Recall = True
                    Else
                        If frmDiscard.ShowME(lngID, strԭ��, Me) Then
                            strSQL = "zl_������ͨ���_����(" & lngID & ",2,'" & strԭ�� & "')"
                            zlDatabase.ExecuteProcedure strSQL, Me.Caption
                            Discard_OR_Recall = True
                        End If
                    End If
                    Set frmDiscard = Nothing
                    If Discard_OR_Recall Then
                        Call zlRefRecord
                        .Select lngRow, lngCol
                        mlngItemID = -1 'ǿ��ˢ��
                        Call zlRefOthers
                    End If
                    Exit Function
                End If
            Else
                MsgBox "��ѡ��һ���ǿյ����ݵ�Ԫ��"
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
    'ȡ��Ԫ��״̬,���ڿ��Ʋ˵�
    '����ֵ  0-����������ع��� 1-�������� 2-��������
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
    '���¼�����
    On Error Resume Next
    With Me.vfgItem
        If intUpDown = 1 Then
            If .Row - 1 > .FixedRows Then .Select .Row - 1, .Col
        Else
            If .Row + 1 < .Rows Then .Select .Row + 1, .Col
        End If
    End With
End Sub
