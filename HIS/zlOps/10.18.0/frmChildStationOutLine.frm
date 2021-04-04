VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmChildStationOutLine 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13125
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3285
      Index           =   2
      Left            =   9300
      ScaleHeight     =   3285
      ScaleWidth      =   5340
      TabIndex        =   3
      Top             =   5115
      Width           =   5340
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1530
         Index           =   2
         Left            =   990
         TabIndex        =   32
         Top             =   105
         Width           =   3900
         _cx             =   6879
         _cy             =   2699
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1365
         Index           =   3
         Left            =   990
         TabIndex        =   33
         Top             =   1695
         Width           =   3675
         _cx             =   6482
         _cy             =   2408
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "术前诊断"
         Height          =   180
         Index           =   12
         Left            =   210
         TabIndex        =   35
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "术后诊断"
         Height          =   180
         Index           =   11
         Left            =   195
         TabIndex        =   34
         Top             =   1725
         Width           =   720
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3450
      Index           =   1
      Left            =   3495
      ScaleHeight     =   3450
      ScaleWidth      =   5370
      TabIndex        =   2
      Top             =   4635
      Width           =   5370
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1695
         Index           =   0
         Left            =   1005
         TabIndex        =   28
         Top             =   90
         Width           =   3915
         _cx             =   6906
         _cy             =   2990
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1395
         Index           =   1
         Left            =   1005
         TabIndex        =   29
         Top             =   1860
         Width           =   3930
         _cx             =   6932
         _cy             =   2461
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
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483638
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   270
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
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "拟行手术"
         Height          =   180
         Index           =   10
         Left            =   180
         TabIndex        =   31
         Top             =   105
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已行手术"
         Height          =   180
         Index           =   9
         Left            =   180
         TabIndex        =   30
         Top             =   1845
         Width           =   720
      End
   End
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   4725
      Index           =   0
      Left            =   150
      ScaleHeight     =   4725
      ScaleWidth      =   9945
      TabIndex        =   1
      Top             =   -75
      Width           =   9945
      Begin VB.Frame fra 
         Height          =   4305
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11655
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   2
            Left            =   5910
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   1290
            Width           =   2310
         End
         Begin VB.CommandButton cmd 
            Height          =   330
            Index           =   1
            Left            =   4725
            Picture         =   "frmChildStationOutLine.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   40
            TabStop         =   0   'False
            ToolTipText     =   "多选，快捷键：F3"
            Top             =   1260
            Width           =   345
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   3300
            ScaleHeight     =   240
            ScaleWidth      =   1755
            TabIndex        =   39
            Top             =   960
            Width           =   1755
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   1200
            ScaleHeight     =   240
            ScaleWidth      =   1770
            TabIndex        =   38
            Top             =   960
            Width           =   1770
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   3285
            ScaleHeight     =   240
            ScaleWidth      =   1755
            TabIndex        =   37
            Top             =   1680
            Width           =   1755
         End
         Begin VB.PictureBox picConver 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   1200
            ScaleHeight     =   240
            ScaleWidth      =   1770
            TabIndex        =   36
            Top             =   1680
            Width           =   1770
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   1
            Left            =   1155
            TabIndex        =   9
            Top             =   1290
            Width           =   3570
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   0
            Left            =   5910
            MaxLength       =   10
            TabIndex        =   8
            Top             =   570
            Width           =   2310
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   3
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   915
            Width           =   2310
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            Left            =   5910
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   210
            Width           =   2310
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   555
            Width           =   3930
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   0
            Left            =   1155
            TabIndex        =   10
            Top             =   195
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   82444291
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   1
            Left            =   3255
            TabIndex        =   11
            Top             =   210
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   82444291
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   2
            Left            =   1155
            TabIndex        =   12
            Top             =   930
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   82444291
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   3
            Left            =   3255
            TabIndex        =   13
            Top             =   930
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   82444291
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   4
            Left            =   1155
            TabIndex        =   14
            Top             =   1650
            Width           =   1845
            _ExtentX        =   3254
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   82444291
            CurrentDate     =   39275
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   300
            Index           =   5
            Left            =   3240
            TabIndex        =   15
            Top             =   1650
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   82444291
            CurrentDate     =   39275
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1785
            Index           =   4
            Left            =   1155
            TabIndex        =   16
            Top             =   2010
            Width           =   7200
            _cx             =   12700
            _cy             =   3149
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
            GridColor       =   -2147483626
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483638
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   270
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
         End
         Begin VB.CheckBox chk 
            Caption         =   "麻醉时间"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   42
            Top             =   990
            Width           =   1095
         End
         Begin VB.CheckBox chk 
            Caption         =   "输氧时间"
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   43
            Top             =   1695
            Width           =   1095
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "麻醉质量"
            Height          =   180
            Index           =   7
            Left            =   5145
            TabIndex        =   27
            Top             =   990
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "麻醉类型"
            Height          =   180
            Index           =   6
            Left            =   5145
            TabIndex        =   26
            Top             =   1350
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "麻醉方式"
            Height          =   180
            Index           =   5
            Left            =   345
            TabIndex        =   25
            Top             =   1335
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "输液总量"
            Height          =   180
            Index           =   3
            Left            =   5145
            TabIndex        =   24
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "手 术 间"
            Height          =   180
            Index           =   2
            Left            =   5145
            TabIndex        =   23
            Top             =   270
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "手术规模"
            Height          =   180
            Index           =   1
            Left            =   345
            TabIndex        =   22
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "手术时间"
            Height          =   180
            Index           =   0
            Left            =   345
            TabIndex        =   21
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "手术人员"
            Height          =   180
            Index           =   13
            Left            =   345
            TabIndex        =   20
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            Height          =   180
            Index           =   14
            Left            =   3015
            TabIndex        =   19
            Top             =   270
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            Height          =   180
            Index           =   15
            Left            =   3015
            TabIndex        =   18
            Top             =   975
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "～"
            Height          =   180
            Index           =   16
            Left            =   3015
            TabIndex        =   17
            Top             =   1695
            Width           =   180
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbc 
      Height          =   2550
      Left            =   615
      TabIndex        =   0
      Top             =   6120
      Width           =   2820
      _Version        =   589884
      _ExtentX        =   4974
      _ExtentY        =   4498
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmChildStationOutLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'######################################################################################################################
'（１）窗体级变量定义
Private mlngLoop As Long
Private mstrSQL As String
Private mblnStartUp As Boolean
Private mlngKey As Long
Private mlngDeptKey As Long
Private mfrmMain As Object
Private mblnDataChanged As Boolean
Private mblnAllowModify As Boolean
Private mblnReading As Boolean

Private WithEvents mclsVsfPerson As clsVsf
Attribute mclsVsfPerson.VB_VarHelpID = -1
Private WithEvents mclsVsfOpsBefore As clsVsf
Attribute mclsVsfOpsBefore.VB_VarHelpID = -1
Private WithEvents mclsVsfOpsAfter As clsVsf
Attribute mclsVsfOpsAfter.VB_VarHelpID = -1
Private WithEvents mclsVsfDiagBefore As clsVsf
Attribute mclsVsfDiagBefore.VB_VarHelpID = -1
Private WithEvents mclsVsfDiagAfter As clsVsf
Attribute mclsVsfDiagAfter.VB_VarHelpID = -1
Private Type Items
    麻醉方式 As String
End Type

Private usrSaveItem As Items
Public Event AfterDataChanged()

'######################################################################################################################
'（２）自定义过程或函数

Public Property Let DataChanged(ByVal blnData As Boolean)
    mblnDataChanged = blnData
    
    If mblnReading = False Then
        RaiseEvent AfterDataChanged
    End If
End Property

Public Property Get DataChanged() As Boolean
    DataChanged = mblnDataChanged
End Property

Public Function InitData(ByVal frmMain As Object, Optional ByVal blnAllowModify As Boolean = True, Optional ByVal lngDeptKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngDeptKey = lngDeptKey
    
    Set mfrmMain = frmMain
    
    
    If ExecuteCommand("初始控件") = False Then Exit Function
    If ExecuteCommand("初始数据") = False Then Exit Function
    
    Call ExecuteCommand("控件状态")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, Optional ByVal blnAllowModify As Boolean = True, Optional ByVal lngDeptKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngDeptKey = lngDeptKey
    mlngKey = lngKey
    
    Call ExecuteCommand("清空数据")
    Call ExecuteCommand("控件状态")
    
    If mlngKey > 0 Then
        If ExecuteCommand("读取数据") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '功能：校验编辑数据的有效性
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    If dtp(0).Value > dtp(1).Value Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "手术开始时间不能大于手术结束时间！"
        Call LocationObj(dtp(0))
        Exit Function
    End If
    
    If Abs(DateDiff("h", CDate(Format(dtp(0).Value, "YYYY-MM-DD HH:MM")), CDate(Format(dtp(1).Value, "YYYY-MM-DD HH:MM")))) > 12 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "手术开始时间和手术结束时间之间不能大于12小时！"
        Call LocationObj(dtp(0))
        Exit Function
    End If
    
    If dtp(2).Value > dtp(3).Value And chk(2).Value = 1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "麻醉开始时间不能大于麻醉结束时间！"
        Call LocationObj(dtp(2))
        Exit Function
    End If
    
    If chk(2).Value = 1 And Trim(txt(1).Text) = "" Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "必须指明麻醉方式！"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    If chk(2).Value = 1 And cbo(3).ListIndex = -1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "必须指明麻醉质量！"
        Call LocationObj(cbo(0))
        Exit Function
    End If
    
    If dtp(4).Value > dtp(5).Value And chk(4).Value = 1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "输氧开始时间不能大于输氧结束时间！"
        Call LocationObj(dtp(4))
        Exit Function
    End If
        
    If CheckAllNumber(txt(0).Text) = False Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "输液总量必须为全数字！"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    With vsf(4)
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 And InStr(.TextMatrix(lngLoop, .ColIndex("岗位")), "主刀医生") > 0 Then
                Exit For
            End If
        Next
        If lngLoop = .Rows Then
            tbc.Item(0).Selected = True
            ShowSimpleMsg " 必须指定手术的主刀医生！"
            Call LocationGrid(vsf(4), 1, .ColIndex("姓名"))
            Exit Function
        End If
    End With
    
    '检查手术名称是否有非法字符、超长、手术个数
    For lngIndex = 0 To 1
        With vsf(lngIndex)
            For lngLoop = 1 To .Rows - 1
                If Val(.RowData(lngLoop)) > 0 Then
                    Exit For
                End If
            Next
                
            If lngLoop = .Rows Then
                tbc.Item(1).Selected = True
                If lngIndex = 0 Then
                    ShowSimpleMsg "至少有一个拟行手术！"
                Else
                    ShowSimpleMsg "至少有一个已行手术！"
                End If

                Call LocationGrid(vsf(lngIndex), 1, .ColIndex("手术名称"))
                Exit Function
            End If
        
        End With
    Next
    
    '检查诊断描述是否有非法字符、超长
    For lngIndex = 2 To 3
        With vsf(lngIndex)
            For lngLoop = 1 To .Rows - 1
                If Val(.RowData(lngLoop)) > 0 Then
                    If StrIsValid(.TextMatrix(lngLoop, .ColIndex("诊断描述")), 100) = False Then
                        tbc.Item(2).Selected = True
                        Call LocationGrid(vsf(lngIndex), lngLoop, .ColIndex("诊断描述"))
                        Exit Function
                    End If
                End If
            Next
        End With
    Next
    
    ValidData = True
    
End Function

Public Function SaveData(ByRef rsSQL As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lngOrderKey As Long
    Dim lng病人id As Long
    Dim lng主页id As Long
    Dim lngRow As Long

    On Error GoTo errHand
    
    strSQL = "Select a.* From 病人医嘱记录 a,病人手术记录 b Where a.ID=b.医嘱id And b.ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
    If rs.BOF = True Then Exit Function
    
    lng病人id = zlCommFun.NVL(rs("病人id").Value, 0)
    lng主页id = zlCommFun.NVL(rs("主页id").Value, 0)
    lngOrderKey = zlCommFun.NVL(rs("ID").Value, 0)
    
    '
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "zl_病人手术记录_Update(" & mlngKey & "," & _
                                        "To_Date('" & Format(dtp(0).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                        "To_Date('" & Format(dtp(0).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                        "To_Date('" & Format(dtp(1).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                        IIf(chk(2).Value = 1, "To_Date('" & Format(dtp(2).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                        IIf(chk(2).Value = 1, "To_Date('" & Format(dtp(3).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                        IIf(chk(2).Value = 1, "'" & txt(1).Text & "'", "Null") & "," & _
                                        IIf(chk(2).Value = 1, "'" & txt(2).Text & "'", "Null") & "," & _
                                        IIf(chk(2).Value = 1, "'" & zlCommFun.GetNeedName(cbo(3).Text) & "'", "Null") & "," & _
                                        Val(txt(0).Text) & "," & _
                                        IIf(chk(4).Value = 1, "To_Date('" & Format(dtp(4).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & "," & _
                                        IIf(chk(4).Value = 1, "To_Date('" & Format(dtp(5).Value, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')", "NULL") & ",'" & _
                                        cbo(1).Text & "'," & _
                                        mlngDeptKey & ",'" & _
                                        cbo(0).Text & "'," & _
                                        "NULL)"
    Call SQLRecordAdd(rsSQL, strSQL)
            
    '手术人员
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "zl_病人手术人员_Delete(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(4)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Then
                strSQL = "zl_病人手术人员_Insert(" & mlngKey & ",'" & .TextMatrix(lngRow, .ColIndex("岗位")) & "'," & Val(.RowData(lngRow)) & ",'" & .TextMatrix(lngRow, .ColIndex("编号")) & "','" & .TextMatrix(lngRow, .ColIndex("姓名")) & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '拟行手术
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_病人手术情况_DELETE(" & mlngKey & ",1)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(0)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Then

                If Left(.TextMatrix(lngRow, .ColIndex("编码方式")), 1) = 1 Then
                    strSQL = "zl_病人手术情况_Insert(" & mlngKey & ",1," & Abs(Val(.TextMatrix(lngRow, .ColIndex("缺省")))) & ",'" & .TextMatrix(lngRow, .ColIndex("手术名称")) & "',Null," & Val(.RowData(lngRow)) & ")"
                Else
                    strSQL = "zl_病人手术情况_Insert(" & mlngKey & ",1," & Abs(Val(.TextMatrix(lngRow, .ColIndex("缺省")))) & ",'" & .TextMatrix(lngRow, .ColIndex("手术名称")) & "'," & Val(.RowData(lngRow)) & ",Null)"
                End If
                
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '已行手术
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_病人手术情况_DELETE(" & mlngKey & ",2)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(1)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Then

                If Left(.TextMatrix(lngRow, .ColIndex("编码方式")), 1) = 1 Then
                    strSQL = "zl_病人手术情况_Insert(" & mlngKey & ",2," & Abs(Val(.TextMatrix(lngRow, .ColIndex("缺省")))) & ",'" & .TextMatrix(lngRow, .ColIndex("手术名称")) & "',Null," & Val(.RowData(lngRow)) & ")"
                Else
                    strSQL = "zl_病人手术情况_Insert(" & mlngKey & ",2," & Abs(Val(.TextMatrix(lngRow, .ColIndex("缺省")))) & ",'" & .TextMatrix(lngRow, .ColIndex("手术名称")) & "'," & Val(.RowData(lngRow)) & ",Null)"
                End If
                
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '术前诊断
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_病人诊断记录_DELETE2(" & lngOrderKey & ",8)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(2)
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) > 0 Or Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) > 0 Then
                
                strSQL = "zl_病人诊断记录_Insert(" & lng病人id & "," & ZVal(lng主页id) & ",1,Null,8," & Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) & "," & Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) & ",Null,'" & .TextMatrix(lngRow, .ColIndex("诊断描述")) & "',Null,Null,Null,Sysdate," & lngOrderKey & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            End If
        Next
    End With
    
    '术后诊断
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_病人诊断记录_DELETE2(" & lngOrderKey & ",9)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(3)
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) > 0 Or Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) > 0 Then
                
                strSQL = "zl_病人诊断记录_Insert(" & lng病人id & "," & ZVal(lng主页id) & ",1,Null,9," & Val(.TextMatrix(lngRow, .ColIndex("疾病id"))) & "," & Val(.TextMatrix(lngRow, .ColIndex("诊断id"))) & ",Null,'" & .TextMatrix(lngRow, .ColIndex("诊断描述")) & "',Null,Null,Null,Sysdate," & lngOrderKey & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            End If
        Next
    End With

    
    SaveData = True
    
    Exit Function
    
    '出错处理
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitTabControl()
    '******************************************************************************************************************
    '功能:
    '参数:
    '返回:
    '******************************************************************************************************************
    With tbc
        With .PaintManager

            .Appearance = xtpTabAppearanceVisio
            .ClientFrame = xtpTabFrameSingleLine
            .COLOR = xtpTabColorOffice2003
            .ColorSet.ButtonSelected = &HFFC0C0     '&HD2BDB6
            .ColorSet.ButtonNormal = &HFFC0C0       '&HD2BDB6
            .ShowIcons = True
        End With
        Set .Icons = frmPubIcons.imgPublic.Icons
        
        .InsertItem 0, "基本情况", picPane(0).hWnd, 0
        .InsertItem 1, "手术情况", picPane(1).hWnd, 0
        .InsertItem 2, "诊断情况", picPane(2).hWnd, 0
        
        .Item(1).Selected = True
        .Item(2).Selected = True
        .Item(0).Selected = True
        
    End With
    
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    
    On Error GoTo errHand
    
    mblnReading = True
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "初始控件"
        
        '手术人员
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfPerson = New clsVsf
        With mclsVsfPerson
            Call .Initialize(Me.Controls, vsf(4), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("岗位", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("姓名", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("编号", 900, flexAlignLeftCenter, flexDTString, "", "编码", True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        '拟行手术
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfOpsBefore = New clsVsf
        With mclsVsfOpsBefore
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("编码方式", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("手术名称", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("缺省", 810, flexAlignCenterCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        '已行手术
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfOpsAfter = New clsVsf
        With mclsVsfOpsAfter
            Call .Initialize(Me.Controls, vsf(1), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", IIf(mblnAllowModify, "[指示器]", "[图标]"), False)
            Call .AppendColumn("编码方式", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("手术名称", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("缺省", 810, flexAlignCenterCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
                
        '术前诊断
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfDiagBefore = New clsVsf
        With mclsVsfDiagBefore
            Call .Initialize(Me.Controls, vsf(2), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("疾病编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("疾病id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("诊断描述", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        '术后诊断
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfDiagAfter = New clsVsf
        With mclsVsfDiagAfter
            Call .Initialize(Me.Controls, vsf(3), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
            Call .AppendColumn("疾病编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断编码", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("诊断id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("疾病id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("诊断描述", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With

        txt(2).BackColor = COLOR.锁色
    '--------------------------------------------------------------------------------------------------------------
    Case "初始数据"
                 
        '设置最大输入长度
        '--------------------------------------------------------------------------------------------------------------
        txt(0).MaxLength = 10

        '诊疗手术规模
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT 名称,0 FROM 诊疗手术规模"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
                
        '麻醉质量
        '--------------------------------------------------------------------------------------------------------------
        With cbo(3)
            .Clear
            .AddItem "1-优"
            .AddItem "2-佳"
            .AddItem "3-劣"
            .AddItem "4-危(急)"
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "控件状态"
    
        blnAllowModify = mblnAllowModify
        If mlngKey <= 0 Then blnAllowModify = False
        
        txt(0).Locked = Not blnAllowModify
        txt(1).Locked = Not blnAllowModify
        txt(2).Locked = Not blnAllowModify
        cbo(0).Locked = Not blnAllowModify
        cbo(1).Locked = Not blnAllowModify
        cbo(3).Locked = Not blnAllowModify
        
        cmd(1).Enabled = blnAllowModify
        dtp(0).Enabled = blnAllowModify
        dtp(1).Enabled = blnAllowModify
        dtp(2).Enabled = blnAllowModify
        dtp(3).Enabled = blnAllowModify
        dtp(4).Enabled = blnAllowModify
        dtp(5).Enabled = blnAllowModify
        
        chk(2).Enabled = blnAllowModify
        chk(4).Enabled = blnAllowModify
        
        If blnAllowModify Then
        
            With mclsVsfPerson
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                Call .InitializeEdit(True, True, True)

                '手术岗位
                '------------------------------------------------------------------------------------------------------
                gstrSQL = "SELECT 名称 FROM 手术岗位 Order by 编码"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                Call .InitializeEditColumn(.ColIndex("岗位"), True, vbVsfEditCombox, vsf(4).BuildComboList(rs, "名称", "名称"))
                
                Call .InitializeEditColumn(.ColIndex("姓名"), True, vbVsfEditCommand)
                
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
                
        
            End With
        
            With mclsVsfOpsBefore
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("编码方式"), True, vbVsfEditCombox, "1-诊疗|2-疾病")
                Call .InitializeEditColumn(.ColIndex("手术名称"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("缺省"), True, vbVsfEditCheck)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            End With
        
            With mclsVsfOpsAfter
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("编码方式"), True, vbVsfEditCombox, "1-诊疗|2-疾病")
                Call .InitializeEditColumn(.ColIndex("手术名称"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("缺省"), True, vbVsfEditCheck)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            End With
            
            With mclsVsfDiagBefore
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("疾病编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("诊断编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("诊断描述"), True, vbVsfEditText)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            End With
            
            With mclsVsfDiagAfter
                Call .ModifyColumn(.ColIndex("图标"), "", 255, flexAlignCenterCenter, flexDTString, "", "[指示器]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("疾病编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("诊断编码"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("诊断描述"), True, vbVsfEditText)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("当前").Picture
            End With
        
        Else
            
            mclsVsfPerson.AllowEdit = False
            mclsVsfOpsBefore.AllowEdit = False
            mclsVsfOpsAfter.AllowEdit = False
            mclsVsfDiagBefore.AllowEdit = False
            mclsVsfDiagAfter.AllowEdit = False
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "清空数据"
        
        txt(0).Text = ""
        txt(1).Text = ""
        txt(2).Text = ""
        chk(2).Value = 0
        chk(4).Value = 0
        
        mclsVsfPerson.ClearGrid
        mclsVsfDiagAfter.ClearGrid
        mclsVsfDiagBefore.ClearGrid
        mclsVsfOpsAfter.ClearGrid
        mclsVsfOpsBefore.ClearGrid
        
        DataChanged = False
    '------------------------------------------------------------------------------------------------------------------
    Case "读取数据"
        
        '医技执行房间
        '--------------------------------------------------------------------------------------------------------------
        cbo(1).Clear
        gstrSQL = "SELECT RowNum As ID,执行间 As 名称 FROM 医技执行房间 WHERE 科室id=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptKey)
        If rs.BOF = False Then Call AddComboData(cbo(1), rs)
        
        '1.读取手术基本资料
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT A.*,C.性别,C.当前科室id,C.住院号 FROM 病人手术记录 A,病人信息 C WHERE A.病人id=C.病人id AND A.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            
'            mlng病人id = zlCommFun.NVL(rs("病人id"), 0)
'            mlng主页id = zlCommFun.NVL(rs("主页id"), 0)
'            mlngDeptKey = zlCommFun.NVL(rs("当前科室id"), 0)
            
'            If zlCommFun.NVL(rs("性别")) Like "*男*" Then mstr性别 = mstr性别 & ",1"
'            If zlCommFun.NVL(rs("性别")) Like "*女*" Then mstr性别 = mstr性别 & ",2"
            
            If IsNull(rs("手术开始时间")) = False Then
                dtp(0).Value = Format(zlCommFun.NVL(rs("手术开始时间")), dtp(0).CustomFormat)
                dtp(1).Value = Format(zlCommFun.NVL(rs("手术结束时间")), dtp(1).CustomFormat)

                
                If IsNull(rs("麻醉开始时间")) = False Then
                    chk(2).Value = 1
                    picConver(2).Visible = False
                    picConver(3).Visible = False
                    dtp(2).Value = Format(zlCommFun.NVL(rs("麻醉开始时间")), dtp(2).CustomFormat)
                    dtp(3).Value = Format(zlCommFun.NVL(rs("麻醉结束时间")), dtp(3).CustomFormat)
                Else
                    chk(2).Value = 0
                    picConver(2).Visible = True
                    picConver(3).Visible = True
                    dtp(2).Value = Format(zlCommFun.NVL(rs("手术开始时间")), dtp(2).CustomFormat)
                    dtp(3).Value = Format(zlCommFun.NVL(rs("手术结束时间")), dtp(3).CustomFormat)
                End If
                
                If IsNull(rs("输氧开始时间")) = False Then
                    chk(4).Value = 1
                    picConver(4).Visible = False
                    picConver(5).Visible = False
                    dtp(4).Value = Format(zlCommFun.NVL(rs("输氧开始时间")), dtp(4).CustomFormat)
                    dtp(5).Value = Format(zlCommFun.NVL(rs("输氧结束时间")), dtp(5).CustomFormat)
                Else
                    chk(4).Value = 0
                    picConver(4).Visible = True
                    picConver(5).Visible = True
                    dtp(4).Value = Format(zlCommFun.NVL(rs("手术开始时间")), dtp(4).CustomFormat)
                    dtp(5).Value = Format(zlCommFun.NVL(rs("手术结束时间")), dtp(5).CustomFormat)
                End If

            End If
            
            zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("手术规模").Value)
            zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("手术间").Value)
            zlControl.CboLocate cbo(3), zlCommFun.NVL(rs("麻醉质量").Value)

            txt(1).Text = zlCommFun.NVL(rs("麻醉方式").Value)
            txt(2).Text = zlCommFun.NVL(rs("麻醉类型").Value)
            txt(0).Text = zlCommFun.NVL(rs("输液总量").Value)
        End If
        
        
        '手术人员
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfPerson.ClearGrid
        gstrSQL = "Select A.人员id As ID,A.岗位,B.编号 As 编码,B.姓名 From 病人手术人员 a,人员表 b Where a.记录id=[1] And a.人员id=b.ID"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsfPerson.LoadGrid(rs)
        
        '拟行手术
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfOpsBefore.ClearGrid
        
        gstrSQL = GetPublicSQL(SQL.病人手术情况)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 1)
        If rs.BOF = False Then Call mclsVsfOpsBefore.LoadGrid(rs)
        
        '已行手术
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfOpsAfter.ClearGrid
        gstrSQL = GetPublicSQL(SQL.病人手术情况)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 2)
        If rs.BOF = False Then Call mclsVsfOpsAfter.LoadGrid(rs)
        
        '拟行诊断
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfDiagBefore.ClearGrid
        gstrSQL = GetPublicSQL(SQL.病人诊断记录)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 8)
        If rs.BOF = False Then Call mclsVsfDiagBefore.LoadGrid(rs)
        
        '已行诊断
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfDiagAfter.ClearGrid
        gstrSQL = GetPublicSQL(SQL.病人诊断记录)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 9)
        If rs.BOF = False Then Call mclsVsfDiagAfter.LoadGrid(rs)
        
    End Select
    
    mblnReading = False
    
    ExecuteCommand = True

    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    mblnReading = False
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub cbo_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub cbo_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub chk_Click(Index As Integer)

    If Index = 2 Then
        picConver(2).Visible = Not (chk(Index).Value = 1)
        picConver(3).Visible = Not (chk(Index).Value = 1)
        
        If cbo(3).Enabled = False Then
            cbo(3).ListIndex = -1
        ElseIf cbo(3).ListIndex = -1 Then
            cbo(3).ListIndex = 0
        End If
    Else
        picConver(4).Visible = Not (chk(Index).Value = 1)
        picConver(5).Visible = Not (chk(Index).Value = 1)
    End If
    
    DataChanged = True
    
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    Select Case Index
    '------------------------------------------------------------------------------------------------------------------
    Case 1      '麻醉方式
        gstrSQL = GetPublicSQL(SQL.麻醉方式选择)
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
 
        If ShowPubSelect(Me, txt(1), 2, "编码,900,0,;名称,2400,0,;麻醉类型,900,0,", Me.Name & "\麻醉方式选择", "请从下表中选择一个麻醉方式", rsData, rs, 8790, 4500) = 1 Then
            

            txt(1).Text = zlCommFun.NVL(rs("名称").Value)
            txt(2).Text = zlCommFun.NVL(rs("麻醉类型").Value)

            txt(1).Tag = ""

            usrSaveItem.麻醉方式 = txt(1).Text
            
            DataChanged = True


        End If

    End Select
End Sub


Private Sub dtp_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub Form_Load()
    Me.BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    fra(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    picPane(0).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    picPane(1).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    picPane(2).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(2).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    chk(4).BackColor = COLOR_NativeXpPlain.SpecialGroupClient
    Call InitTabControl
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tbc.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set mclsVsfPerson = Nothing
    Set mclsVsfOpsBefore = Nothing
    Set mclsVsfOpsAfter = Nothing
    Set mclsVsfDiagBefore = Nothing
    Set mclsVsfDiagAfter = Nothing
    
End Sub

Private Sub mclsVsfDiagAfter_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(3).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfDiagAfter_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(3)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub mclsVsfDiagBefore_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(2).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfDiagBefore_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(2)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub mclsVsfOpsAfter_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(1).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfOpsAfter_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(1)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub mclsVsfOpsBefore_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Val(vsf(0).RowData(Row)) > 0 Then
        DataChanged = True
    End If
End Sub

Private Sub mclsVsfOpsBefore_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf(0)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
End Sub

Private Sub mclsVsfPerson_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If Val(vsf(4).RowData(Row)) > 0 Then
        DataChanged = True
    End If
    
End Sub

Private Sub mclsVsfPerson_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)

    With vsf(4)
        Cancel = (Val(.RowData(Row)) <= 0)
    End With
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 0
        fra(0).Move 0, -75, picPane(Index).Width, picPane(Index).Height + 75
        
        cbo(1).Move cbo(1).Left, cbo(1).Top, fra(0).Width - cbo(1).Left - 75
        cbo(3).Move cbo(3).Left, cbo(3).Top, fra(0).Width - cbo(3).Left - 75
        txt(0).Move txt(0).Left, txt(0).Top, fra(0).Width - txt(0).Left - 75
        txt(2).Move txt(2).Left, txt(2).Top, fra(0).Width - txt(2).Left - 75
        
        vsf(4).Move vsf(4).Left, vsf(4).Top, fra(0).Width - vsf(4).Left - 75, fra(0).Height - vsf(4).Top - 75
        mclsVsfPerson.AppendRows = True

    Case 1
    
        vsf(0).Move vsf(0).Left, vsf(0).Top, picPane(Index).Width - vsf(0).Left - 75
        vsf(1).Move vsf(1).Left, vsf(1).Top, picPane(Index).Width - vsf(1).Left - 75, picPane(Index).Height - vsf(1).Top - 75
        
        mclsVsfOpsBefore.AppendRows = True
        mclsVsfOpsAfter.AppendRows = True
    Case 2
        vsf(2).Move vsf(2).Left, vsf(2).Top, picPane(Index).Width - vsf(2).Left - 75
        vsf(3).Move vsf(3).Left, vsf(3).Top, picPane(Index).Width - vsf(3).Left - 75, picPane(Index).Height - vsf(3).Top - 75
        
        mclsVsfDiagBefore.AppendRows = True
        mclsVsfDiagAfter.AppendRows = True
        
    End Select
End Sub


Private Sub txt_Change(Index As Integer)
    If mblnReading Then Exit Sub
    
    DataChanged = True
    
    Select Case Index
    Case 1
        txt(Index).Tag = "Changed"
    End Select

End Sub

Private Sub txt_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txt(Index)
    
    Select Case Index
    Case 1
        zlCommFun.OpenIme True
    End Select
        
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 1
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            txt(Index).Text = ""
            txt(2).Text = ""
            cmd(1).Tag = ""
            txt(Index).Tag = ""
            usrSaveItem.麻醉方式 = ""
        End If
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytMode As Byte
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case Index
        Case 1
            If txt(Index).Tag <> "" Then
                txt(Index).Tag = ""
                
                strText = UCase(txt(Index).Text)
                bytMode = GetApplyMode(strText)

                strText = strText & "%"
                strTmp = strText & IIf(ParamInfo.项目输入匹配方式 = 1, "", "%")
                
                gstrSQL = GetPublicSQL(SQL.麻醉方式过滤, bytMode)
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "编码,990,0,1;名称,1500,0,0;麻醉类型,900,0,0", Me.Name & "\麻醉方式过滤", "请从下面选择一个麻醉方式", rsData, rs) = 1 Then

                    txt(Index).Text = zlCommFun.NVL(rs("名称").Value)
                    txt(2).Text = zlCommFun.NVL(rs("麻醉类型").Value)
                    
                    DataChanged = True
                    
                    usrSaveItem.麻醉方式 = txt(Index).Text

                Else
                    txt(Index).Text = usrSaveItem.麻醉方式
                    txt(Index).Tag = ""
                    Exit Sub
                End If

            End If
        Case Else
            zlCommFun.PressKey vbKeyTab
        End Select
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)

    Select Case Index
    Case 1
        zlCommFun.OpenIme False
    End Select
    
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
    If Cancel Then Exit Sub

    Select Case Index
    Case 1
        If (txt(Index).Tag = "Changed") Then
            txt(Index).Text = usrSaveItem.麻醉方式
            txt(Index).Tag = ""
        End If
    End Select
    
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.AfterEdit(Row, Col)
    Case 1
        Call mclsVsfOpsAfter.AfterEdit(Row, Col)
    Case 2
        Call mclsVsfDiagBefore.AfterEdit(Row, Col)
    Case 3
        Call mclsVsfDiagAfter.AfterEdit(Row, Col)
    Case 4
        Call mclsVsfPerson.AfterEdit(Row, Col)
    End Select
    
    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Case 1
        Call mclsVsfOpsAfter.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Case 2
        Call mclsVsfDiagBefore.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Case 3
        Call mclsVsfDiagAfter.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Case 4
        Call mclsVsfPerson.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    End Select
End Sub

Private Sub vsf_AfterMoveColumn(Index As Integer, ByVal Col As Long, Position As Long)
    Select Case Index
    Case 0
        mclsVsfOpsBefore.AppendRows = True
    Case 1
        mclsVsfOpsAfter.AppendRows = True
    Case 2
        mclsVsfDiagBefore.AppendRows = True
    Case 3
        mclsVsfDiagAfter.AppendRows = True
    Case 4
        mclsVsfPerson.AppendRows = True
    End Select
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Select Case Index
    Case 0
        mclsVsfOpsBefore.AppendRows = True
    Case 1
        mclsVsfOpsAfter.AppendRows = True
    Case 2
        mclsVsfDiagBefore.AppendRows = True
    Case 3
        mclsVsfDiagAfter.AppendRows = True
    Case 4
        mclsVsfPerson.AppendRows = True
    End Select
End Sub


Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Select Case Index
    Case 0
        mclsVsfOpsBefore.AppendRows = True
    Case 1
        mclsVsfOpsAfter.AppendRows = True
    Case 2
        mclsVsfDiagBefore.AppendRows = True
    Case 3
        mclsVsfDiagAfter.AppendRows = True
    Case 4
        mclsVsfPerson.AppendRows = True
    End Select
End Sub

Private Sub vsf_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim bytRet As Byte
    Dim strTmp As String
    
    With vsf(Index)
        Select Case Index
        '--------------------------------------------------------------------------------------------------------------
        Case 0, 1                               '拟行、已行手术
            
            If Col = .ColIndex("手术名称") Then

                If Val(Left(.TextMatrix(Row, .ColIndex("编码方式")), 1)) = 1 Then
                    '诊疗编码
                    gstrSQL = GetPublicSQL(SQL.手术项目选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目选择", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                Else
                    '疾病编码
                    gstrSQL = GetPublicSQL(SQL.疾病编码选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "S")
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码选择", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                End If

                If bytRet = 1 Then
                    If Index = 0 Then
                        If mclsVsfOpsBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
                    Else
                        If mclsVsfOpsAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
                    End If
                    
                    .EditText = zlCommFun.NVL(rs("名称").Value)
                    .TextMatrix(Row, .ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 2, 3                               '术前术后诊断
            If Col = .ColIndex("疾病编码") Or Col = .ColIndex("诊断编码") Then
                Select Case Col
                '------------------------------------------------------------------------------------------------------
                Case .ColIndex("疾病编码")
                
                    gstrSQL = GetPublicSQL(SQL.疾病编码选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "D")
        
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码选择", "请从下表中选择一个疾病编码项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then
                        
                        If Index = 2 Then
                            If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                Exit Sub
                            End If
                        Else
                            If mclsVsfDiagAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                Exit Sub
                            End If
                        End If
                        .EditText = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                        .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)
                        
                        DataChanged = True
                    End If
                '----------------------------------------------------------------------------------------------------------
                Case .ColIndex("诊断编码")
                
                    gstrSQL = GetPublicSQL(SQL.疾病诊断选择)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
        
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "编码,1200,0,;名称,2700,0,", Me.Name & "\疾病诊断选择", "请从下表中选择一个疾病诊断项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then
                        If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                            Exit Sub
                        End If
            
                        .EditText = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                        .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))
                
                        DataChanged = True
                    End If
                End Select
                
                '----------------------------------------------------------------------------------------------------------
                If bytRet = 1 Then
                    If rsData.BOF = False Then
                        .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("疾病编码").Value)
                        .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("诊断编码").Value)
                        .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("疾病id").Value, 0)
                        .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("诊断id").Value, 0)
                    End If
                End If
                
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 4                                      '手术人员
            
            If Col = .ColIndex("姓名") Then
    
                gstrSQL = GetPublicSQL(SQL.人员信息选择)
                
                strTmp = "医生"
                If InStr(.TextMatrix(.Row, .ColIndex("岗位")), "护士") > 0 Then strTmp = "护士"
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strTmp, mlngDeptKey)
    
                If ShowPubSelect(Me, vsf(Index), 2, "编号,1200,0,;姓名,1500,0,;简码,900,0,;科室,1200,0,", Me.Name & "\人员信息选择", "请从下表中选择一个人员", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsfPerson.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "选择的人员“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
                        Exit Sub
                    End If
                           
                    .EditText = zlCommFun.NVL(rs("姓名").Value)
                    .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
                    .TextMatrix(Row, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                End If
            End If
            
        End Select
    End With
End Sub

Private Sub vsf_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.KeyDown(KeyCode, Shift)
    Case 1
        Call mclsVsfOpsAfter.KeyDown(KeyCode, Shift)
    Case 2
        Call mclsVsfDiagBefore.KeyDown(KeyCode, Shift)
    Case 3
        Call mclsVsfDiagAfter.KeyDown(KeyCode, Shift)
    Case 4
        Call mclsVsfPerson.KeyDown(KeyCode, Shift)
    End Select
End Sub

Private Sub vsf_KeyDownEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    Dim bytMode As Byte
    Dim bytRet As Byte
    Dim strClass As String
    
    With vsf(Index)
        If KeyCode = vbKeyReturn Then
        
            If InStr(.EditText, "'") > 0 Then
                KeyCode = 0
                .EditText = ""
                Exit Sub
            End If
            strText = UCase(.EditText)
            bytMode = GetApplyMode(strText)
            strText = strText & "%"
            strTmp = IIf(ParamInfo.项目输入匹配方式 = 1, strText, "%" & strText)
                    
            Select Case Index
            Case 0, 1
                If Col = .ColIndex("手术名称") Then
                    
                    If Val(Left(.TextMatrix(Row, .ColIndex("编码方式")), 1)) = 1 Then
                        gstrSQL = GetPublicSQL(SQL.手术项目过滤, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\手术项目过滤", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    Else
                        gstrSQL = GetPublicSQL(SQL.疾病编码过滤, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "S")
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码过滤", "请从下表中选择一个手术项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    End If

                    If bytRet = 1 Then
                        
                        If Index = 0 Then
                            If mclsVsfOpsBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                Exit Sub
                            End If
                        Else
                            If mclsVsfOpsAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                Exit Sub
                            End If
                        End If
    
                        .EditText = zlCommFun.NVL(rs("名称").Value)
                        .TextMatrix(Row, .ColIndex("手术名称")) = zlCommFun.NVL(rs("名称").Value)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
    
                        DataChanged = True
    
                    Else
                        KeyCode = 0

                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
    
                    End If
                End If
            Case 2, 3
                
                If Col = .ColIndex("疾病编码") Or Col = .ColIndex("诊断编码") Then
                                                
                    Select Case Col
                    '--------------------------------------------------------------------------------------------------
                    Case .ColIndex("疾病编码")
    
                        gstrSQL = GetPublicSQL(SQL.疾病编码过滤, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "D")
        
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "编码,1200,0,;名称,2700,0,;简码,900,0,;附码,900,0,", Me.Name & "\疾病编码过滤", "请从下表中选择一个疾病编码项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        If bytRet = 1 Then
                            
                            If Index = 2 Then
                                If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            Else
                                If mclsVsfDiagAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            End If
        
                            .EditText = zlCommFun.NVL(rs("编码").Value)
                            .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("编码").Value)
                            .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                            .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("ID").Value, 0)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                            
                            gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)
                        
                            DataChanged = True
                        End If
                    '--------------------------------------------------------------------------------------------------
                    Case .ColIndex("诊断编码")
                        gstrSQL = GetPublicSQL(SQL.疾病诊断过滤, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
        
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "编码,1200,0,;名称,2700,0,", Me.Name & "\疾病诊断过滤", "请从下表中选择一个疾病诊断项目", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        If bytRet = 1 Then
                            
                            If Index = 2 Then
                                If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            Else
                                If mclsVsfDiagAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "选择的项目“" & zlCommFun.NVL(rs("名称").Value) & "”已被选择！"
                                    Exit Sub
                                End If
                            End If
        
                            .EditText = zlCommFun.NVL(rs("编码").Value)
                            .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("编码").Value)
                            .TextMatrix(Row, .ColIndex("诊断描述")) = zlCommFun.NVL(rs("名称").Value)
                            .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("ID").Value, 0)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
        
                            gstrSQL = GetPublicSQL(SQL.疾病诊断对照)
                            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))
                        
                            DataChanged = True
                        End If
                    End Select
                    
                    If bytRet = 1 Then
                    
                        '--------------------------------------------------------------------------------------------------
                        If rsData.BOF = False Then
                            
                            .TextMatrix(Row, .ColIndex("疾病编码")) = zlCommFun.NVL(rs("疾病编码").Value)
                            .TextMatrix(Row, .ColIndex("诊断编码")) = zlCommFun.NVL(rs("诊断编码").Value)
                            
                            .TextMatrix(Row, .ColIndex("疾病id")) = zlCommFun.NVL(rs("疾病id").Value, 0)
                            .TextMatrix(Row, .ColIndex("诊断id")) = zlCommFun.NVL(rs("诊断id").Value, 0)
                        End If
                
                    Else
                        KeyCode = 0
                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    End If
                    
                End If
            Case 4
                
                If Col = .ColIndex("姓名") Then
        
                    gstrSQL = GetPublicSQL(SQL.人员信息过滤, bytMode)
                    
                    strClass = "医生"
                    If InStr(.TextMatrix(.Row, .ColIndex("岗位")), "护士") > 0 Then strClass = "护士"
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strClass, mlngDeptKey, strText, strTmp)
        
                    If ShowPubSelect(Me, vsf(Index), 2, "编号,1200,0,;姓名,1500,0,;简码,900,0,;科室,1200,0,", Me.Name & "\人员信息过滤", "请从下表中选择一个人员", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then
    
                        If mclsVsfPerson.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "选择的人员“" & zlCommFun.NVL(rs("姓名").Value) & "”已被选择！"
                            Exit Sub
                        End If
                               
                        .EditText = zlCommFun.NVL(rs("姓名").Value)
                        .TextMatrix(Row, .ColIndex("姓名")) = zlCommFun.NVL(rs("姓名").Value)
                        .TextMatrix(Row, .ColIndex("编号")) = zlCommFun.NVL(rs("编号").Value)
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        DataChanged = True
                    End If
                End If
            
            End Select
        Else
            DataChanged = True
        End If
    End With
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.KeyPress(KeyAscii)
    Case 1
        Call mclsVsfOpsAfter.KeyPress(KeyAscii)
    Case 2
        Call mclsVsfDiagBefore.KeyPress(KeyAscii)
    Case 3
        Call mclsVsfDiagAfter.KeyPress(KeyAscii)
    Case 4
        Call mclsVsfPerson.KeyPress(KeyAscii)
    End Select
    
End Sub

Private Sub vsf_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.KeyPressEdit(KeyAscii)
    Case 1
        Call mclsVsfOpsAfter.KeyPressEdit(KeyAscii)
    Case 2
        Call mclsVsfDiagBefore.KeyPressEdit(KeyAscii)
    Case 3
        Call mclsVsfDiagAfter.KeyPressEdit(KeyAscii)
    Case 4
        Call mclsVsfPerson.KeyPressEdit(KeyAscii)
    End Select
End Sub

Private Sub vsf_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Select Case Index
        Case 0
            Call mclsVsfOpsBefore.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 1
            Call mclsVsfOpsAfter.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 2
            Call mclsVsfDiagBefore.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 3
            Call mclsVsfDiagAfter.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        Case 4
            Call mclsVsfPerson.AutoAddRow(vsf(Index).MouseRow, vsf(Index).MouseCol)
        End Select
    End Select
End Sub

Private Sub vsf_SetupEditWindow(Index As Integer, ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.EditSelAll
    Case 1
        Call mclsVsfOpsAfter.EditSelAll
    Case 2
        Call mclsVsfDiagBefore.EditSelAll
    Case 3
        Call mclsVsfDiagAfter.EditSelAll
    Case 4
        Call mclsVsfPerson.EditSelAll
    End Select
End Sub

Private Sub vsf_StartEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.BeforeEdit(Row, Col, Cancel)
    Case 1
        Call mclsVsfOpsAfter.BeforeEdit(Row, Col, Cancel)
    Case 2
        Call mclsVsfDiagBefore.BeforeEdit(Row, Col, Cancel)
    Case 3
        Call mclsVsfDiagAfter.BeforeEdit(Row, Col, Cancel)
    Case 4
        Call mclsVsfPerson.BeforeEdit(Row, Col, Cancel)
    End Select
End Sub

Private Sub vsf_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Select Case Index
    Case 0
        Call mclsVsfOpsBefore.ValidateEdit(Col, Cancel)
    Case 1
        Call mclsVsfOpsAfter.ValidateEdit(Col, Cancel)
    Case 2
        Call mclsVsfDiagBefore.ValidateEdit(Col, Cancel)
    Case 3
        Call mclsVsfDiagAfter.ValidateEdit(Col, Cancel)
    Case 4
        Call mclsVsfPerson.ValidateEdit(Col, Cancel)
    End Select
End Sub


