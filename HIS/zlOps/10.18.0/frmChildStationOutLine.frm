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
   StartUpPosition =   3  '����ȱʡ
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
         Caption         =   "��ǰ���"
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
         Caption         =   "�������"
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
         Caption         =   "��������"
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
         Caption         =   "��������"
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
            ToolTipText     =   "��ѡ����ݼ���F3"
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
            Caption         =   "����ʱ��"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   42
            Top             =   990
            Width           =   1095
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ��"
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
            Caption         =   "��������"
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
            Caption         =   "��������"
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
            Caption         =   "����ʽ"
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
            Caption         =   "��Һ����"
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
            Caption         =   "�� �� ��"
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
            Caption         =   "������ģ"
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
            Caption         =   "����ʱ��"
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
            Caption         =   "������Ա"
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
            Caption         =   "��"
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
            Caption         =   "��"
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
            Caption         =   "��"
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
'���������弶��������
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
    ����ʽ As String
End Type

Private usrSaveItem As Items
Public Event AfterDataChanged()

'######################################################################################################################
'�������Զ�����̻���

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
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngDeptKey = lngDeptKey
    
    Set mfrmMain = frmMain
    
    
    If ExecuteCommand("��ʼ�ؼ�") = False Then Exit Function
    If ExecuteCommand("��ʼ����") = False Then Exit Function
    
    Call ExecuteCommand("�ؼ�״̬")
    
    DataChanged = False
    
    InitData = True
    
End Function

Public Function RefreshData(ByVal lngKey As Long, Optional ByVal blnAllowModify As Boolean = True, Optional ByVal lngDeptKey As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mblnAllowModify = blnAllowModify
    mlngDeptKey = lngDeptKey
    mlngKey = lngKey
    
    Call ExecuteCommand("�������")
    Call ExecuteCommand("�ؼ�״̬")
    
    If mlngKey > 0 Then
        If ExecuteCommand("��ȡ����") = False Then Exit Function
    End If

    DataChanged = False
    
    RefreshData = True
    
End Function

Public Function ValidData() As Boolean
    '******************************************************************************************************************
    '���ܣ�У��༭���ݵ���Ч��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim lngIndex As Long
    
    If dtp(0).Value > dtp(1).Value Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "������ʼʱ�䲻�ܴ�����������ʱ�䣡"
        Call LocationObj(dtp(0))
        Exit Function
    End If
    
    If Abs(DateDiff("h", CDate(Format(dtp(0).Value, "YYYY-MM-DD HH:MM")), CDate(Format(dtp(1).Value, "YYYY-MM-DD HH:MM")))) > 12 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "������ʼʱ�����������ʱ��֮�䲻�ܴ���12Сʱ��"
        Call LocationObj(dtp(0))
        Exit Function
    End If
    
    If dtp(2).Value > dtp(3).Value And chk(2).Value = 1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "����ʼʱ�䲻�ܴ����������ʱ�䣡"
        Call LocationObj(dtp(2))
        Exit Function
    End If
    
    If chk(2).Value = 1 And Trim(txt(1).Text) = "" Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "����ָ������ʽ��"
        Call LocationObj(txt(1))
        Exit Function
    End If
    
    If chk(2).Value = 1 And cbo(3).ListIndex = -1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "����ָ������������"
        Call LocationObj(cbo(0))
        Exit Function
    End If
    
    If dtp(4).Value > dtp(5).Value And chk(4).Value = 1 Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "������ʼʱ�䲻�ܴ�����������ʱ�䣡"
        Call LocationObj(dtp(4))
        Exit Function
    End If
        
    If CheckAllNumber(txt(0).Text) = False Then
        tbc.Item(0).Selected = True
        ShowSimpleMsg "��Һ��������Ϊȫ���֣�"
        Call LocationObj(txt(0))
        Exit Function
    End If
    
    With vsf(4)
        For lngLoop = 1 To .Rows - 1
            If Val(.RowData(lngLoop)) > 0 And InStr(.TextMatrix(lngLoop, .ColIndex("��λ")), "����ҽ��") > 0 Then
                Exit For
            End If
        Next
        If lngLoop = .Rows Then
            tbc.Item(0).Selected = True
            ShowSimpleMsg " ����ָ������������ҽ����"
            Call LocationGrid(vsf(4), 1, .ColIndex("����"))
            Exit Function
        End If
    End With
    
    '������������Ƿ��зǷ��ַ�����������������
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
                    ShowSimpleMsg "������һ������������"
                Else
                    ShowSimpleMsg "������һ������������"
                End If

                Call LocationGrid(vsf(lngIndex), 1, .ColIndex("��������"))
                Exit Function
            End If
        
        End With
    Next
    
    '�����������Ƿ��зǷ��ַ�������
    For lngIndex = 2 To 3
        With vsf(lngIndex)
            For lngLoop = 1 To .Rows - 1
                If Val(.RowData(lngLoop)) > 0 Then
                    If StrIsValid(.TextMatrix(lngLoop, .ColIndex("�������")), 100) = False Then
                        tbc.Item(2).Selected = True
                        Call LocationGrid(vsf(lngIndex), lngLoop, .ColIndex("�������"))
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
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim lngOrderKey As Long
    Dim lng����id As Long
    Dim lng��ҳid As Long
    Dim lngRow As Long

    On Error GoTo errHand
    
    strSQL = "Select a.* From ����ҽ����¼ a,����������¼ b Where a.ID=b.ҽ��id And b.ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngKey)
    If rs.BOF = True Then Exit Function
    
    lng����id = zlCommFun.NVL(rs("����id").Value, 0)
    lng��ҳid = zlCommFun.NVL(rs("��ҳid").Value, 0)
    lngOrderKey = zlCommFun.NVL(rs("ID").Value, 0)
    
    '
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "zl_����������¼_Update(" & mlngKey & "," & _
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
            
    '������Ա
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "zl_����������Ա_Delete(" & mlngKey & ")"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(4)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Then
                strSQL = "zl_����������Ա_Insert(" & mlngKey & ",'" & .TextMatrix(lngRow, .ColIndex("��λ")) & "'," & Val(.RowData(lngRow)) & ",'" & .TextMatrix(lngRow, .ColIndex("���")) & "','" & .TextMatrix(lngRow, .ColIndex("����")) & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_�����������_DELETE(" & mlngKey & ",1)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(0)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Then

                If Left(.TextMatrix(lngRow, .ColIndex("���뷽ʽ")), 1) = 1 Then
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",1," & Abs(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "',Null," & Val(.RowData(lngRow)) & ")"
                Else
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",1," & Abs(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "'," & Val(.RowData(lngRow)) & ",Null)"
                End If
                
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_�����������_DELETE(" & mlngKey & ",2)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(1)
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) > 0 Then

                If Left(.TextMatrix(lngRow, .ColIndex("���뷽ʽ")), 1) = 1 Then
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",2," & Abs(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "',Null," & Val(.RowData(lngRow)) & ")"
                Else
                    strSQL = "zl_�����������_Insert(" & mlngKey & ",2," & Abs(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ")))) & ",'" & .TextMatrix(lngRow, .ColIndex("��������")) & "'," & Val(.RowData(lngRow)) & ",Null)"
                End If
                
                Call SQLRecordAdd(rsSQL, strSQL)
            End If
        Next
    End With
    
    '��ǰ���
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_������ϼ�¼_DELETE2(" & lngOrderKey & ",8)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(2)
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("����id"))) > 0 Or Val(.TextMatrix(lngRow, .ColIndex("���id"))) > 0 Then
                
                strSQL = "zl_������ϼ�¼_Insert(" & lng����id & "," & ZVal(lng��ҳid) & ",1,Null,8," & Val(.TextMatrix(lngRow, .ColIndex("����id"))) & "," & Val(.TextMatrix(lngRow, .ColIndex("���id"))) & ",Null,'" & .TextMatrix(lngRow, .ColIndex("�������")) & "',Null,Null,Null,Sysdate," & lngOrderKey & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            End If
        Next
    End With
    
    '�������
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "ZL_������ϼ�¼_DELETE2(" & lngOrderKey & ",9)"
    Call SQLRecordAdd(rsSQL, strSQL)
    With vsf(3)
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("����id"))) > 0 Or Val(.TextMatrix(lngRow, .ColIndex("���id"))) > 0 Then
                
                strSQL = "zl_������ϼ�¼_Insert(" & lng����id & "," & ZVal(lng��ҳid) & ",1,Null,9," & Val(.TextMatrix(lngRow, .ColIndex("����id"))) & "," & Val(.TextMatrix(lngRow, .ColIndex("���id"))) & ",Null,'" & .TextMatrix(lngRow, .ColIndex("�������")) & "',Null,Null,Null,Sysdate," & lngOrderKey & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            End If
        Next
    End With

    
    SaveData = True
    
    Exit Function
    
    '������
    '------------------------------------------------------------------------------------------------------------------
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitTabControl()
    '******************************************************************************************************************
    '����:
    '����:
    '����:
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
        
        .InsertItem 0, "�������", picPane(0).hWnd, 0
        .InsertItem 1, "�������", picPane(1).hWnd, 0
        .InsertItem 2, "������", picPane(2).hWnd, 0
        
        .Item(1).Selected = True
        .Item(2).Selected = True
        .Item(0).Selected = True
        
    End With
    
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rs As New ADODB.Recordset
    Dim blnAllowModify As Boolean
    Dim strSQL As String
    
    On Error GoTo errHand
    
    mblnReading = True
    
    Select Case strCommand
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
        
        '������Ա
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfPerson = New clsVsf
        With mclsVsfPerson
            Call .Initialize(Me.Controls, vsf(4), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("��λ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("����", 2100, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���", 900, flexAlignLeftCenter, flexDTString, "", "����", True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        '��������
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfOpsBefore = New clsVsf
        With mclsVsfOpsBefore
            Call .Initialize(Me.Controls, vsf(0), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("���뷽ʽ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("ȱʡ", 810, flexAlignCenterCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        '��������
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfOpsAfter = New clsVsf
        With mclsVsfOpsAfter
            Call .Initialize(Me.Controls, vsf(1), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", IIf(mblnAllowModify, "[ָʾ��]", "[ͼ��]"), False)
            Call .AppendColumn("���뷽ʽ", 1200, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��������", 2400, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("ȱʡ", 810, flexAlignCenterCenter, flexDTBoolean, "", , True)
            Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
                
        '��ǰ���
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfDiagBefore = New clsVsf
        With mclsVsfDiagBefore
            Call .Initialize(Me.Controls, vsf(2), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("��������", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ϱ���", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("�������", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With
        
        '�������
        '--------------------------------------------------------------------------------------------------------------
        Set mclsVsfDiagAfter = New clsVsf
        With mclsVsfDiagAfter
            Call .Initialize(Me.Controls, vsf(3), True, True, frmPubResource.GetImageList(16))
            Call .ClearColumn
            Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
            Call .AppendColumn("��������", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("��ϱ���", 990, flexAlignLeftCenter, flexDTString, "", , True)
            Call .AppendColumn("���id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("����id", 0, flexAlignLeftCenter, flexDTDecimal, "", , True, , , True)
            Call .AppendColumn("�������", 1500, flexAlignLeftCenter, flexDTString, "", , True)
            .AppendRows = True
        End With

        txt(2).BackColor = COLOR.��ɫ
    '--------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
                 
        '����������볤��
        '--------------------------------------------------------------------------------------------------------------
        txt(0).MaxLength = 10

        '����������ģ
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT ����,0 FROM ����������ģ"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then Call AddComboData(cbo(0), rs, False)
                
        '��������
        '--------------------------------------------------------------------------------------------------------------
        With cbo(3)
            .Clear
            .AddItem "1-��"
            .AddItem "2-��"
            .AddItem "3-��"
            .AddItem "4-Σ(��)"
        End With
    
    '------------------------------------------------------------------------------------------------------------------
    Case "�ؼ�״̬"
    
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
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)

                '������λ
                '------------------------------------------------------------------------------------------------------
                gstrSQL = "SELECT ���� FROM ������λ Order by ����"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                Call .InitializeEditColumn(.ColIndex("��λ"), True, vbVsfEditCombox, vsf(4).BuildComboList(rs, "����", "����"))
                
                Call .InitializeEditColumn(.ColIndex("����"), True, vbVsfEditCommand)
                
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
                
        
            End With
        
            With mclsVsfOpsBefore
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("���뷽ʽ"), True, vbVsfEditCombox, "1-����|2-����")
                Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("ȱʡ"), True, vbVsfEditCheck)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            End With
        
            With mclsVsfOpsAfter
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("���뷽ʽ"), True, vbVsfEditCombox, "1-����|2-����")
                Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("ȱʡ"), True, vbVsfEditCheck)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            End With
            
            With mclsVsfDiagBefore
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("��ϱ���"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("�������"), True, vbVsfEditText)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            End With
            
            With mclsVsfDiagAfter
                Call .ModifyColumn(.ColIndex("ͼ��"), "", 255, flexAlignCenterCenter, flexDTString, "", "[ָʾ��]", False)
                Call .InitializeEdit(True, True, True)
                Call .InitializeEditColumn(.ColIndex("��������"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("��ϱ���"), True, vbVsfEditCommand)
                Call .InitializeEditColumn(.ColIndex("�������"), True, vbVsfEditText)
                        
                .IndicatorCol = 0
                Set .IndicatorIcon = frmPubResource.GetImageList(16).ListImages("��ǰ").Picture
            End With
        
        Else
            
            mclsVsfPerson.AllowEdit = False
            mclsVsfOpsBefore.AllowEdit = False
            mclsVsfOpsAfter.AllowEdit = False
            mclsVsfDiagBefore.AllowEdit = False
            mclsVsfDiagAfter.AllowEdit = False
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"
        
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
    Case "��ȡ����"
        
        'ҽ��ִ�з���
        '--------------------------------------------------------------------------------------------------------------
        cbo(1).Clear
        gstrSQL = "SELECT RowNum As ID,ִ�м� As ���� FROM ҽ��ִ�з��� WHERE ����id=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngDeptKey)
        If rs.BOF = False Then Call AddComboData(cbo(1), rs)
        
        '1.��ȡ������������
        '--------------------------------------------------------------------------------------------------------------
        gstrSQL = "SELECT A.*,C.�Ա�,C.��ǰ����id,C.סԺ�� FROM ����������¼ A,������Ϣ C WHERE A.����id=C.����id AND A.ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            
'            mlng����id = zlCommFun.NVL(rs("����id"), 0)
'            mlng��ҳid = zlCommFun.NVL(rs("��ҳid"), 0)
'            mlngDeptKey = zlCommFun.NVL(rs("��ǰ����id"), 0)
            
'            If zlCommFun.NVL(rs("�Ա�")) Like "*��*" Then mstr�Ա� = mstr�Ա� & ",1"
'            If zlCommFun.NVL(rs("�Ա�")) Like "*Ů*" Then mstr�Ա� = mstr�Ա� & ",2"
            
            If IsNull(rs("������ʼʱ��")) = False Then
                dtp(0).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(0).CustomFormat)
                dtp(1).Value = Format(zlCommFun.NVL(rs("��������ʱ��")), dtp(1).CustomFormat)

                
                If IsNull(rs("����ʼʱ��")) = False Then
                    chk(2).Value = 1
                    picConver(2).Visible = False
                    picConver(3).Visible = False
                    dtp(2).Value = Format(zlCommFun.NVL(rs("����ʼʱ��")), dtp(2).CustomFormat)
                    dtp(3).Value = Format(zlCommFun.NVL(rs("�������ʱ��")), dtp(3).CustomFormat)
                Else
                    chk(2).Value = 0
                    picConver(2).Visible = True
                    picConver(3).Visible = True
                    dtp(2).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(2).CustomFormat)
                    dtp(3).Value = Format(zlCommFun.NVL(rs("��������ʱ��")), dtp(3).CustomFormat)
                End If
                
                If IsNull(rs("������ʼʱ��")) = False Then
                    chk(4).Value = 1
                    picConver(4).Visible = False
                    picConver(5).Visible = False
                    dtp(4).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(4).CustomFormat)
                    dtp(5).Value = Format(zlCommFun.NVL(rs("��������ʱ��")), dtp(5).CustomFormat)
                Else
                    chk(4).Value = 0
                    picConver(4).Visible = True
                    picConver(5).Visible = True
                    dtp(4).Value = Format(zlCommFun.NVL(rs("������ʼʱ��")), dtp(4).CustomFormat)
                    dtp(5).Value = Format(zlCommFun.NVL(rs("��������ʱ��")), dtp(5).CustomFormat)
                End If

            End If
            
            zlControl.CboLocate cbo(0), zlCommFun.NVL(rs("������ģ").Value)
            zlControl.CboLocate cbo(1), zlCommFun.NVL(rs("������").Value)
            zlControl.CboLocate cbo(3), zlCommFun.NVL(rs("��������").Value)

            txt(1).Text = zlCommFun.NVL(rs("����ʽ").Value)
            txt(2).Text = zlCommFun.NVL(rs("��������").Value)
            txt(0).Text = zlCommFun.NVL(rs("��Һ����").Value)
        End If
        
        
        '������Ա
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfPerson.ClearGrid
        gstrSQL = "Select A.��Աid As ID,A.��λ,B.��� As ����,B.���� From ����������Ա a,��Ա�� b Where a.��¼id=[1] And a.��Աid=b.ID"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then Call mclsVsfPerson.LoadGrid(rs)
        
        '��������
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfOpsBefore.ClearGrid
        
        gstrSQL = GetPublicSQL(SQL.�����������)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 1)
        If rs.BOF = False Then Call mclsVsfOpsBefore.LoadGrid(rs)
        
        '��������
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfOpsAfter.ClearGrid
        gstrSQL = GetPublicSQL(SQL.�����������)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 2)
        If rs.BOF = False Then Call mclsVsfOpsAfter.LoadGrid(rs)
        
        '�������
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfDiagBefore.ClearGrid
        gstrSQL = GetPublicSQL(SQL.������ϼ�¼)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, 8)
        If rs.BOF = False Then Call mclsVsfDiagBefore.LoadGrid(rs)
        
        '�������
        '--------------------------------------------------------------------------------------------------------------
        mclsVsfDiagAfter.ClearGrid
        gstrSQL = GetPublicSQL(SQL.������ϼ�¼)
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
    Case 1      '����ʽ
        gstrSQL = GetPublicSQL(SQL.����ʽѡ��)
        
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
 
        If ShowPubSelect(Me, txt(1), 2, "����,900,0,;����,2400,0,;��������,900,0,", Me.Name & "\����ʽѡ��", "����±���ѡ��һ������ʽ", rsData, rs, 8790, 4500) = 1 Then
            

            txt(1).Text = zlCommFun.NVL(rs("����").Value)
            txt(2).Text = zlCommFun.NVL(rs("��������").Value)

            txt(1).Tag = ""

            usrSaveItem.����ʽ = txt(1).Text
            
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
            usrSaveItem.����ʽ = ""
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
                strTmp = strText & IIf(ParamInfo.��Ŀ����ƥ�䷽ʽ = 1, "", "%")
                
                gstrSQL = GetPublicSQL(SQL.����ʽ����, bytMode)
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                If ShowPubSelect(Me, txt(Index), 2, "����,990,0,1;����,1500,0,0;��������,900,0,0", Me.Name & "\����ʽ����", "�������ѡ��һ������ʽ", rsData, rs) = 1 Then

                    txt(Index).Text = zlCommFun.NVL(rs("����").Value)
                    txt(2).Text = zlCommFun.NVL(rs("��������").Value)
                    
                    DataChanged = True
                    
                    usrSaveItem.����ʽ = txt(Index).Text

                Else
                    txt(Index).Text = usrSaveItem.����ʽ
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
            txt(Index).Text = usrSaveItem.����ʽ
            txt(Index).Tag = ""
        End If
    End Select
    
End Sub

Private Sub vsf_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

    '�༭����
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
        Case 0, 1                               '���С���������
            
            If Col = .ColIndex("��������") Then

                If Val(Left(.TextMatrix(Row, .ColIndex("���뷽ʽ")), 1)) = 1 Then
                    '���Ʊ���
                    gstrSQL = GetPublicSQL(SQL.������Ŀѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,", Me.Name & "\������Ŀѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                Else
                    '��������
                    gstrSQL = GetPublicSQL(SQL.��������ѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "S")
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\��������ѡ��", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                End If

                If bytRet = 1 Then
                    If Index = 0 Then
                        If mclsVsfOpsBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                    Else
                        If mclsVsfOpsAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                    End If
                    
                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                    .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                    
                    DataChanged = True
                End If
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 2, 3                               '��ǰ�������
            If Col = .ColIndex("��������") Or Col = .ColIndex("��ϱ���") Then
                Select Case Col
                '------------------------------------------------------------------------------------------------------
                Case .ColIndex("��������")
                
                    gstrSQL = GetPublicSQL(SQL.��������ѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, "D")
        
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\��������ѡ��", "����±���ѡ��һ������������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then
                        
                        If Index = 2 Then
                            If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        Else
                            If mclsVsfDiagAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        End If
                        .EditText = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("�������")) = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("����id")) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        gstrSQL = GetPublicSQL(SQL.������϶���)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)
                        
                        DataChanged = True
                    End If
                '----------------------------------------------------------------------------------------------------------
                Case .ColIndex("��ϱ���")
                
                    gstrSQL = GetPublicSQL(SQL.�������ѡ��)
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption)
        
                    bytRet = ShowPubSelect(Me, vsf(Index), 3, "����,1200,0,;����,2700,0,", Me.Name & "\�������ѡ��", "����±���ѡ��һ�����������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    If bytRet = 1 Then
                        If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
            
                        .EditText = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("��ϱ���")) = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("�������")) = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("���id")) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        gstrSQL = GetPublicSQL(SQL.������϶���)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))
                
                        DataChanged = True
                    End If
                End Select
                
                '----------------------------------------------------------------------------------------------------------
                If bytRet = 1 Then
                    If rsData.BOF = False Then
                        .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("��������").Value)
                        .TextMatrix(Row, .ColIndex("��ϱ���")) = zlCommFun.NVL(rs("��ϱ���").Value)
                        .TextMatrix(Row, .ColIndex("����id")) = zlCommFun.NVL(rs("����id").Value, 0)
                        .TextMatrix(Row, .ColIndex("���id")) = zlCommFun.NVL(rs("���id").Value, 0)
                    End If
                End If
                
            End If
        '--------------------------------------------------------------------------------------------------------------
        Case 4                                      '������Ա
            
            If Col = .ColIndex("����") Then
    
                gstrSQL = GetPublicSQL(SQL.��Ա��Ϣѡ��)
                
                strTmp = "ҽ��"
                If InStr(.TextMatrix(.Row, .ColIndex("��λ")), "��ʿ") > 0 Then strTmp = "��ʿ"
                
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strTmp, mlngDeptKey)
    
                If ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1500,0,;����,900,0,;����,1200,0,", Me.Name & "\��Ա��Ϣѡ��", "����±���ѡ��һ����Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then

                    If mclsVsfPerson.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                        ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                        Exit Sub
                    End If
                           
                    .EditText = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                    .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
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
            strTmp = IIf(ParamInfo.��Ŀ����ƥ�䷽ʽ = 1, strText, "%" & strText)
                    
            Select Case Index
            Case 0, 1
                If Col = .ColIndex("��������") Then
                    
                    If Val(Left(.TextMatrix(Row, .ColIndex("���뷽ʽ")), 1)) = 1 Then
                        gstrSQL = GetPublicSQL(SQL.������Ŀ����, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,", Me.Name & "\������Ŀ����", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    Else
                        gstrSQL = GetPublicSQL(SQL.�����������, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "S")
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\�����������", "����±���ѡ��һ��������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                    End If

                    If bytRet = 1 Then
                        
                        If Index = 0 Then
                            If mclsVsfOpsBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        Else
                            If mclsVsfOpsAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                Exit Sub
                            End If
                        End If
    
                        .EditText = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                        
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
                
                If Col = .ColIndex("��������") Or Col = .ColIndex("��ϱ���") Then
                                                
                    Select Case Col
                    '--------------------------------------------------------------------------------------------------
                    Case .ColIndex("��������")
    
                        gstrSQL = GetPublicSQL(SQL.�����������, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp, "D")
        
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,;����,900,0,;����,900,0,", Me.Name & "\�����������", "����±���ѡ��һ������������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        If bytRet = 1 Then
                            
                            If Index = 2 Then
                                If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                    Exit Sub
                                End If
                            Else
                                If mclsVsfDiagAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                    Exit Sub
                                End If
                            End If
        
                            .EditText = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("�������")) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("����id")) = zlCommFun.NVL(rs("ID").Value, 0)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
                            
                            gstrSQL = GetPublicSQL(SQL.������϶���)
                            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, Val(.RowData(Row)), 0)
                        
                            DataChanged = True
                        End If
                    '--------------------------------------------------------------------------------------------------
                    Case .ColIndex("��ϱ���")
                        gstrSQL = GetPublicSQL(SQL.������Ϲ���, bytMode)
                        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strText, strTmp)
        
                        bytRet = ShowPubSelect(Me, vsf(Index), 2, "����,1200,0,;����,2700,0,", Me.Name & "\������Ϲ���", "����±���ѡ��һ�����������Ŀ", rsData, rs, 8790, 4500, , Val(.RowData(Row)))
                        If bytRet = 1 Then
                            
                            If Index = 2 Then
                                If mclsVsfDiagBefore.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                    Exit Sub
                                End If
                            Else
                                If mclsVsfDiagAfter.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                                    Exit Sub
                                End If
                            End If
        
                            .EditText = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("��ϱ���")) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("�������")) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, .ColIndex("���id")) = zlCommFun.NVL(rs("ID").Value, 0)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value, 0)
        
                            gstrSQL = GetPublicSQL(SQL.������϶���)
                            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, 0, Val(.RowData(Row)))
                        
                            DataChanged = True
                        End If
                    End Select
                    
                    If bytRet = 1 Then
                    
                        '--------------------------------------------------------------------------------------------------
                        If rsData.BOF = False Then
                            
                            .TextMatrix(Row, .ColIndex("��������")) = zlCommFun.NVL(rs("��������").Value)
                            .TextMatrix(Row, .ColIndex("��ϱ���")) = zlCommFun.NVL(rs("��ϱ���").Value)
                            
                            .TextMatrix(Row, .ColIndex("����id")) = zlCommFun.NVL(rs("����id").Value, 0)
                            .TextMatrix(Row, .ColIndex("���id")) = zlCommFun.NVL(rs("���id").Value, 0)
                        End If
                
                    Else
                        KeyCode = 0
                        .Cell(flexcpData, Row, Col) = .Cell(flexcpData, Row, Col)
                        .EditText = .Cell(flexcpData, Row, Col)
                        .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                    End If
                    
                End If
            Case 4
                
                If Col = .ColIndex("����") Then
        
                    gstrSQL = GetPublicSQL(SQL.��Ա��Ϣ����, bytMode)
                    
                    strClass = "ҽ��"
                    If InStr(.TextMatrix(.Row, .ColIndex("��λ")), "��ʿ") > 0 Then strClass = "��ʿ"
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, mfrmMain.Caption, strClass, mlngDeptKey, strText, strTmp)
        
                    If ShowPubSelect(Me, vsf(Index), 2, "���,1200,0,;����,1500,0,;����,900,0,;����,1200,0,", Me.Name & "\��Ա��Ϣ����", "����±���ѡ��һ����Ա", rsData, rs, 8790, 4500, , Val(.RowData(Row))) = 1 Then
    
                        If mclsVsfPerson.CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ա��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                               
                        .EditText = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("����")) = zlCommFun.NVL(rs("����").Value)
                        .TextMatrix(Row, .ColIndex("���")) = zlCommFun.NVL(rs("���").Value)
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
    '�༭����
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
    '�༭����
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
    '�༭����
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
    '�༭����
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
    '�༭����
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


