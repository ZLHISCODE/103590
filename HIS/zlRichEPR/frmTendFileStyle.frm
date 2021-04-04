VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileStyle 
   Caption         =   "护理文件样式"
   ClientHeight    =   10275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14955
   Icon            =   "frmTendFileStyle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10275
   ScaleWidth      =   14955
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   300
      Top             =   375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picpaper 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   3240
      ScaleHeight     =   1770
      ScaleWidth      =   11115
      TabIndex        =   91
      Top             =   1305
      Visible         =   0   'False
      Width           =   11115
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   0
         Left            =   7935
         MaxLength       =   6
         TabIndex        =   100
         Text            =   "25"
         Top             =   615
         Width           =   615
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "横向"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   9420
         TabIndex        =   95
         Top             =   135
         Width           =   1095
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "纵向"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   7710
         TabIndex        =   94
         Top             =   135
         Value           =   -1  'True
         Width           =   1080
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   1
         Left            =   9660
         MaxLength       =   6
         TabIndex        =   106
         Text            =   "25"
         Top             =   615
         Width           =   585
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   2
         Left            =   7935
         MaxLength       =   6
         TabIndex        =   116
         Text            =   "19"
         Top             =   1020
         Width           =   600
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   3
         Left            =   9660
         MaxLength       =   6
         TabIndex        =   120
         Text            =   "19"
         Top             =   1020
         Width           =   585
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   110
         Text            =   "210.05"
         Top             =   885
         Width           =   945
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   113
         Text            =   "297.08"
         Top             =   900
         Width           =   945
      End
      Begin VB.ComboBox cboPaperKind 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   420
         Width           =   5355
      End
      Begin MSComCtl2.UpDown udHeight 
         Height          =   300
         Left            =   4560
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   900
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196614
         OrigLeft        =   4170
         OrigTop         =   900
         OrigRight       =   4425
         OrigBottom      =   1185
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udWidth 
         Height          =   300
         Left            =   2235
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   900
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196613
         OrigLeft        =   1830
         OrigTop         =   893
         OrigRight       =   2070
         OrigBottom      =   1178
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   0
         Left            =   8565
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(0)"
         BuddyDispid     =   196611
         BuddyIndex      =   0
         OrigLeft        =   8655
         OrigTop         =   630
         OrigRight       =   8910
         OrigBottom      =   930
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   1
         Left            =   10260
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   615
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(1)"
         BuddyDispid     =   196611
         BuddyIndex      =   1
         OrigLeft        =   10335
         OrigTop         =   630
         OrigRight       =   10590
         OrigBottom      =   930
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   2
         Left            =   8550
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(2)"
         BuddyDispid     =   196611
         BuddyIndex      =   2
         OrigLeft        =   8670
         OrigTop         =   1035
         OrigRight       =   8925
         OrigBottom      =   1335
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   3
         Left            =   10275
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   1020
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(3)"
         BuddyDispid     =   196611
         BuddyIndex      =   3
         OrigLeft        =   10335
         OrigTop         =   1035
         OrigRight       =   10590
         OrigBottom      =   1335
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   5
         Left            =   10590
         TabIndex        =   122
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   4
         Left            =   8880
         TabIndex        =   118
         Top             =   1080
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   3
         Left            =   10575
         TabIndex        =   108
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   2
         Left            =   8865
         TabIndex        =   104
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "上"
         Height          =   180
         Index           =   0
         Left            =   7700
         TabIndex        =   98
         Top             =   675
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "下"
         Height          =   180
         Index           =   1
         Left            =   9420
         TabIndex        =   105
         Top             =   675
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "左"
         Height          =   180
         Index           =   2
         Left            =   7700
         TabIndex        =   115
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "右"
         Height          =   180
         Index           =   3
         Left            =   9420
         TabIndex        =   119
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label lblRound 
         AutoSize        =   -1  'True
         Caption         =   "页边距:"
         Height          =   180
         Left            =   6615
         TabIndex        =   99
         Top             =   690
         Width           =   630
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   1
         Left            =   4890
         TabIndex        =   103
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lblWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "宽"
         Height          =   180
         Left            =   1065
         TabIndex        =   109
         Top             =   945
         Width           =   180
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "高"
         Height          =   180
         Left            =   3375
         TabIndex        =   112
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "毫米"
         Height          =   180
         Index           =   0
         Left            =   2535
         TabIndex        =   102
         Top             =   945
         Width           =   360
      End
      Begin VB.Label lblPaper 
         AutoSize        =   -1  'True
         Caption         =   "纸张种类"
         Height          =   180
         Left            =   270
         TabIndex        =   96
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblPaperHint 
         AutoSize        =   -1  'True
         Caption         =   "注意:  如果实际打印机和当前打印机不符，可能导致纸张设置失效！"
         Height          =   180
         Left            =   270
         TabIndex        =   123
         Top             =   1440
         Width           =   5490
      End
      Begin VB.Label lblOrient 
         AutoSize        =   -1  'True
         Caption         =   "纸张方向:"
         Height          =   180
         Left            =   6615
         TabIndex        =   93
         Top             =   165
         Width           =   810
      End
      Begin VB.Label lblPrinter 
         AutoSize        =   -1  'True
         Caption         =   "打印机:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   270
         TabIndex        =   92
         Top             =   135
         Width           =   690
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8340
      Left            =   375
      ScaleHeight     =   8310
      ScaleWidth      =   11055
      TabIndex        =   1
      Top             =   1215
      Width           =   11085
      Begin VB.PictureBox picSum 
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   3195
         ScaleHeight     =   1875
         ScaleWidth      =   3435
         TabIndex        =   89
         Top             =   4200
         Visible         =   0   'False
         Width           =   3435
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1770
            Left            =   210
            TabIndex        =   90
            Top             =   60
            Width           =   2760
            _cx             =   4868
            _cy             =   3122
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
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
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
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
         Begin VB.Image imgout 
            Height          =   240
            Index           =   4
            Left            =   3165
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.PictureBox picColRelation 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   3795
         Left            =   2895
         ScaleHeight     =   3765
         ScaleWidth      =   3180
         TabIndex        =   85
         Top             =   3405
         Visible         =   0   'False
         Width           =   3210
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   1575
            Picture         =   "frmTendFileStyle.frx":058A
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   86
            Top             =   195
            Width           =   270
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgColRelation 
            Height          =   3075
            Left            =   90
            TabIndex        =   88
            Top             =   600
            Width           =   2970
            _cx             =   5239
            _cy             =   5424
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
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   2
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendFileStyle.frx":0B14
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
         Begin VB.Image imgout 
            Height          =   240
            Index           =   3
            Left            =   2925
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblColRelation 
            Caption         =   "设置汇总对照关系   ："
            Height          =   225
            Left            =   105
            TabIndex        =   87
            Top             =   225
            Width           =   1890
         End
      End
      Begin zlRichEPR.F1ColorPicker ColorForeColor 
         Height          =   2190
         Left            =   4125
         TabIndex        =   79
         Top             =   600
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   3863
         AutoColor       =   16777215
      End
      Begin zlRichEPR.F1ColorPicker ColorFillColor 
         Height          =   2190
         Left            =   4065
         TabIndex        =   80
         Top             =   630
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   3863
         AutoColor       =   16777215
      End
      Begin VB.PictureBox picColDoctor 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   2925
         ScaleHeight     =   3825
         ScaleWidth      =   3285
         TabIndex        =   81
         Top             =   2805
         Visible         =   0   'False
         Width           =   3315
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   2115
            Picture         =   "frmTendFileStyle.frx":0BD0
            ScaleHeight     =   240
            ScaleWidth      =   270
            TabIndex        =   82
            Top             =   195
            Width           =   270
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfColDoctor 
            Height          =   3000
            Left            =   120
            TabIndex        =   84
            Top             =   735
            Width           =   3030
            _cx             =   5345
            _cy             =   5292
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
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   2
            FixedCols       =   2
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendFileStyle.frx":115A
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
         Begin VB.Image imgout 
            Height          =   240
            Index           =   2
            Left            =   3030
            Stretch         =   -1  'True
            Top             =   0
            Width           =   240
         End
         Begin VB.Label lblColDoctor 
            Caption         =   "设置医嘱导入对应列关系   ："
            Height          =   240
            Left            =   135
            TabIndex        =   83
            Top             =   225
            Width           =   2445
         End
      End
      Begin VB.PictureBox picPane 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   7275
         Index           =   1
         Left            =   2655
         ScaleHeight     =   7275
         ScaleWidth      =   8430
         TabIndex        =   40
         Top             =   -270
         Width           =   8430
         Begin VB.PictureBox picCloumn 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            ForeColor       =   &H80000008&
            Height          =   4335
            Left            =   2805
            ScaleHeight     =   4305
            ScaleWidth      =   3300
            TabIndex        =   59
            Top             =   2925
            Visible         =   0   'False
            Width           =   3330
            Begin VB.CommandButton cmdOk 
               Height          =   315
               Left            =   2580
               Picture         =   "frmTendFileStyle.frx":11EA
               Style           =   1  'Graphical
               TabIndex        =   71
               ToolTipText     =   "确认"
               Top             =   3885
               Width           =   495
            End
            Begin VB.ComboBox cboItemSearch 
               Height          =   300
               Left            =   810
               TabIndex        =   64
               Top             =   450
               Width           =   1785
            End
            Begin VB.ListBox lstColumnUsed 
               Appearance      =   0  'Flat
               Height          =   1650
               ItemData        =   "frmTendFileStyle.frx":1774
               Left            =   75
               List            =   "frmTendFileStyle.frx":1776
               TabIndex        =   65
               Top             =   870
               Width           =   2625
            End
            Begin VB.TextBox txtColumnPostfix 
               Enabled         =   0   'False
               Height          =   300
               Left            =   510
               TabIndex        =   69
               Top             =   3075
               Width           =   2175
            End
            Begin VB.TextBox txtColumnPrefix 
               Enabled         =   0   'False
               Height          =   300
               Left            =   510
               TabIndex        =   67
               Top             =   2700
               Width           =   2175
            End
            Begin VB.CheckBox chk 
               Caption         =   "对角线"
               Height          =   210
               Left            =   495
               TabIndex        =   70
               Top             =   3540
               Width           =   1020
            End
            Begin VB.TextBox txtColumnNo 
               Enabled         =   0   'False
               Height          =   300
               Left            =   315
               MaxLength       =   2
               TabIndex        =   60
               Text            =   "1"
               Top             =   75
               Width           =   390
            End
            Begin MSComCtl2.UpDown udColumnNo 
               Height          =   300
               Left            =   705
               TabIndex        =   61
               Top             =   75
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txtColumnNo"
               BuddyDispid     =   196649
               OrigLeft        =   5400
               OrigTop         =   75
               OrigRight       =   5640
               OrigBottom      =   375
               Max             =   5
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.Label lblSearch 
               AutoSize        =   -1  'True
               Caption         =   "查找："
               Height          =   180
               Index           =   1
               Left            =   165
               TabIndex        =   63
               Top             =   510
               Width           =   540
            End
            Begin VB.Image imgout 
               Height          =   240
               Index           =   0
               Left            =   3030
               Stretch         =   -1  'True
               Top             =   0
               Width           =   240
            End
            Begin VB.Label lblColumnPrefix 
               AutoSize        =   -1  'True
               Caption         =   "前缀"
               Height          =   180
               Left            =   75
               TabIndex        =   66
               Top             =   2745
               Width           =   330
            End
            Begin VB.Label lblColumnPostfix 
               AutoSize        =   -1  'True
               Caption         =   "后缀"
               Height          =   180
               Left            =   75
               TabIndex        =   68
               Top             =   3120
               Width           =   330
            End
            Begin VB.Label lblColumnNo 
               AutoSize        =   -1  'True
               Caption         =   "第        列内容项目:"
               Height          =   180
               Left            =   105
               TabIndex        =   62
               Top             =   135
               Width           =   1890
            End
            Begin VB.Image imgColDelete 
               Height          =   240
               Left            =   2745
               Top             =   885
               Width           =   240
            End
         End
         Begin VB.PictureBox picLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            ForeColor       =   &H80000008&
            Height          =   4035
            Left            =   4215
            ScaleHeight     =   4005
            ScaleWidth      =   3180
            TabIndex        =   49
            Top             =   2880
            Visible         =   0   'False
            Width           =   3210
            Begin VB.CommandButton cmdLabOK 
               Height          =   315
               Left            =   2445
               Picture         =   "frmTendFileStyle.frx":1778
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "确认"
               Top             =   3600
               Width           =   510
            End
            Begin VB.PictureBox picDown 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   2490
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   54
               Top             =   1995
               Width           =   480
            End
            Begin VB.PictureBox picUP 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   420
               Left            =   2505
               ScaleHeight     =   420
               ScaleWidth      =   315
               TabIndex        =   53
               Top             =   870
               Width           =   315
            End
            Begin VB.TextBox txtLabelPrefix 
               Enabled         =   0   'False
               Height          =   300
               Left            =   90
               TabIndex        =   56
               Top             =   3345
               Width           =   975
            End
            Begin VB.CheckBox chkLabelCrLf 
               Caption         =   "换行"
               Height          =   180
               Left            =   90
               TabIndex        =   55
               Top             =   3045
               Width           =   870
            End
            Begin VB.ListBox lstLabelUsed 
               Appearance      =   0  'Flat
               Height          =   2370
               Left            =   60
               TabIndex        =   52
               Top             =   495
               Width           =   1920
            End
            Begin VB.ComboBox cboLableSearch 
               Height          =   300
               Left            =   720
               TabIndex        =   51
               Top             =   75
               Width           =   1920
            End
            Begin VB.Image imgout 
               Height          =   240
               Index           =   1
               Left            =   2940
               Stretch         =   -1  'True
               Top             =   0
               Width           =   240
            End
            Begin VB.Image imgLabDelete 
               Height          =   240
               Left            =   1980
               Top             =   480
               Width           =   240
            End
            Begin VB.Label lblLabelPrefix 
               AutoSize        =   -1  'True
               Caption         =   "前缀文本"
               Height          =   180
               Left            =   1230
               TabIndex        =   57
               Top             =   3390
               Width           =   720
            End
            Begin VB.Label lblSearch 
               AutoSize        =   -1  'True
               Caption         =   "查找："
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   50
               Top             =   135
               Width           =   540
            End
         End
         Begin VB.TextBox txtTitleText 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   2430
            TabIndex        =   45
            Text            =   "记录单"
            Top             =   1635
            Width           =   4425
         End
         Begin VB.PictureBox picFoot页码 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1200
            Left            =   0
            ScaleHeight     =   1170
            ScaleWidth      =   7830
            TabIndex        =   72
            Top             =   3930
            Width           =   7860
            Begin VB.OptionButton optPageAlign 
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   2775
               Picture         =   "frmTendFileStyle.frx":1D02
               Style           =   1  'Graphical
               TabIndex        =   76
               TabStop         =   0   'False
               Top             =   435
               Width           =   345
            End
            Begin VB.OptionButton optPageAlign 
               Enabled         =   0   'False
               Height          =   315
               Index           =   0
               Left            =   2085
               Picture         =   "frmTendFileStyle.frx":205B
               Style           =   1  'Graphical
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   435
               Width           =   345
            End
            Begin VB.ComboBox cbo页码 
               Enabled         =   0   'False
               Height          =   300
               Left            =   4185
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   450
               Width           =   1575
            End
            Begin VB.CheckBox chk页码 
               Caption         =   "打印页码"
               Height          =   195
               Left            =   840
               TabIndex        =   73
               Top             =   495
               Width           =   1155
            End
            Begin VB.OptionButton optPageAlign 
               Enabled         =   0   'False
               Height          =   315
               Index           =   1
               Left            =   2430
               Picture         =   "frmTendFileStyle.frx":23E1
               Style           =   1  'Graphical
               TabIndex        =   75
               TabStop         =   0   'False
               Top             =   435
               Width           =   345
            End
            Begin VB.Label lbl页码 
               AutoSize        =   -1  'True
               Caption         =   "页码位置"
               Enabled         =   0   'False
               Height          =   180
               Left            =   3375
               TabIndex        =   78
               Top             =   480
               Width           =   975
            End
         End
         Begin VB.PictureBox picFoot 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1400
            Left            =   0
            ScaleHeight     =   1481.395
            ScaleMode       =   0  'User
            ScaleWidth      =   1365
            TabIndex        =   46
            Top             =   2355
            Width           =   1400
            Begin VB.Label lblFoot 
               AutoSize        =   -1  'True
               Caption         =   "页脚"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   420
               TabIndex        =   47
               Top             =   570
               Width           =   480
            End
         End
         Begin VB.PictureBox picHead 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   1400
            Left            =   0
            ScaleHeight     =   1400
            ScaleMode       =   0  'User
            ScaleWidth      =   1365
            TabIndex        =   42
            Top             =   285
            Width           =   1400
            Begin VB.Label lblHead 
               AutoSize        =   -1  'True
               Caption         =   "页眉"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   420
               TabIndex        =   43
               Top             =   525
               Width           =   480
            End
         End
         Begin RichTextLib.RichTextBox rtbHead 
            Height          =   1200
            Left            =   1170
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   0
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   2117
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   2
            OLEDragMode     =   0
            OLEDropMode     =   0
            TextRTF         =   $"frmTendFileStyle.frx":2771
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin RichTextLib.RichTextBox rtbFoot 
            Height          =   1200
            Left            =   1185
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   2625
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   2117
            _Version        =   393217
            BorderStyle     =   0
            Enabled         =   -1  'True
            ScrollBars      =   2
            OLEDragMode     =   0
            OLEDropMode     =   0
            TextRTF         =   $"frmTendFileStyle.frx":280E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgThis 
            Height          =   1425
            Left            =   15
            TabIndex        =   44
            Top             =   1155
            Width           =   7830
            _cx             =   13811
            _cy             =   2514
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
            BackColorFixed  =   8421504
            ForeColorFixed  =   12632256
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483644
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendFileStyle.frx":28AB
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   1
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   0   'False
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   2
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
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
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picPane 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   6990
         Index           =   0
         Left            =   0
         ScaleHeight     =   6960
         ScaleWidth      =   2565
         TabIndex        =   2
         Top             =   0
         Width           =   2595
         Begin VB.PictureBox picBase 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            ForeColor       =   &H80000008&
            Height          =   2520
            Index           =   0
            Left            =   0
            ScaleHeight     =   2490
            ScaleWidth      =   2535
            TabIndex        =   5
            Top             =   480
            Width           =   2565
            Begin VB.TextBox txtTabCols 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1425
               MaxLength       =   2
               TabIndex        =   8
               Text            =   "5"
               Top             =   330
               Width           =   360
            End
            Begin VB.TextBox txtTabRowHeight 
               Height          =   300
               Left            =   1410
               MaxLength       =   3
               TabIndex        =   11
               Text            =   "300"
               Top             =   690
               Width           =   510
            End
            Begin VB.OptionButton optTabTiers 
               Caption         =   "三(&3)"
               Height          =   180
               Index           =   2
               Left            =   1665
               TabIndex        =   15
               Top             =   1515
               Width           =   780
            End
            Begin VB.OptionButton optTabTiers 
               Caption         =   "双(&2)"
               Height          =   180
               Index           =   1
               Left            =   870
               TabIndex        =   14
               Top             =   1515
               Width           =   780
            End
            Begin VB.OptionButton optTabTiers 
               Caption         =   "单(&1)"
               Height          =   180
               Index           =   0
               Left            =   75
               TabIndex        =   13
               Top             =   1515
               Value           =   -1  'True
               Width           =   780
            End
            Begin VB.CheckBox chk首列合并 
               Caption         =   "日期时间列合并打印"
               Height          =   195
               Left            =   105
               TabIndex        =   16
               Top             =   1875
               Width           =   1980
            End
            Begin VB.CheckBox chkHideTime 
               Caption         =   "时间列隐藏打印"
               Height          =   195
               Left            =   105
               TabIndex        =   17
               Top             =   2175
               Width           =   1635
            End
            Begin MSComCtl2.UpDown udTabCols 
               Height          =   300
               Left            =   1785
               TabIndex        =   9
               Top             =   330
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   5
               BuddyControl    =   "txtTabCols"
               BuddyDispid     =   196680
               OrigLeft        =   1725
               OrigTop         =   315
               OrigRight       =   1965
               OrigBottom      =   600
               Max             =   60
               Min             =   3
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.Label lblBasic 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "基本形态："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   45
               TabIndex        =   6
               Top             =   45
               Width           =   975
            End
            Begin VB.Label lblTabCols 
               AutoSize        =   -1  'True
               Caption         =   "表格列数:"
               Height          =   180
               Left            =   75
               TabIndex        =   7
               Top             =   390
               Width           =   810
            End
            Begin VB.Label lblTabRowHeight 
               AutoSize        =   -1  'True
               Caption         =   "最小行高:"
               Height          =   180
               Left            =   75
               TabIndex        =   10
               Top             =   750
               Width           =   810
            End
            Begin VB.Label lblTabTiers 
               AutoSize        =   -1  'True
               Caption         =   "表头层数:"
               Height          =   180
               Left            =   90
               TabIndex        =   12
               Top             =   1140
               Width           =   810
            End
         End
         Begin VB.PictureBox picBase 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            ForeColor       =   &H80000008&
            Height          =   1845
            Index           =   1
            Left            =   0
            ScaleHeight     =   1815
            ScaleWidth      =   2535
            TabIndex        =   20
            Top             =   3480
            Width           =   2565
            Begin VB.TextBox txtHeadText 
               Height          =   765
               Left            =   -15
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   29
               Top             =   1050
               Width           =   2805
            End
            Begin VB.TextBox txtHeadCol 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1665
               MaxLength       =   2
               TabIndex        =   27
               Text            =   "1"
               Top             =   315
               Width           =   330
            End
            Begin VB.TextBox txtHeadRow 
               Enabled         =   0   'False
               Height          =   300
               Left            =   465
               MaxLength       =   1
               TabIndex        =   23
               Text            =   "2"
               Top             =   315
               Width           =   330
            End
            Begin MSComCtl2.UpDown udHeadCol 
               Height          =   300
               Left            =   1995
               TabIndex        =   26
               Top             =   315
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   1
               BuddyControl    =   "txtHeadCol"
               BuddyDispid     =   196691
               OrigLeft        =   5985
               OrigTop         =   2085
               OrigRight       =   6225
               OrigBottom      =   2370
               Max             =   5
               Min             =   1
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown udHeadRow 
               Height          =   285
               Left            =   765
               TabIndex        =   24
               Top             =   315
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               Value           =   2
               BuddyControl    =   "txtHeadRow"
               BuddyDispid     =   196692
               OrigLeft        =   4920
               OrigTop         =   2085
               OrigRight       =   5160
               OrigBottom      =   2385
               Max             =   3
               Min             =   2
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.Label lblHeadSet 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "表头单元："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   15
               TabIndex        =   21
               Top             =   45
               Width           =   975
            End
            Begin VB.Label lblHeadText 
               AutoSize        =   -1  'True
               Caption         =   "文本"
               Height          =   180
               Left            =   60
               TabIndex        =   28
               Top             =   765
               Width           =   360
            End
            Begin VB.Label lblHeadCol 
               AutoSize        =   -1  'True
               Caption         =   "列号"
               Height          =   180
               Left            =   1245
               TabIndex        =   25
               Top             =   375
               Width           =   360
            End
            Begin VB.Label lblHeadRow 
               AutoSize        =   -1  'True
               Caption         =   "层号"
               Height          =   180
               Left            =   60
               TabIndex        =   22
               Top             =   375
               Width           =   360
            End
         End
         Begin VB.PictureBox picBase 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            ForeColor       =   &H80000008&
            Height          =   1140
            Index           =   2
            Left            =   15
            ScaleHeight     =   1110
            ScaleWidth      =   2535
            TabIndex        =   32
            Top             =   5820
            Width           =   2565
            Begin VB.TextBox txtRecordTo 
               Enabled         =   0   'False
               Height          =   300
               Left            =   1620
               MaxLength       =   2
               TabIndex        =   38
               Text            =   "8"
               Top             =   390
               Width           =   345
            End
            Begin VB.TextBox txtRecordFrom 
               Enabled         =   0   'False
               Height          =   300
               Left            =   240
               MaxLength       =   2
               TabIndex        =   34
               Text            =   "18"
               Top             =   390
               Width           =   360
            End
            Begin MSComCtl2.UpDown udRecordTo 
               Height          =   300
               Left            =   1950
               TabIndex        =   37
               Top             =   390
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   8
               BuddyControl    =   "txtRecordTo"
               BuddyDispid     =   196699
               OrigLeft        =   5985
               OrigTop         =   405
               OrigRight       =   6225
               OrigBottom      =   705
               Max             =   23
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown udRecordFrom 
               Height          =   300
               Left            =   600
               TabIndex        =   35
               Top             =   390
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   18
               BuddyControl    =   "txtRecordFrom"
               BuddyDispid     =   196700
               OrigLeft        =   4440
               OrigTop         =   405
               OrigRight       =   4680
               OrigBottom      =   705
               Max             =   23
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.Label lblRecordStyle 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "条件样式："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   15
               TabIndex        =   33
               Top             =   90
               Width           =   975
            End
            Begin VB.Label lblRecordTo 
               AutoSize        =   -1  'True
               Caption         =   "次日       点"
               Height          =   180
               Left            =   1230
               TabIndex        =   39
               Top             =   450
               Width           =   1170
            End
            Begin VB.Label lblRecordFrom 
               AutoSize        =   -1  'True
               Caption         =   "从       点至"
               Height          =   180
               Left            =   45
               TabIndex        =   36
               Top             =   450
               Width           =   1170
            End
         End
         Begin VB.PictureBox picSize 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   0
            Left            =   15
            Picture         =   "frmTendFileStyle.frx":2993
            ScaleHeight     =   420
            ScaleWidth      =   2505
            TabIndex        =   3
            Top             =   15
            Width           =   2535
            Begin VB.Image ImgUpdown 
               Height          =   360
               Index           =   0
               Left            =   2115
               Top             =   15
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "基本形态："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   90
               TabIndex        =   4
               Top             =   105
               Width           =   975
            End
         End
         Begin VB.PictureBox picSize 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   1
            Left            =   15
            Picture         =   "frmTendFileStyle.frx":6281
            ScaleHeight     =   420
            ScaleWidth      =   2505
            TabIndex        =   18
            Top             =   3015
            Width           =   2535
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "表头单元："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   1
               Left            =   90
               TabIndex        =   19
               Top             =   105
               Width           =   975
            End
            Begin VB.Image ImgUpdown 
               Height          =   360
               Index           =   1
               Left            =   2145
               Top             =   30
               Width           =   360
            End
         End
         Begin VB.PictureBox picSize 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   450
            Index           =   2
            Left            =   15
            Picture         =   "frmTendFileStyle.frx":9B6F
            ScaleHeight     =   420
            ScaleWidth      =   2505
            TabIndex        =   30
            Top             =   5355
            Width           =   2535
            Begin VB.Image ImgUpdown 
               Height          =   360
               Index           =   2
               Left            =   2160
               Top             =   45
               Width           =   360
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               Caption         =   "条件样式："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   2
               Left            =   90
               TabIndex        =   31
               Top             =   105
               Width           =   975
            End
         End
      End
   End
   Begin VB.CheckBox chk重算数据行 
      Caption         =   "重算数据行"
      Height          =   195
      Left            =   6555
      TabIndex        =   0
      Top             =   690
      Width           =   1395
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   124
      Top             =   9915
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendFileStyle.frx":D45D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   22490
            Text            =   "可以根据医院实际情况，设置调整护理记录的查看与输出样式。"
            TextSave        =   "可以根据医院实际情况，设置调整护理记录的查看与输出样式。"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "编辑"
            TextSave        =   "编辑"
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
   Begin MSComctlLib.ImageList Img 
      Left            =   5655
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendFileStyle.frx":DCF1
            Key             =   "out"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendFileStyle.frx":14553
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendFileStyle.frx":1ADB5
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendFileStyle.frx":1B34F
            Key             =   "down"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendFileStyle.frx":1B7D0
            Key             =   "up"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendFileStyle.frx":1BC50
            Key             =   "BaseDown"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendFileStyle.frx":1C34A
            Key             =   "BaseUp"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager ImageManager 
      Left            =   3960
      Top             =   465
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmTendFileStyle.frx":1CA44
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1065
      Top             =   525
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmTendFileStyle.frx":73C14
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendFileStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conRowHeight = 300        '基准行高度

'打印纸张常量(256=自定义)
Const PageSize1 = "信笺， 8 1/2×11 英寸"
Const PageSize2 = "+A611 小型信笺， 8 1/2×11 英寸"
Const PageSize3 = "小型报， 11×17 英寸"
Const PageSize4 = "分类帐， 17×11 英寸"
Const PageSize5 = "法律文件， 8 1/2×14 英寸"
Const PageSize6 = "声明书，5 1/2×8 1/2 英寸"
Const PageSize7 = "行政文件，7 1/2×10 1/2 英寸"
Const PageSize8 = "A3, 297×420 毫米"
Const PageSize9 = "A4, 210×297 毫米"
Const PageSize10 = "A4小号， 210×297 毫米"
Const PageSize11 = "A5, 148×210 毫米"
Const PageSize12 = "B4, 250×354 毫米"
Const PageSize13 = "B5, 182×257 毫米"
Const PageSize14 = "对开本， 8 1/2×13 英寸"
Const PageSize15 = "四开本， 215×275 毫米"
Const PageSize16 = "10×14 英寸"
Const PageSize17 = "11×17 英寸"
Const PageSize18 = "便条，8 1/2×11 英寸"
Const PageSize19 = "#9 信封， 3 7/8×8 7/8 英寸"
Const PageSize20 = "#10 信封， 4 1/8×9 1/2 英寸"
Const PageSize21 = "#11 信封， 4 1/2×10 3/8 英寸"
Const PageSize22 = "#12 信封， 4 1/2×11 英寸"
Const PageSize23 = "#14 信封， 5×11 1/2 英寸"
Const PageSize24 = "C 尺寸工作单"
Const PageSize25 = "D 尺寸工作单"
Const PageSize26 = "E 尺寸工作单"
Const PageSize27 = "DL 型信封， 110×220 毫米"
Const PageSize28 = "C5 型信封， 162×229 毫米"
Const PageSize29 = "C3 型信封， 324×458 毫米"
Const PageSize30 = "C4 型信封， 229×324 毫米"
Const PageSize31 = "C6 型信封， 114×162 毫米"
Const PageSize32 = "C65 型信封，114×229 毫米"
Const PageSize33 = "B4 型信封， 250×353 毫米"
Const PageSize34 = "B5 型信封，176×250 毫米"
Const PageSize35 = "B6 型信封， 176×125 毫米"
Const PageSize36 = "信封， 110×230 毫米"
Const PageSize37 = "信封大王， 3 7/8×7 1/2 英寸"
Const PageSize38 = "信封， 3 5/8×6 1/2 英寸"
Const PageSize39 = "U.S. 标准复写簿， 14 7/8×11 英寸"
Const PageSize40 = "德国标准复写簿， 8 1/2×12 英寸"
Const PageSize41 = "德国法律复写簿， 8 1/2×13 英寸"

Private Const RGN_DIFF = 4

Private mlngFileID As Long          '被编辑的记录ID，修改、查阅时由上级程序通过ShowMe传递进入,新增时为0，
Private mblnRTBFoot As Boolean
Private mblnOk As Boolean           '是否完成编辑退出
Private mlngPageRow As Long         '当前页面格式定义的有效数据行
Private mintType As Integer         '当前选中的空间类型
Private mintBlnReCulat As Boolean   '是否需要重算

Private Enum TYPE_Type
    Type_表头文本 = 0
    Type_表格文本 = 1
    Type_条件文本 = 2
    Type_页眉页脚 = 3
End Enum

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

'页眉页脚相关
'######################################################################################################
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'矩形
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'包含用于格式化指定设备的相关信息
Private Type FORMATRANGE
    hDC As Long             '渲染设备
    hdcTarget As Long       '目标设备
    rc As RECT              '渲染区域，单位：缇。
    rcPage As RECT          '渲染设备的整体区域，单位：缇。
    chrg As CHARRANGE       '用于格式化的文本范围。
End Type

Private Type PageInfo
    PageNumber As Long      '页码
    Start As Long           '字符起始位置
    End As Long             '字符终止位置
    ActualHeight As Long    '本页实际打印高度
End Type
Private AllPages() As PageInfo   '页信息
Private Const WM_PASTE = &H302&              '粘贴
Private Const WM_USER = &H400                '通常用 WM_USER + X 来自定义消息
Private Const EM_FORMATRANGE = (WM_USER + 57)    '为某一设备格式化指定范围的文本。
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '设置用于所见即所得的目标设备和行宽。
Private Const EM_HIDESELECTION = (WM_USER + 63)  '显示/隐藏文本。
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_FINDSTRINGEXACT = &H158  '在ComboBox中精确查找
Private Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Private Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '获取中英文混合字符串长度
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
'######################################################################################################

'临时变量
Private rsItems As New ADODB.Recordset   '绑定项目数据集
Private rsLabItems As New ADODB.Recordset   '表上标签数据集
Private rsTemp As New ADODB.Recordset
Private lngCount As Long
Private strTemp As String
Private strCurFont As String
Private objFont As StdFont
Private mblnChanged As Boolean
Private rtbThis As Object
Private mstrColDoctor As String     '项目名称;总给予量|项目名称;医嘱内容 (绑定医嘱列使用)
Private mstrTextFont  As String     '表格文本字体
Private mlngTabTextColor As Long    '表格文本颜色
Private mlngTabGridColor As Long    '表格颜色
Private mstrTitleFont As String     '标题字体
Private mstrRecordFont As String    '条件字体
Private mlngRecordColor As Long     '条件颜色
Private mintTableAlign As Integer   '表列项目对齐方式


'项目性质：录入项，选择，单选，多选
'一列一个项目的，程序不控制
'一列有两个项目的，项目之间必须加前缀或后缀加以识别且项目性质必须相同，只能是录入项或选择项
'一列有两个以上项目的，只能是录入项
'规则修改:如果列头文本含/且只绑定两个项目,项目的分隔符也是/,则类型必须相同,只能是录入或选择项;否则不控制
'
'日期提供一种格式：日/月，如8/6


Private Property Let DataChanged(vData As Boolean)
    
    mblnChanged = vData
        
    If mblnChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
    
End Property

Private Property Get DataChanged() As Boolean
    
    DataChanged = mblnChanged

End Property

Private Function GetPaperName(ByVal intSize As Integer) As String
    '功能： 根据当前打印机的设置，获取纸张名称
    '返回： 纸张名称
    If intSize = 256 Then
        GetPaperName = "用户自定义 ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
            intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "不可测的纸张 ..."
    End If
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenu As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrToolBar As CommandBar
    Dim Combo As CommandBarComboBox                     '工具栏下拉框控件
    Dim objCustControl As CommandBarControlCustom
    Dim i As Long
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.Options.LargeIcons = False
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = ImageManager.Icons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    
    
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set imgout(0) = Img.ListImages("out").Picture
    Set imgout(1) = Img.ListImages("out").Picture
    Set imgout(2) = Img.ListImages("out").Picture
    Set imgout(3) = Img.ListImages("out").Picture
    Set imgout(4) = Img.ListImages("out").Picture
    Set imgLabDelete = Img.ListImages("delete").Picture
    Set imgColDelete = Img.ListImages("delete").Picture
    Set picUP = Img.ListImages("up").Picture
    Set picDown = Img.ListImages("down").Picture
    Set ImgUpdown(0) = Img.ListImages("BaseDown").Picture
    Set ImgUpdown(1) = Img.ListImages("BaseDown").Picture
    Set ImgUpdown(2) = Img.ListImages("BaseDown").Picture
    Call ShapeMe(RGB(255, 255, 255), True, picUP)
    Call ShapeMe(RGB(255, 255, 255), True, picDown)
     '快键绑定
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.SetIconSize 24, 24
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, "保存并退出"): cbrControl.BeginGroup = True
        cbrControl.STYLE = xtpButtonIconAndCaption '同时显示图标和文字
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存"): cbrControl.ToolTipText = "保存已更改的数据(Ctrl+S,F2)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "恢复"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "恢复到上次保存时的数据状态"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "预览当前设置的记录单样式"
        Set cbrMenu = .Add(xtpControlPopup, conMenu_ManagePopup, "管理"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "管理同步、汇总和重算": cbrMenu.IconId = conMenu_ManagePopup
        With cbrMenu.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_MarkMap, "插入图片")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SummaryTime, "汇总时段")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewCompute, "护理文件重算")
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PageSYN, "页面格式同步")
        End With
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "帮助(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.ToolTipText = "退出当前的设计窗体(Esc)"

    End With
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
        
    Set cbrToolBar = cbsThis.Add("工具栏", xtpBarTop)
    cbrToolBar.SetIconSize 16, 16
    cbrToolBar.ShowTextBelowIcons = False '工具栏中的按钮文字显示在图标右侧
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
'        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置")
'        cbrControl.STYLE = xtpButtonIconAndCaption '同时显示图标和文字
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Corresponding, "汇总对照设置")
        cbrControl.STYLE = xtpButtonIconAndCaption '同时显示图标和文字
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Orders, "医嘱对照设置")
        cbrControl.BeginGroup = True
        cbrControl.STYLE = xtpButtonIconAndCaption '同时显示图标和文字
        
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ID_FORMAT_FORECOLOR, "字体颜色")
        cbrControl.BeginGroup = True
        Set objCustControl = cbrControl.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorForeColor.hWnd
        
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ID_DRAW_LINECOLOR, "表格颜色")
        cbrControl.BeginGroup = True
        Set objCustControl = cbrControl.CommandBar.Controls.Add(xtpControlCustom, 0, "")
        objCustControl.Handle = ColorFillColor.hWnd
        
        Set cbrControl = .Add(xtpControlButton, 31, "字体")
        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTNAME, "字体名称")
        Combo.BeginGroup = True
        For i = 0 To gfrmPublic.cmbFont.ListCount - 1
            Combo.AddItem gfrmPublic.cmbFont.List(i), i + 1
            If gfrmPublic.cmbFont.List(i) = "宋体" Then Combo.ListIndex = i + 1
        Next
        Combo.Width = 90
        Combo.DropDownWidth = 250
        Combo.DropDownListStyle = True

        Set Combo = .Add(xtpControlComboBox, ID_FORMAT_FONTSIZE, "字体尺寸")
        '字号列表
        Combo.AddItem "初号"
        Combo.AddItem "小初"
        Combo.AddItem "一号"
        Combo.AddItem "小一"
        Combo.AddItem "二号"
        Combo.AddItem "小二"
        Combo.AddItem "三号"
        Combo.AddItem "小三"
        Combo.AddItem "四号"
        Combo.AddItem "小四"
        Combo.AddItem "五号"
        Combo.AddItem "小五"
        Combo.AddItem "六号"
        Combo.AddItem "小六"
        Combo.AddItem "七号"
        Combo.AddItem "八号"
        Combo.AddItem 5
        Combo.AddItem 5.5
        Combo.AddItem 6.5
        Combo.AddItem 7.5
        Combo.AddItem 8
        Combo.AddItem 9
        Combo.AddItem 10
        Combo.AddItem 10.5
        Combo.AddItem 11
        Combo.AddItem 12
        Combo.AddItem 14
        Combo.AddItem 16
        Combo.AddItem 18
        Combo.AddItem 20
        Combo.AddItem 22
        Combo.AddItem 24
        Combo.AddItem 26
        Combo.AddItem 28
        Combo.AddItem 36
        Combo.AddItem 48
        Combo.AddItem 72
        Combo.ListIndex = 10
        Combo.Width = 50
        Combo.DropDownWidth = 80
        Combo.DropDownListStyle = True
        
        Set cbrCustom = cbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
        chk重算数据行.Visible = True
        cbrCustom.Handle = chk重算数据行.hWnd
        
    End With
    
    Set cbrToolBar = cbsThis.Add("表格", xtpBarTop): cbrToolBar.BarId = ID_Main_TABLE
    cbrToolBar.SetIconSize 16, 16
    cbrToolBar.EnableDocking xtpFlagHideWrap
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlLabel, 32, "表格：")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_AddUP, "增加行(在上方)(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_AddBottom, "增加行(在下方)(&B)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_DeleteRow, "删除行(&R)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_AddLeft, "增加列(在左侧)(&L)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_AddRight, "增加列(在右侧)(&R")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_DeleteCol, "删除列(&C)"): cbrControl.BeginGroup = True
        
        .Add xtpControlButton, conMenu_Edit_BlodFont, "粗体"
        .Add xtpControlButton, conMenu_Edit_Ttalic, "斜体"
        .Add xtpControlButton, conMenu_Edit_Underline, "下划线"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Left_Alignment, "左对齐"): cbrControl.BeginGroup = True
        .Add xtpControlButton, conMenu_Edit_Center_Alignment, "居中"
        .Add xtpControlButton, conMenu_Edit_Right_Alignment, "右对齐"
        Set cbrControl = .Add(xtpControlLabel, ID_TABLE_FORMATCOLWIDTH, "W：")
        Set cbrControl = .Add(xtpControlEdit, ID_TABLE_FORMATCOLWIDTH, "")
        Set cbrControl = .Add(xtpControlLabel, ID_TABLE_FORMATROWHEIGHT, "H：")
        Set cbrControl = .Add(xtpControlEdit, ID_TABLE_FORMATROWHEIGHT, "")
    End With
        
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel And cbrControl.Type <> xtpControlEdit Then
            cbrControl.STYLE = xtpButtonIcon
        End If
    Next
    
     '快键绑定
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
        
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Public Sub ShapeMe(COLOR As Long, HorizontalScan As Boolean, Optional PicOne As PictureBox, Optional PicTwo As PictureBox = Nothing)
'========================================
'函数名称：ShapeMe
'函数作用：透明化PictureBox背景
'参数(Color透明颜色,横纵向,透明化控件)
'========================================
    
    Dim intX As Integer, intY As Integer
    Dim dblHeight As Double, dblWidth As Double
    Dim lngHDc As Long
    Dim booMiddleOfSet As Boolean
    Dim colPoints As Collection
    Set colPoints = New Collection
    Dim Z As Variant
    Dim dblTransY As Double
    Dim dblTransStartX As Double
    Dim dblTransEndX As Double
    Dim Name As Object
    Dim CurRgn As Long, TempRgn As Long
    If Not PicOne Is Nothing Then
       Set Name = PicOne
    Else
       Set Name = PicTwo
    End If
    
   With Name
        .AutoRedraw = True 'object must have this setting
       .ScaleMode = 3 'object must have this setting
       lngHDc = .hDC 'faster to use a variable; VB help recommends using the property, but I didn't encounter any problems
       If HorizontalScan = True Then 'look for lines of transparency horizontally
           dblHeight = .ScaleHeight 'faster to use a variable
           dblWidth = .ScaleWidth 'faster to use a variable
       Else 'look vertically (note that the names "dblHeight" and "dblWidth" are non-sensical now, but this was an easy way to do this
           dblHeight = .ScaleWidth 'faster to use a variable
           dblWidth = .ScaleHeight 'faster to use a variable
       End If 'HorizontalScan = True
   End With
    booMiddleOfSet = False
   
    'gather all points that need to be made transparent
   For intY = 0 To dblHeight  ' Go through each column of pixels on form
       dblTransY = intY
        For intX = 0 To dblWidth ' Go through each line of pixels on form
           'note that using GetPixel appears to be faster than using VB's Point
           If TypeOf Name Is Form Then 'check to see if this is a form and use GetPixel function which is a little faster
               If GetPixel(lngHDc, intX, intY) = COLOR Then  ' If the pixel's color is the transparency color, record it
                   If booMiddleOfSet = False Then
                        dblTransStartX = intX
                        dblTransEndX = intX
                        booMiddleOfSet = True
                    Else
                        dblTransEndX = intX
                    End If 'booMiddleOfSet = False
               Else
                    If booMiddleOfSet Then
                        colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                        booMiddleOfSet = False
                    End If 'booMiddleOfSet = True
               End If 'GetPixel(lngHDC, X, Y) = Color
           ElseIf TypeOf Name Is PictureBox Then 'if a PictureBox then use Point; a little slower but works when GetPixel doesn't
               If Name.Point(intX, intY) = COLOR Then
                    If booMiddleOfSet = False Then
                        dblTransStartX = intX
                        dblTransEndX = intX
                        booMiddleOfSet = True
                    Else
                        dblTransEndX = intX
                    End If 'booMiddleOfSet = False
               Else
                    If booMiddleOfSet Then
                        colPoints.Add Array(dblTransY, dblTransStartX, dblTransEndX)
                        booMiddleOfSet = False
                    End If 'booMiddleOfSet = True
               End If 'Name.Point(X, Y) = Color
           End If 'TypeOf Name Is Form
           
        Next intX
    Next intY
   
    CurRgn = CreateRectRgn(0, 0, dblWidth, dblHeight)  ' Create base region which is the current whole window
   
    For Each Z In colPoints 'now make it transparent
       TempRgn = CreateRectRgn(Z(1), Z(0), Z(2) + 1, Z(0) + 1)  ' Create a temporary pixel region for this pixel
       CombineRgn CurRgn, CurRgn, TempRgn, RGN_DIFF  ' Combine temp pixel region with base region using RGN_DIFF to extract the pixel and make it transparent
       DeleteObject (TempRgn)  ' Delete the temporary region and free resources
   Next
   
    SetWindowRgn Name.hWnd, CurRgn, True
   
    Set colPoints = Nothing
   
End Sub

Public Function ShowMe(ByVal frmParent As Object, Optional ByVal lngFileID As Long) As Boolean
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '返回：确定返回新增或修改的ID；取消返回0
    '---------------------------------------------------
    
    mlngFileID = lngFileID

    Err = 0: On Error GoTo errHand
    
    If RefreshData = False Then
        DataChanged = False
        Unload Me
        Exit Function
    End If
    
    DataChanged = False
    
    '---------------------------------------------------
    '窗体显示
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    
    ShowMe = mblnOk
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function


Private Function RefreshData() As Boolean


    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '返回：确定返回新增或修改的ID；取消返回0
    '---------------------------------------------------
    Dim strTitle As String
    Dim strColRelation As String
    Dim arrColName, arrTmp, lngRow As Long, i As Integer, j As Integer
    Dim strColName  As String, strColNumber As String
    '
    With vfgThis
        .Cols = 6
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vsf
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vfgColRelation
        .Rows = 3
        .Tag = ""
        .Cell(flexcpText, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vsfColDoctor
        .Rows = 2
        .Tag = ""
    End With
    '---------------------------------------------------
    '根据当前打印机，装入可选纸张
    '---------------------------------------------------
    Call LoadPaper
    
    '46251,刘鹏飞,2012-09-11,装载页码输出位置
    With cbo页码
        .Clear
        .AddItem "页眉上方": .ItemData(.NewIndex) = 1
        .AddItem "页眉下方": .ItemData(.NewIndex) = 2
        .AddItem "页脚上方": .ItemData(.NewIndex) = 3
        .AddItem "页脚下方": .ItemData(.NewIndex) = 4
        cbo页码.Tag = 3
        Call zlcontrol.CboSetIndex(cbo页码.hWnd, 2)
    End With
    
    '---------------------------------------------------
    '基础数据获取
    '---------------------------------------------------
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select l.编号, l.名称, l.说明 From 病历文件列表 l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    Me.Caption = "护理记录样式 - " & rsTemp!名称
    strTitle = rsTemp!名称
    

                gstrSQL = _
                    " Select 中文名,替换域" & vbNewLine & _
                    " From 诊治所见项目 i, 诊治所见分类 k" & vbNewLine & _
                    " Where k.Id = i.分类id And ((k.编码 In ('02', '05', '06') And i.替换域 = 1) Or (k.性质 = 2 And k.编码 = '06' And NVL(i.替换域,0) = 0))" & vbNewLine & _
                    " Order By k.性质, k.编码, i.编码"
                Set rsLabItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                With rsLabItems
                    Me.cboLableSearch.Clear
                    Do While Not .EOF
                        Me.cboLableSearch.AddItem "" & !中文名
                        .MoveNext
                    Loop
                End With
    
    gstrSQL = "Select 项目序号,项目名称,项目类型,项目表示,项目长度 From 护理记录项目 Order By 项目名称"
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsItems
        Me.cboItemSearch.Clear
        Me.cboItemSearch.AddItem "日期"
        Me.cboItemSearch.AddItem "时间"
        Do While Not .EOF
            Me.cboItemSearch.AddItem "" & !项目名称
            .MoveNext
        Loop
        Me.cboItemSearch.AddItem "护士"
        Me.cboItemSearch.AddItem "签名人"
        Me.cboItemSearch.AddItem "签名时间"
        Me.cboItemSearch.ListIndex = 0
        .MoveFirst
    End With
    
    '---------------------------------------------------
    '定义样式获取
    '---------------------------------------------------
    '空表格时未能设置其表头层次,导致缺省表头层数是单层,但表头层数计数器最小值为2,在点击该表头时报错
    Me.optTabTiers(0).Value = True
    Call optTabTiers_Click(0)
    
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    mlngPageRow = 0
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数"
                If Val("" & !内容文本) = 1 Then
                    Me.optTabTiers(0).Value = True
                    Call optTabTiers_Click(0)
                ElseIf Val("" & !内容文本) = 2 Then
                    Me.optTabTiers(1).Value = True
                    Call optTabTiers_Click(1)
                Else
                    Me.optTabTiers(2).Value = True
                    Call optTabTiers_Click(2)
                End If
            Case "总列数":  Me.udTabCols.Value = Val("" & !内容文本)
            Case "最小行高"
                Me.txtTabRowHeight.Text = Val("" & !内容文本)
                Call txtTabRowHeight_Change
            Case "文本字体"
                mstrTextFont = "" & !内容文本
                strCurFont = mstrTextFont
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    cbsThis.FindControl(xtpControlButton, conMenu_Edit_BlodFont).Checked = False
                    cbsThis.FindControl(xtpControlButton, conMenu_Edit_Ttalic).Checked = False
                    cbsThis.FindControl(xtpControlButton, conMenu_Edit_Underline).Checked = False
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then
                        .Bold = True
                        cbsThis.FindControl(xtpControlButton, conMenu_Edit_BlodFont).Checked = True
                    End If
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True: cbsThis.FindControl(xtpControlButton, conMenu_Edit_Ttalic).Checked = True
                End With
                Set Me.vfgThis.Font = objFont
            Case "文本颜色"
                mlngTabTextColor = Val("" & !内容文本)
                vfgThis.Cell(flexcpForeColor, 6, 1, 6, vfgThis.Cols - 1) = mlngTabTextColor
'                Me.vfgThis.ForeColor = mlngTabTextColor
            Case "表格颜色"
                mlngTabGridColor = Val("" & !内容文本)
                With Me.vfgThis
                    .GridColor = mlngTabGridColor
                    .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
                    .CellBorderRange 1, .FixedCols, 2, .Cols - 1, vbWhite, 0, 0, 1, 0, 0, 0
                    .CellBorderRange 3, .FixedCols, 8, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
                End With
            
            Case "标题文本"
                Me.txtTitleText.Text = "" & !内容文本 '
            Case "标题字体"
                mstrTitleFont = "" & !内容文本
                strCurFont = mstrTitleFont
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                
                With Me.txtTitleText
                    Set txtTitleText.Font = objFont
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
                    .ROWHEIGHT(1) = objFont.Size * 20 + 150
                End With
                Call txtTitleText_Change
            Case "开始时间": Me.udRecordFrom.Value = Val("" & !内容文本)
            Case "终止时间": Me.udRecordTo.Value = Val("" & !内容文本)
            Case "条件字体"
                mstrRecordFont = "" & !内容文本
                strCurFont = mstrRecordFont
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
                End With
            Case "条件颜色"
                mlngRecordColor = Val("" & !内容文本)
                With Me.vfgThis
                    .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = mlngRecordColor
                End With
            Case "有效数据行"
                mlngPageRow = Val("" & !内容文本)
            Case "日期时间合并"
                chk首列合并.Value = Val("" & !内容文本)
            Case "时间列隐藏"
                chkHideTime.Value = Val("" & !内容文本)
            End Select
            .MoveNext
        Loop
    End With
    chk重算数据行.Value = (IIf(mlngPageRow = 0, 1, 0))
    
    Dim strPaper As String, blnHead As Boolean, blnFoot As Boolean
    gstrSQL = "Select 种类||'-'||编号 AS KEY,格式,页眉,页脚 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If Not rsTemp.EOF Then
        strPaper = "" & rsTemp!格式:
        blnHead = ReadPageHead(rtbHead, rsTemp!Key)
        blnFoot = ReadPageFoot(rtbFoot, rsTemp!Key)
        picFoot.Tag = rsTemp!Key
        
        chk页码.Value = IIf(Val(NVL(rsTemp!页脚, 0)) > 0, 1, 0)
        If chk页码.Value = 1 Then optPageAlign(Val(NVL(rsTemp!页脚, 0)) - 1).Value = True
        '46251,刘鹏飞,2012-09-11,装载页码输出位置
        If CInt(Val(NVL(rsTemp!页眉, 0))) > 0 And CInt(Val(NVL(rsTemp!页眉, 0))) < 5 Then
            cbo页码.ListIndex = CInt(Val(NVL(rsTemp!页眉, 0))) - 1
        End If
    End If
    
    If UBound(Split(strPaper, ";")) >= 0 Then
        For lngCount = 0 To Me.cboPaperKind.ListCount - 1
            If Me.cboPaperKind.ItemData(lngCount) = Val(Split(strPaper, ";")(0)) Then Me.cboPaperKind.ListIndex = lngCount: Exit For
        Next
        If Me.cboPaperKind.ListIndex = 0 Then
            If UBound(Split(strPaper, ";")) >= 2 Then Me.txtHeight.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(2)), vbTwips, vbMillimeters), 2)
            If UBound(Split(strPaper, ";")) >= 3 Then Me.txtWidth.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(3)), vbTwips, vbMillimeters), 2)
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 1 Then
        If Val(Split(strPaper, ";")(1)) = 2 Then
            Me.optOrient(1).Value = True
        Else
            Me.optOrient(0).Value = True
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 4 Then Me.txtMarjin(2).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 5 Then Me.txtMarjin(3).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 6 Then Me.txtMarjin(0).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 7 Then Me.txtMarjin(1).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(7)), vbTwips, vbMillimeters), 2)
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Me.lstLabelUsed.Clear
        Do While Not .EOF
            Me.lstLabelUsed.AddItem !内容文本 & "{" & !要素名称 & "}"
            Me.lstLabelUsed.ItemData(Me.lstLabelUsed.NewIndex) = !是否换行
            .MoveNext
        Loop
        If Me.lstLabelUsed.ListCount > 0 Then
            Me.lstLabelUsed.ListIndex = 0
            cboLableSearch.Enabled = True
            Me.chkLabelCrLf.Enabled = True
            Me.txtLabelPrefix.Enabled = True
        Else
            Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
            Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
        End If
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            Me.vfgThis.TextMatrix(!内容行次 + 2, !对象序号) = "" & !内容文本
            .MoveNext
        Loop
        Call udHeadCol_Change
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.对象序号,d.对象标记, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示,d.要素值域 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    strColRelation = ""
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.vfgThis.ColWidth(!对象序号) = Val(Split("" & !对象属性, "`")(0))
            If InStr(1, "" & !对象属性, "`") <> 0 Then
                vfgThis.Cell(flexcpAlignment, 6, !对象序号, 7, !对象序号) = Val(Split("" & !对象属性, "`")(1))
            Else
                vfgThis.Cell(flexcpAlignment, 6, !对象序号, 7, !对象序号) = flexAlignLeftCenter
            End If
            If Me.udColumnNo.Value <> !对象序号 Then Me.udColumnNo.Value = !对象序号
            Me.lstColumnUsed.AddItem !内容文本 & "{" & !要素名称 & "}" & !要素单位
            Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = ZLCommFun.NVL(!要素表示, 0)
            
            If IsNumeric(NVL(!对象标记)) And Val(NVL(!对象标记)) > 0 And Val(NVL(!对象标记)) <> Val(NVL(!对象序号)) Then
                strColRelation = strColRelation & "|" & Val(NVL(!对象序号)) & "," & Val(NVL(!对象标记))
            End If
            
            .MoveNext
            If .EOF Then
                vfgColRelation.Tag = "OK"
                Call cmdOK_Click
                strColRelation = Mid(strColRelation, 2)
                '初始化时加载之前设置的汇总列对照关系
                If strColRelation <> "" Then
                    arrColName = Split(strColRelation, "|")
                    With vfgColRelation
                        If .ColComboList(.ColIndex("对照名称")) <> "" Then
                            arrTmp = Split(.ColComboList(.ColIndex("对照名称")), "|#")
                            For lngRow = .FixedRows To .Rows - 1
                                If Val(.TextMatrix(lngRow, .ColIndex("汇总列号"))) > 0 Then
                                    For i = 0 To UBound(arrColName)
                                        If Val(Split(arrColName(i), ",")(0)) = Val(.TextMatrix(lngRow, .ColIndex("汇总列号"))) Then
                                            For j = 0 To UBound(arrTmp)
                                                If InStr(1, arrTmp(j), ";") > 0 And Val(Split(arrColName(i), ",")(1)) > 0 Then
                                                    If Val(Split(arrColName(i), ",")(1)) = Val(Mid(arrTmp(j), 1, InStr(1, arrTmp(j), ";") - 1)) Then
                                                        .TextMatrix(lngRow, .ColIndex("对照列号")) = Mid(arrTmp(j), 1, InStr(1, arrTmp(j), ";") - 1)
                                                        .TextMatrix(lngRow, .ColIndex("对照名称")) = Mid(arrTmp(j), InStr(1, arrTmp(j), ";") + 1)
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    Next i
                                End If
                            Next lngRow
                        End If
                    End With
                End If
            ElseIf Me.udColumnNo.Value <> !对象序号 Then
                Call cmdOK_Click
            End If
        Loop
        .Filter = " 内容行次 > 1"
        .Sort = "对象序号"
        strColNumber = ""
        Do While Not .EOF
            If Not InStr(1, strColNumber & ",", "," & !对象序号 & ",") > 0 Then
            strColNumber = strColNumber & "," & !对象序号
            End If
            .MoveNext
        Loop
        .Filter = " 内容行次 = 1"
        .Sort = "对象序号"
        mstrColDoctor = ""
        vsfColDoctor.Tag = "OK"
        strColName = " |用量|医嘱内容|总给予量|执行频次|用药目的|用药理由|医生嘱托|开始执行时间|给药途径"
        Do While Not .EOF
            If Not InStr(1, ",日期,时间,护士,签名人,签名时间,", "," & !要素名称 & ",") > 0 And Not InStr(strColNumber & ",", "," & !对象序号 & ",") > 0 Then
                vsfColDoctor.Rows = vsfColDoctor.Rows + 1
                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("项目列号")) = Val(NVL(!对象序号))
                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("项目名称")) = NVL(!要素名称)
                If Not NVL(!要素值域) = "" Then mstrColDoctor = mstrColDoctor & "|" & NVL(!要素名称) & ";" & NVL(!要素值域)
                vsfColDoctor.ComboList = strColName
            End If
            .MoveNext
        Loop
        Me.udColumnNo.Value = Me.vfgThis.Col
    End With
    If mstrTitleFont = "" Then mstrTitleFont = "宋体,9"
    If mstrRecordFont = "" Then mstrRecordFont = "宋体,9"
    If mstrTextFont = "" Then mstrTextFont = "宋体,9"
    
    
    mstrColDoctor = Mid(mstrColDoctor, 2)
    Call cmdOK_Click
    '汇总时段
    '------------------------------------------------------------------------------------------------------------------
    Dim aryTmp As Variant
    
    gstrSQL = "Select d.对象序号, d.对象属性, d.内容行次, d.内容文本 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '汇总时段'" & _
        " Order By d.对象序号, d.内容行次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With vsf
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                strTemp = ZLCommFun.NVL(rsTemp!内容文本)
                If strTemp <> "" Then

                    aryTmp = Split(strTemp, ",")

                    If UBound(aryTmp) >= 2 Then
                        If .TextMatrix(.Rows - 1, 1) <> "" And .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1

                        .TextMatrix(.Rows - 1, 1) = Trim(aryTmp(0))
                        .TextMatrix(.Rows - 1, 2) = Trim(aryTmp(1))
                        .TextMatrix(.Rows - 1, 3) = Trim(aryTmp(2))
                    End If
                End If
                rsTemp.MoveNext
            Loop
            mclsVsf.AppendRows = True
        End If
    End With
    
    '再按列合并
    For lngCount = 1 To vfgThis.Cols - 1
        vfgThis.MergeCol(lngCount) = True
    Next
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    Call SetTxtTitleSize
    Call cmdLabOK_Click
    RefreshData = True
    vfgThis.Row = 6
    vfgThis.Col = 1
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetItemName(ByVal strData As String, ByVal intOrder As Integer) As String
    Dim intPos As Integer
    '获取指定格式串中指定序号的项目名称，格式如：{收缩压}/ {舒张压}mmHg
    
    intPos = InStr(1, strData, "{")
    If intOrder > 0 Then intPos = InStr(intPos + 1, strData, "{")
    strData = Mid(strData, intPos + 1)
    strData = Mid(strData, 1, InStr(1, strData, "}") - 1)
    GetItemName = strData
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lngCol As Long, StrText As String) As Boolean

    Dim intPos As Integer
    Dim strTmp As String
    Dim lngHour As Long, lngMin As Long
    
    intPos = InStr(StrText, ":")
    
    If StrText = "" Then
        CheckTime = True
        Exit Function
    End If
    
    If intPos > 0 Then
        
        strTmp = Mid(StrText, 1, intPos - 1)
        lngHour = Val(strTmp)
        
        If Val(strTmp) < 0 Or Val(strTmp) > 23 Then
            MsgBox "小时只能在0-23之间！", vbInformation, gstrSysName
            Exit Function
        End If
        
        StrText = Mid(StrText, intPos + 1)
        intPos = InStr(StrText, ":")
        If intPos > 0 Then
            strTmp = Mid(StrText, 1, intPos - 1)
            lngMin = Val(strTmp)
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "分钟只能在0-59之间！", vbInformation, gstrSysName
                Exit Function
            End If
            
            If InStr(Mid(StrText, intPos + 1), ":") > 0 Then
                
                MsgBox "时间格式不正确！", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = Mid(StrText, intPos + 1)
                If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                    MsgBox "秒钟只能在0-59之间！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            strTmp = StrText
            lngMin = Val(strTmp)
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "分钟只能在0-59之间！", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    Else
        If Val(StrText) < 0 Or Val(StrText) > 23 Then
            MsgBox "小时只能在0-23之间！", vbInformation, gstrSysName
            Exit Function
        End If
        lngHour = Val(StrText)
    End If
    
    StrText = String(2 - Len(CStr(lngHour)), "0") & CStr(lngHour) & ":" & String(2 - Len(CStr(lngMin)), "0") & CStr(lngMin)
    vsf.TextMatrix(lngRow, lngCol) = StrText
    
    CheckTime = True
    
End Function

Private Function CheckData() As Boolean
    Dim intType As Integer, intFace As Integer, intLen As Integer                  '两个项目的表示方式一样才允许
    Dim bln护士 As Boolean
    Dim StrText As String, strItem As String
    Dim lngCol As Long, lngCount As Long
    Dim intDo As Integer, intHead As Integer
    Dim intRow As Integer
    
    '每种护理记录单都必须要有一列绑定护士才行
    lngCount = vfgThis.Cols - 1
    For lngCol = 1 To lngCount
        If InStr(1, ",{护士},{签名人},", "," & vfgThis.TextMatrix(6, lngCol) & ",") <> 0 Then
            bln护士 = True
        End If
    Next
    If Not bln护士 Then
        MsgBox "必须有一列绑定护士项目或签名人项目！", vbInformation, gstrSysName
        Exit Function
    End If
    If vfgThis.TextMatrix(6, 1) <> "{日期}" Then
        MsgBox "第一列必须绑定日期项目！", vbInformation, gstrSysName
        Exit Function
    End If
    If vfgThis.TextMatrix(6, 2) <> "{时间}" Then
        MsgBox "第二列必须绑定时间项目！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '只有绑定两个项目的列，项目的分隔符也是/，且项目类型必须相同,只能是录入、选择项或汇总项才允许设置列格式为对角线（特例：日期设置对角线表示日期简写）
    '只取列最近的列头（3列表头取5，2列表头取4，1列表头取3)
    If optTabTiers(0).Value Then
        intHead = 3
    ElseIf optTabTiers(1).Value Then
        intHead = 4
    Else
        intHead = 5
    End If
    For lngCol = 1 To lngCount
        If vfgThis.Cell(flexcpData, 6, lngCol) <> "" Then
            StrText = Val(Split(vfgThis.Cell(flexcpData, 6, lngCol), "`")(1))
            
            If StrText = 1 Then
                '格式：{A}{B}，按}分解，2列分解出来的数组就=2
                StrText = vfgThis.TextMatrix(6, lngCol)
                If UBound(Split(StrText, "}")) <> 2 Then
                    If StrText <> "{日期}" Then
                        MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 不允许设置列对角线！" & vbCrLf & "[只有日期列、绑定两个项目的列才允许设置列对角线]", vbInformation, gstrSysName
                        Exit Function
                    Else
                        GoTo ntloop
                    End If
                End If
                
                '两个项目的分隔符也必须是/
                If Trim(Mid(StrText, InStr(1, StrText, "}") + 1, InStr(InStr(1, StrText, "{") + 1, StrText, "{") - InStr(1, StrText, "}") - 1)) <> "/" Then
                    MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 设置了列对角线，要求绑定的项目格式应该是:A/B", vbInformation, gstrSysName
                    Exit Function
                End If
                
                '两个项目的项目类型必须一致
                For intDo = 0 To 1
                    strItem = Trim(GetItemName(StrText, intDo))
                    rsItems.Filter = "项目名称='" & strItem & "'"
                    '录入的时候检查过了,系统项如日期,时间等不允许与固定项绑定,所以,此时不存在找不到的情况
                    If rsItems.RecordCount = 0 Then
                        MsgBox "项目:" & strItem & "已改名或已经被删除！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If intDo > 0 Then
                        If Not (intFace = rsItems!项目表示 And intType = rsItems!项目类型) Then
                            MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 绑定的两个项目的编辑方式必须一致！", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If rsItems!项目类型 = 1 And rsItems!项目表示 = 0 And NVL(rsItems!项目长度, 1) > 3 Then
                            MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 绑定的两个文本项目长度不能大于3！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    Else
                        intFace = rsItems!项目表示
                        intType = rsItems!项目类型
                        intLen = NVL(rsItems!项目长度, 1)
                        If InStr(1, "0,2,4,5", intFace) = 0 Then
                            MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 绑定的两个项目必须都是数值型、选择项、单选项、汇总项或长度小于等于3的文本项目！", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If intType = 1 And intFace = 0 And intLen > 3 Then
                            MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 绑定的两个文本项目长度不能大于3！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
ntloop:
    Next
    
    '检查汇总列关系
    With vfgColRelation
        For intRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("汇总列号"))) > 0 And Val(.TextMatrix(intRow, .ColIndex("对照列号"))) > 0 Then
                If Val(.TextMatrix(intRow, .ColIndex("汇总列号"))) = Val(.TextMatrix(intRow, .ColIndex("对照列号"))) Then
                    MsgBox "表列关系中第" & intRow & "行汇总列号和对照列号属于同一列，请检查！", vbInformation, gstrSysName
                    Exit Function
                End If
                '检查设置的对照项目名称是否重复
                For lngCount = intRow + 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, .ColIndex("对照列号"))) = Val(.TextMatrix(lngCount, .ColIndex("对照列号"))) Then
                        MsgBox "表列关系中第" & intRow & "行和第" & lngCount & "行设置的对照项目重复,请检查！", vbInformation, gstrSysName
                        Exit Function
                    End If
                Next lngCount
            End If
        Next intRow
    End With
    
    '检查汇总时间
    With vsf
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) <> "" And .TextMatrix(intRow, 2) <> "" And .TextMatrix(intRow, 3) <> "" Then
                
                If CheckTime(intRow, 2, .TextMatrix(intRow, 2)) = False Then Exit Function
                If CheckTime(intRow, 3, .TextMatrix(intRow, 3)) = False Then Exit Function
                
            End If
        Next
    End With
    
    rsItems.Filter = 0
    CheckData = True
End Function

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean
    Dim blnReCalc As Boolean    '因每页的可打印数据行发生变化，对所有未打印完的文件进行重算
    Dim strCol As String
    Dim strInput As String
    Dim intPageAlign As Integer, intPageOption As Integer
    Dim lngRows As Long, lngFixedRows As Long
    Dim lngRow As Long, strRelation As String, blnRelation As Boolean '汇总列对照关系设定变量
    Dim strHead As String
    Dim strColItemRelation As String
    Dim str科室id As String
    Dim strHeadLab As String
    
    If CheckData = False Then Exit Function
        
    '保存数据
    If Me.optOrient(0).Value = True Then
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "上边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "下边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "左边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "右边距太大！", vbExclamation, gstrSysName: Exit Function
    Else
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "上边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "下边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "左边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "右边距太大！", vbExclamation, gstrSysName: Exit Function
    End If
    
    '计算有效数据行(纸张|纸向|高|宽|上边距|下边距|左边距|右边距|行高|固定行数|标题栏字体名|标题栏字体大小|标题文本|表上项字体名|表上项字体大小|表上项文本)
    If optTabTiers(0).Value Then
        lngFixedRows = 1
    ElseIf optTabTiers(1).Value Then
        lngFixedRows = 2
    Else
        lngFixedRows = 3
    End If
    
    strHeadLab = ""
    '获取表头信息
    For lngCount = vfgThis.FixedCols To vfgThis.Cols - 1
        If vfgThis.RowHidden(3) = False Then strHead = strHead & "'" & lngCount - 1 & ",0," & Trim(vfgThis.TextMatrix(3, lngCount)) & "," & vfgThis.ColWidth(lngCount)
        If vfgThis.RowHidden(4) = False Then strHead = strHead & "'" & lngCount - 1 & ",1," & Trim(vfgThis.TextMatrix(4, lngCount)) & "," & vfgThis.ColWidth(lngCount)
        If vfgThis.RowHidden(5) = False Then strHead = strHead & "'" & lngCount - 1 & ",2," & Trim(vfgThis.TextMatrix(5, lngCount)) & "," & vfgThis.ColWidth(lngCount)
    Next
    strHead = Mid(strHead, 2)
    strTemp = vfgThis.TextMatrix(2, 2)
    strInput = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex) & "|" & IIf(Me.optOrient(0).Value, 1, 2) & "|" & _
               Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips)) & "|" & _
               Val(txtTabRowHeight.Text) & "|" & lngFixedRows & "|" & Split(mstrTitleFont, ",")(0) & "|" & Split(mstrTitleFont, ",")(1) & "|" & _
               txtTitleText.Text & "|" & Split(mstrTextFont, ",")(0) & "|" & Split(mstrTextFont, ",")(1) & "|" & strHeadLab & "|" & strHead
    lngRows = frmTendFilePreview.ShowMe(Me, strInput)
    
    
    If mintBlnReCulat = True Then
        str科室id = frmTendRecalCulation.ShowEditor(mlngFileID)
        If str科室id = "" Then
            MsgBox "取消数据重算，请重新设置数据重算！", vbInformation, gstrSysName
            Exit Function
        End If
        If lngRows <> mlngPageRow And mlngPageRow > 0 Then
            '数据行发生变化，对已打印数据有影响，提示
            If MsgBox("    本次修改导致每页可打印的数据行发生变化，已打印的数据若修改或重打将会导致打印混乱，是否继续？" & vbCrLf & "    原来每页可打印" & mlngPageRow & "行；现在每页可打印" & lngRows & "行", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnReCalc = True
            '87990,陈刘,2015-09-02
        Else
            If MsgBox("    本次将按照当前选择重算方式,对使用该格式的文件重新计算数据行数，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnReCalc = True
        End If
    Else
        lngRows = mlngPageRow
    End If
    
    '如果超出页面宽度则禁止保存
    If OverRun Then
        MsgBox "表格的宽度超出了纸张有效打印范围！", vbInformation, gstrSysName
        Exit Function
    End If
    If Not PageHeadTest Then Exit Function
    
    '拼串
    If chk页码.Value = 1 Then
        If optPageAlign(0).Value Then
            intPageAlign = 1
        ElseIf optPageAlign(1).Value Then
            intPageAlign = 2
        Else
            intPageAlign = 3
        End If
        '46251,刘鹏飞,2012-09-11
        intPageOption = Val(cbo页码.ItemData(cbo页码.ListIndex))
    End If
    If Me.optTabTiers(0).Value Then
        gstrSQL = mlngFileID & ",1," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    ElseIf Me.optTabTiers(1).Value Then
        gstrSQL = mlngFileID & ",2," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    Else
        gstrSQL = mlngFileID & ",3," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    End If

    gstrSQL = gstrSQL & ",'" & mstrTextFont & "'," & mlngTabTextColor & "," & mlngTabGridColor
    gstrSQL = gstrSQL & ",'" & Trim(Me.txtTitleText.Text) & "','" & mstrTitleFont & "'"
    gstrSQL = gstrSQL & "," & Me.udRecordFrom.Value & "," & Me.udRecordTo.Value & ",'" & mstrRecordFont & "'," & mlngRecordColor & "," & lngRows & "," & chk首列合并.Value & "," & chkHideTime.Value
    
    gstrSQL = gstrSQL & ",'" & Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
    gstrSQL = gstrSQL & ";" & IIf(Me.optOrient(0).Value, 1, 2)
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "'"
    gstrSQL = gstrSQL & "," & intPageOption & "," & intPageAlign
    
    With Me.vfgThis
        gstrSQL = gstrSQL & ",'" & Replace(Replace(.TextMatrix(2, .FixedCols), " ", "|"), vbCrLf, "'||Chr(13)||Chr(10)||'") & "'"
        strTemp = ""
        For lngCount = .FixedCols To .Cols - 1
            If .RowHidden(3) = False Then strTemp = strTemp & "|" & lngCount & ",1," & Trim(.TextMatrix(3, lngCount))
            If .RowHidden(4) = False Then strTemp = strTemp & "|" & lngCount & ",2," & Trim(.TextMatrix(4, lngCount))
            If .RowHidden(5) = False Then strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
'            strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strTemp, 2) & "'"
        
        strCol = ""
        blnRelation = False
        For lngCount = .FixedCols To .Cols - 1
            '汇总列与名称列对照关系的设定
            strRelation = ""
            If Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                For lngRow = vfgColRelation.FixedRows To vfgColRelation.Rows - 1
                    If Val(vfgColRelation.TextMatrix(lngRow, vfgColRelation.ColIndex("汇总列号"))) = lngCount And Val(vfgColRelation.TextMatrix(lngRow, vfgColRelation.ColIndex("对照列号"))) > 0 Then
                        strRelation = Val(vfgColRelation.TextMatrix(lngRow, vfgColRelation.ColIndex("对照列号")))
                        Exit For
                    End If
                Next lngRow
            End If
            
            strColItemRelation = ""
            If Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                For lngRow = vsfColDoctor.FixedRows To vsfColDoctor.Rows - 1
                    If Val(vsfColDoctor.TextMatrix(lngRow, vsfColDoctor.ColIndex("项目列号"))) = lngCount And Trim(vsfColDoctor.TextMatrix(lngRow, vsfColDoctor.ColIndex("对照名称"))) <> "" Then
                        strColItemRelation = vsfColDoctor.TextMatrix(lngRow, vsfColDoctor.ColIndex("对照名称"))
                        Exit For
                    End If
                Next lngRow
            End If
            
            If strRelation = "" Then
                strRelation = lngCount
            Else
                strRelation = lngCount & "`" & Val(strRelation)
                blnRelation = True
            End If
            strCol = strCol & "|" & strRelation & "," & .ColWidth(lngCount) & "`" & .Cell(flexcpAlignment, 6, lngCount) & "," & Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount)) & "," & strColItemRelation
'            strTemp = strTemp & "|" & lngCount & "," & .ColWidth(lngCount) & "," & Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strCol, 2) & "'"
    End With
    
    '填写汇总时段
    '------------------------------------------------------------------------------------------------------------------
    With vsf
        strTemp = ""
        
        For lngCount = 1 To .Rows - 1
            If .TextMatrix(lngCount, 1) <> "" And .TextMatrix(lngCount, 2) <> "" And .TextMatrix(lngCount, 3) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(lngCount, 1)) & "," & Trim(.TextMatrix(lngCount, 2)) & "," & Trim(.TextMatrix(lngCount, 3))
            Else
                .TextMatrix(lngCount, 1) = ""
                .TextMatrix(lngCount, 2) = ""
                .TextMatrix(lngCount, 3) = ""
            End If
        Next
        
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        gstrSQL = gstrSQL & ",'" & strTemp & "'"
    End With
    
    gstrSQL = gstrSQL & ",NULL"
    gstrSQL = "Zl_护理文件样式_Update(" & gstrSQL & ")"
    
    Err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    blnTrans = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If Not SavePageHead(picFoot.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    If Not SavePageFoot(picFoot.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    
    SaveData = True
    mblnOk = True
    mlngPageRow = lngRows
    mintBlnReCulat = False
'    cmd同步.Enabled = True
    
    '因数据行发生变化，进行数据重算
    If blnReCalc Then
        
        If str科室id <> "" Or Val(str科室id) = -1 Then
            strInput = strInput & "|" & str科室id
            If frmTendFilePreview.AnaliseData(Me, mlngFileID, strInput) Then MsgBox "打印数据已更新！", vbInformation, gstrSysName
        End If
    End If
    Exit Function

errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePageHead(ByVal StrKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:产生文件并压缩;True:已产生压缩文件
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Head_S.rtf"
        rtbHead.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageHead = zlBlobSave(12, StrKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function


Private Function SavePageFoot(ByVal StrKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:产生文件并压缩;True:已产生压缩文件
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Foot_S.rtf"
        rtbFoot.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageFoot = zlBlobSave(13, StrKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function

Private Sub cboItemSearch_Click()
    If Me.cboItemSearch.ListIndex = -1 Then Exit Sub
    lstColumnUsed.AddItem "{" & Me.cboItemSearch.List(Me.cboItemSearch.ListIndex) & "}"
    lstColumnUsed.ListIndex = lstColumnUsed.NewIndex
    Me.txtColumnPrefix.Enabled = True
    Me.txtColumnPostfix.Enabled = True
    chk.Enabled = True
End Sub

Private Sub InstallColRelation()
'功能:装载汇总列与名称列关系
'-------------------------------------
    Dim strTmp As String
    Dim intType As Integer, intFace As Integer, intLen As Integer      '项目类型
    Dim strName As String  '项目名称
    Dim intCount As Integer, i As Integer
    Dim arrColName, arrTmp
    Dim arrColNo, strColName As String, j As Integer
    
    If vfgColRelation.Tag <> "OK" Then Exit Sub
    
    With vfgColRelation
        '首先获取之前设定的列的关系
        arrColName = Array()
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, .ColIndex("汇总列号"))) > 0 And Val(.TextMatrix(lngCount, .ColIndex("对照列号"))) > 0 And Val(.TextMatrix(lngCount, .ColIndex("汇总列号"))) <> Val(.TextMatrix(lngCount, .ColIndex("对照列号"))) Then
                ReDim Preserve arrColName(UBound(arrColName) + 1)
                arrColName(UBound(arrColName)) = .TextMatrix(lngCount, .ColIndex("汇总名称")) & "'" & .TextMatrix(lngCount, .ColIndex("对照名称"))
            End If
        Next lngCount
        
        '汇总项目与名称列的关系设定，必须满足：列只绑定了一个项目,名称列可为文本或单选，文本项目长度小于100(因为记录单分组数据处理文本长度>=100认为是大文本项目)
        .Rows = .FixedRows + 1
        .Cell(flexcpText, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        intCount = .Rows
        strTmp = "": .ColComboList(.ColIndex("对照名称")) = ""
        For lngCount = 1 To vfgThis.Cols - 1
            If Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                arrTmp = Split(Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)), " ")
                If UBound(arrTmp) = 0 Then
                    If InStr(1, arrTmp(0), "`") > 0 Then
                        strName = CStr(Split(arrTmp(0), "`")(0))
                    Else
                        strName = CStr(arrTmp(0))
                    End If
                    strName = Mid(strName, InStr(1, strName, "{") + 1)
                    strName = Mid(strName, 1, InStr(1, strName, "}") - 1)
                    If strName <> "" Then
                        rsItems.Filter = "项目名称='" & strName & "'"
                        If rsItems.RecordCount > 0 Then
                            intType = rsItems!项目类型
                            intFace = rsItems!项目表示
                            intLen = NVL(rsItems!项目长度, 1)
                            If intType = 0 And intFace = 4 Then '汇总项目确定
                                If intCount > .Rows Then .Rows = .Rows + 1
                                .TextMatrix(.Rows - 1, .ColIndex("汇总列号")) = lngCount
                                .TextMatrix(.Rows - 1, .ColIndex("汇总名称")) = strName
                                intCount = intCount + 1
                            '名称列项目确定
                            ElseIf intType = 1 And intFace = 0 And intLen < 100 Then '文本项目
                                strTmp = strTmp & "|#" & lngCount & ";" & strName
                            ElseIf intType = 1 And intFace = 2 Then '单选项目
                                strTmp = strTmp & "|#" & lngCount & ";" & strName
                            End If
                        End If
                    End If
                End If
            End If
        Next lngCount
        If strTmp <> "" Then
            arrTmp = Split(Mid(strTmp, 3), "|#")
            .ColComboList(.ColIndex("对照名称")) = "#; " & strTmp
            '恢复原有绑定的列内容
            If UBound(arrColName) >= 0 Then
                For lngCount = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(lngCount, .ColIndex("汇总列号"))) > 0 Then
                        For intCount = 0 To UBound(arrColName)
                            If .TextMatrix(lngCount, .ColIndex("汇总名称")) = Split(arrColName(intCount), "'")(0) Then
                                For i = 0 To UBound(arrTmp)
                                    If Split(arrColName(intCount), "'")(1) = Mid(arrTmp(i), InStr(1, arrTmp(i), ";") + 1) Then
                                        .TextMatrix(lngCount, .ColIndex("对照列号")) = Mid(arrTmp(i), 1, InStr(1, arrTmp(i), ";") - 1)
                                        .TextMatrix(lngCount, .ColIndex("对照名称")) = Split(arrColName(intCount), "'")(1)
                                    End If
                                Next i
                            End If
                        Next intCount
                    End If
                Next lngCount
            End If
        End If
    End With
    
    '加载医嘱绑定表格
    With vsfColDoctor
        .Rows = .FixedRows + 1
        .Cell(flexcpText, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        If vsfColDoctor.Tag = "OK" Then
            arrColNo = Split(mstrColDoctor, "|")
            
            intCount = .Rows
            For lngCount = 1 To vfgThis.Cols - 1
                If Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                    arrTmp = Split(Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)), " ")
                    If UBound(arrTmp) = 0 Then
                        If InStr(1, arrTmp(0), "`") > 0 Then
                            strName = CStr(Split(arrTmp(0), "`")(0))
                        Else
                            strName = CStr(arrTmp(0))
                        End If
                        strName = Mid(strName, InStr(1, strName, "{") + 1)
                        strName = Mid(strName, 1, InStr(1, strName, "}") - 1)
                        If strName <> "" Then
                            rsItems.Filter = "项目名称='" & strName & "'"
                            If rsItems.RecordCount > 0 Then
                                intType = rsItems!项目类型
                                intFace = rsItems!项目表示
                                If intCount > .Rows Then .Rows = .Rows + 1
                                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("项目列号")) = lngCount
                                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("项目名称")) = strName
                                vsfColDoctor.ComboList = " |用量|医嘱内容|总给予量|执行频次|用药目的|用药理由|医生嘱托|开始执行时间|给药途径"
                                intCount = intCount + 1
                                
                            End If
                        End If
                    End If
                End If
            Next lngCount
            '设置已经绑定的列
            For i = 0 To UBound(arrColNo)
                strColName = Split(arrColNo(i), ";")(0)
                strName = NVL(Split(arrColNo(i), ";")(1))
                For j = vsfColDoctor.FixedRows To vsfColDoctor.Rows - 1
                    If vsfColDoctor.TextMatrix(j, .ColIndex("项目名称")) = strColName Then vsfColDoctor.TextMatrix(j, .ColIndex("对照名称")) = strName: Exit For
                Next
            Next
        End If
    End With
End Sub

Private Sub cboItemSearch_DropDown()
    If Not cboItemSearch.ListCount > 0 Then
        Me.cboItemSearch.AddItem "日期"
        Me.cboItemSearch.AddItem "时间"
        Me.cboItemSearch.AddItem "护士"
        Me.cboItemSearch.AddItem "签名人"
        Me.cboItemSearch.AddItem "签名时间"
    End If
End Sub

Private Sub cboItemSearch_KeyPress(KeyAscii As Integer)
    Dim lngRet As Long
    Dim StrText As String
    If KeyAscii = 13 Then
        StrText = Me.cboItemSearch.Text
        lngRet = SendMessage(cboItemSearch.hWnd, CB_FINDSTRINGEXACT, -1, ByVal StrText)
        rsItems.Filter = ""
        If lngRet = -1 Then
            Me.cboItemSearch.Clear
            StrText = Trim(StrText)
            If StrText <> "" Then
                rsItems.Filter = "项目名称 like *" & StrText & "*"
                If InStr(1, "日期", StrText) > 0 Then Me.cboItemSearch.AddItem "日期"
                If InStr(1, "时间", StrText) > 0 Then Me.cboItemSearch.AddItem "时间"
                If InStr(1, "护士", StrText) > 0 Then Me.cboItemSearch.AddItem "护士"
                If InStr(1, "签名人", StrText) > 0 Then Me.cboItemSearch.AddItem "签名人"
                If InStr(1, "签名时间", StrText) > 0 Then Me.cboItemSearch.AddItem "签名时间"
            End If
            If rsItems.RecordCount > 0 Then
                    
                Do While Not rsItems.EOF
                    Me.cboItemSearch.AddItem "" & rsItems!项目名称
                    rsItems.MoveNext
                Loop
            Else
                
            End If
            SendMessage cboItemSearch.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
        Else
            cboItemSearch.ListIndex = lngRet
        End If
    End If
    rsItems.Filter = ""
End Sub

Private Sub cboLableSearch_Click()
    With Me.lstLabelUsed
        If Me.cboLableSearch.ListIndex = -1 Then Exit Sub
        .AddItem Me.cboLableSearch.List(Me.cboLableSearch.ListIndex) & "：{" & Me.cboLableSearch.List(Me.cboLableSearch.ListIndex) & "}"
        .ListIndex = .NewIndex
        Me.imgLabDelete.Enabled = True
        Me.chkLabelCrLf.Enabled = True
        Me.txtLabelPrefix.Enabled = True
    End With
End Sub

Private Sub cboLableSearch_DblClick()
    Call cboLableSearch_Click
End Sub

Private Sub cboLableSearch_KeyPress(KeyAscii As Integer)
    Dim lngRet As Long
    Dim StrText As String
    If KeyAscii = 13 Then
        StrText = Me.cboLableSearch.Text
        lngRet = SendMessage(cboLableSearch.hWnd, CB_FINDSTRINGEXACT, -1, ByVal StrText)
       
        If lngRet = -1 Then
            Me.cboLableSearch.Clear
            If StrText = "" Then
                rsLabItems.Filter = ""
            Else
                rsLabItems.Filter = "中文名 like *" & StrText & "*"
            End If
            Do While Not rsLabItems.EOF
                Me.cboLableSearch.AddItem "" & rsLabItems!中文名
                rsLabItems.MoveNext
            Loop
            SendMessage cboLableSearch.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
        Else
            cboLableSearch.ListIndex = lngRet
        End If
    End If
    rsLabItems.Filter = ""
End Sub

Private Sub cboPaperKind_Click()
    Dim intOrientation As Integer
    If Me.optOrient(0).Value = True Then
        intOrientation = 1
    Else
        intOrientation = 2
    End If
    '设置打印机纸张方向固定为1，保存时在重新设置
    Printer.Orientation = 1
    If Me.cboPaperKind.ListIndex <= 0 Then
        Me.txtWidth.Enabled = True: Me.udWidth.Enabled = True
        Me.txtHeight.Enabled = True: Me.udHeight.Enabled = True
        Me.optOrient(0).Value = True
        Me.optOrient(0).Enabled = True: Me.optOrient(1).Enabled = True
        If intOrientation = 1 Then
            Me.optOrient(0).Value = True
        Else
            Me.optOrient(1).Value = True
        End If
    Else
        Me.txtWidth.Enabled = False: Me.udWidth.Enabled = False
        Me.txtHeight.Enabled = False: Me.udHeight.Enabled = False
        Me.optOrient(0).Enabled = True: Me.optOrient(1).Enabled = True
        Err = 0: On Error Resume Next
        Printer.PaperSize = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
        Me.txtWidth.Text = Me.ScaleX(Printer.Width, vbTwips, vbMillimeters)
        Me.txtHeight.Text = Me.ScaleY(Printer.Height, vbTwips, vbMillimeters)
        If intOrientation = 1 Then
            Me.optOrient(0).Value = True
        Else
            Me.optOrient(1).Value = True
        End If
    End If
    DataChanged = True
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim intIndex As Integer
    Dim CbsEdit As CommandBarEdit
    Dim CbsCombo As CommandBarComboBox
    Dim cbrControl As CommandBarControl
    Dim lngCol As Long, lngRow As Long
    Dim strFont As String, strFontNameSize As String
    Dim intType As Integer
    Dim intAlign As Integer
    Dim blnChecked As Boolean
    Dim strValue As String
    Select Case Control.ID
    
    Case conMenu_Edit_Left_Alignment, conMenu_Edit_Center_Alignment, conMenu_Edit_Right_Alignment
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Right_Alignment)
        cbrControl.Checked = False
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Center_Alignment)
        cbrControl.Checked = False
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Left_Alignment)
        cbrControl.Checked = False
        Select Case Control.ID
        Case conMenu_Edit_Left_Alignment
            Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Left_Alignment)
            cbrControl.Checked = True
            intAlign = flexAlignLeftCenter
            vfgThis.TextMatrix(7, vfgThis.Col) = vfgThis.TextMatrix(6, vfgThis.Col) & " "
            DataChanged = True
        Case conMenu_Edit_Center_Alignment
            Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Center_Alignment)
            cbrControl.Checked = True
            intAlign = flexAlignCenterCenter
            vfgThis.TextMatrix(7, vfgThis.Col) = " " & vfgThis.TextMatrix(6, vfgThis.Col) & " "
            DataChanged = True
        Case conMenu_Edit_Right_Alignment
            Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Right_Alignment)
            cbrControl.Checked = True
            intAlign = flexAlignRightCenter
            vfgThis.TextMatrix(7, vfgThis.Col) = " " & vfgThis.TextMatrix(6, vfgThis.Col)
        End Select
        vfgThis.Cell(flexcpAlignment, 6, vfgThis.Col, 7, vfgThis.Col) = intAlign
        DataChanged = True
    Case conMenu_Edit_BlodFont
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_BlodFont)
        If cbrControl.Checked Then
            cbrControl.Checked = False
        Else
            cbrControl.Checked = True
        End If
        blnChecked = cbrControl.Checked
        Select Case mintType
            Case TYPE_Type.Type_表格文本
                If blnChecked Then
                    mstrTextFont = mstrTextFont & ",粗"
                Else
                    mstrTextFont = Replace(mstrTextFont, ",粗", "")
                End If
                strFont = mstrTextFont
            Case TYPE_Type.Type_表头文本
                If blnChecked Then
                    mstrTitleFont = mstrTitleFont & ",粗"
                Else
                    mstrTitleFont = Replace(mstrTitleFont, ",粗", "")
                End If
                strFont = mstrTitleFont
            Case TYPE_Type.Type_条件文本
                If blnChecked Then
                    mstrRecordFont = mstrRecordFont & ",粗"
                Else
                    mstrRecordFont = Replace(mstrRecordFont, ",粗", "")
                End If
                strFont = mstrRecordFont
            Case TYPE_Type.Type_页眉页脚
                Call GetrtbObject
                If blnChecked = True Then
                    rtbThis.SelBold = True
                Else
                    rtbThis.SelBold = False
                End If
        End Select
        If strFont <> "" Then Call SetEditorFont(strFont)
        mblnChanged = True
    Case conMenu_Edit_Ttalic
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Ttalic)
        If cbrControl.Checked Then
            cbrControl.Checked = False
        Else
            cbrControl.Checked = True
        End If
        blnChecked = cbrControl.Checked
        Select Case mintType
            Case TYPE_Type.Type_表格文本
                If blnChecked Then
                    mstrTextFont = mstrTextFont & ",斜"
                Else
                    mstrTextFont = Replace(mstrTextFont, ",斜", "")
                End If
                strFont = mstrTextFont
            Case TYPE_Type.Type_表头文本
                If blnChecked Then
                    mstrTitleFont = mstrTitleFont & ",斜"
                Else
                    mstrTitleFont = Replace(mstrTitleFont, ",斜", "")
                End If
                strFont = mstrTitleFont
            Case TYPE_Type.Type_条件文本
                If blnChecked Then
                    mstrRecordFont = mstrRecordFont & ",斜"
                Else
                    mstrRecordFont = Replace(mstrRecordFont, ",斜", "")
                End If
                strFont = mstrRecordFont
            Case TYPE_Type.Type_页眉页脚
                Call GetrtbObject
                If blnChecked = True Then
                    rtbThis.SelItalic = True
                Else
                    rtbThis.SelItalic = False
                End If
        End Select
            If strFont <> "" Then Call SetEditorFont(strFont)
        mblnChanged = True
    Case conMenu_Edit_Underline
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Underline)
        If cbrControl.Checked Then
            cbrControl.Checked = False
        Else
            cbrControl.Checked = True
        End If
        blnChecked = cbrControl.Checked
        If mintType = TYPE_Type.Type_页眉页脚 Then
            Call GetrtbObject
            If blnChecked = True Then
                rtbThis.SelUnderline = True
            Else
                rtbThis.SelUnderline = False
            End If
        End If
        
        mblnChanged = True
    Case conMenu_File_Preview
        strValue = GetTab
        Call frmTendWavePrint.TendPreview(strValue, rtbHead, rtbFoot)
    Case conMenu_Edit_MarkMap
        Dim picTemp As StdPicture
    
        With Me.dlgThis
            .DialogTitle = "标志图选择"
            .Filename = ""
            .Filter = "图像|*.jpg;*.bmp;*.ico;*.gif"
            .CancelError = True
            On Error Resume Next
            .ShowOpen
            If Err.Number <> 0 Then
                Err.Clear
                Exit Sub
            End If
        End With
        Set picTemp = Nothing
        Set picTemp = LoadPicture(Me.dlgThis.Filename)
        If picTemp Is Nothing Then MsgBox "不是有效的图片文件！", vbExclamation, Me.Caption: Exit Sub
        
        Clipboard.Clear
        Clipboard.SetData picTemp
        
        Call GetrtbObject
        SendMessageLong rtbThis.hWnd, WM_PASTE, 0, 0
        mblnChanged = True
        
    Case conMenu_Edit_Curve_AddLeft
        Me.vfgThis.Cols = Me.vfgThis.Cols + 1
        Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
        Me.vfgThis.ColPosition(vfgThis.Cols - 1) = Val(vfgThis.Tag)
        With Me.vfgThis
            For lngCount = .FixedCols To .Cols - 1
                .TextMatrix(0, lngCount) = lngCount
                .TextMatrix(1, lngCount) = .TextMatrix(1, IIf(Val(vfgThis.Tag) = .FixedCols, .FixedCols + 1, .FixedCols))
                .TextMatrix(2, lngCount) = .TextMatrix(2, IIf(Val(vfgThis.Tag) = .FixedCols, .FixedCols + 1, .FixedCols))
                .TextMatrix(.Rows - 1, lngCount) = .TextMatrix(.Rows - 1, IIf(Val(vfgThis.Tag) = .FixedCols, .FixedCols + 1, .FixedCols))
                .ColAlignment(lngCount) = flexAlignCenterCenter
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .CellBorderRange 2, .FixedCols, 2, .Cols - 1, mlngTabGridColor, 0, 0, 0, 1, 0, 0
            .CellBorderRange 3, .FixedCols, 8, .Cols - 1, mlngTabGridColor, 1, 1, 1, 1, 1, 1
            .MergeCol(-1) = True
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        Me.udHeadCol.Max = vfgThis.Cols - 1
        Me.udTabCols.Value = vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Col)
        mblnChanged = True
    Case conMenu_Edit_Curve_AddRight
        Me.vfgThis.Cols = Me.vfgThis.Cols + 1
        Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
        Me.vfgThis.ColPosition(vfgThis.Cols - 1) = Val(vfgThis.Tag) + 1
        With Me.vfgThis
            For lngCount = .FixedCols To .Cols - 1
                .TextMatrix(0, lngCount) = lngCount
                .TextMatrix(1, lngCount) = .TextMatrix(1, .FixedCols)
                .TextMatrix(2, lngCount) = .TextMatrix(2, .FixedCols)
                .TextMatrix(.Rows - 1, lngCount) = .TextMatrix(.Rows - 1, .FixedCols)
                .ColAlignment(lngCount) = flexAlignCenterCenter
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .CellBorderRange 2, .FixedCols, 2, .Cols - 1, mlngTabGridColor, 0, 0, 0, 1, 0, 0
            .CellBorderRange 3, .FixedCols, 8, .Cols - 1, mlngTabGridColor, 1, 1, 1, 1, 1, 1
            .MergeCol(-1) = True
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        Me.udHeadCol.Max = vfgThis.Cols - 1
        Me.udTabCols.Value = vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Col)
        mblnChanged = True
    Case conMenu_Edit_Curve_DeleteCol
        If vfgThis.Cols <= 4 Then Exit Sub
        With Me.vfgThis
            For lngCount = vfgThis.Col To .Cols - 2
                .TextMatrix(0, lngCount) = .TextMatrix(0, lngCount + 1)
                .TextMatrix(1, lngCount) = .TextMatrix(1, lngCount + 1)
                .TextMatrix(2, lngCount) = .TextMatrix(2, lngCount + 1)
                .TextMatrix(3, lngCount) = .TextMatrix(3, lngCount + 1)
                .TextMatrix(4, lngCount) = .TextMatrix(4, lngCount + 1)
                .TextMatrix(5, lngCount) = .TextMatrix(5, lngCount + 1)
                .TextMatrix(6, lngCount) = .TextMatrix(6, lngCount + 1)
                .TextMatrix(7, lngCount) = .TextMatrix(7, lngCount + 1)
                .Cell(flexcpData, 6, lngCount, 6, lngCount) = .Cell(flexcpData, 6, lngCount + 1, 6, lngCount + 1)
            Next
            For lngCount = vfgThis.Col To .Cols - 2
                .ColWidth(lngCount) = .ColWidth(lngCount + 1)
            Next
        End With
        Me.vfgThis.Cols = Me.vfgThis.Cols - 1
        With Me.vfgThis
            For lngCount = .FixedCols To .Cols - 1
                .TextMatrix(0, lngCount) = lngCount
                .ColAlignment(lngCount) = flexAlignCenterCenter
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .CellBorderRange 2, .FixedCols, 2, .Cols - 1, mlngTabGridColor, 0, 0, 0, 1, 0, 0
            .CellBorderRange 3, .FixedCols, 8, .Cols - 1, mlngTabGridColor, 1, 1, 1, 1, 1, 1
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        Me.udHeadCol.Max = vfgThis.Cols - 1
        Me.udTabCols.Value = vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Col)
        mblnChanged = True
    Case conMenu_Edit_Curve_AddUP
        
        With vfgThis
            If optTabTiers(0).Value Then
                intIndex = 0
            ElseIf optTabTiers(1).Value Then
                intIndex = 1
            Else
                intIndex = 2
            End If
            If intIndex < 2 Then optTabTiers(intIndex + 1).Value = True
            For lngRow = 5 To .Row Step -1
                For lngCol = .FixedCols To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow - 1, lngCol)
                Next
            Next
            vfgThis.Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
        End With
        Me.udTabCols.Value = vfgThis.Cols - 1
        vfgThis.AutoSize 0, vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Col)
        DataChanged = True
    Case conMenu_Edit_Curve_AddBottom
        With vfgThis
            If optTabTiers(0).Value Then
                intIndex = 0
            ElseIf optTabTiers(1).Value Then
                intIndex = 1
            Else
                intIndex = 2
            End If
            If intIndex < 2 Then optTabTiers(intIndex + 1).Value = True
            For lngRow = 5 To .Row + 1 Step -1
                For lngCol = .FixedCols To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow - 1, lngCol)
                Next
            Next
            vfgThis.Cell(flexcpText, .Row + 1, .FixedCols, .Row + 1, .Cols - 1) = ""
        End With
        vfgThis.AutoSize 0, vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Col)
        DataChanged = True
    Case conMenu_Edit_Curve_DeleteRow
        For lngRow = vfgThis.Row To 5
            For lngCol = vfgThis.FixedCols To vfgThis.Cols - 1
               vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow + 1, lngCol)
            Next
        Next
        If optTabTiers(1).Value Then
            optTabTiers(0).Value = True
        ElseIf optTabTiers(2).Value Then
            optTabTiers(1).Value = True
        End If
        vfgThis.Cell(flexcpText, 5, vfgThis.FixedCols, 5, vfgThis.Cols - 1) = ""
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Col)
    Case conMenu_Edit_Curve_BuddySingle
        Call vfgThis_DblClick
    Case conMenu_Edit_Curve_BuddyDouble
        picCloumn.Visible = True
        picCloumn.Left = vfgThis.Left + vfgThis.CellLeft
        picCloumn.Top = vfgThis.Top + vfgThis.CellTop
        picCloumn.Tag = 2
        udColumnNo.Value = vfgThis.Col
        txtColumnNo.Text = vfgThis.Col
        
        
    Case ID_FORMAT_FONTNAME, ID_FORMAT_FONTSIZE
    
        If Control.ID = ID_FORMAT_FONTNAME Then
            intType = 0
            Set CbsCombo = cbsThis.FindControl(xtpControlComboBox, ID_FORMAT_FONTNAME)
            strFontNameSize = CbsCombo.Text
        Else
            intType = 1
            Set CbsCombo = cbsThis.FindControl(xtpControlComboBox, ID_FORMAT_FONTSIZE)
            strFontNameSize = CbsCombo.Text
            If Not IsNumeric(strFontNameSize) Then
            strFontNameSize = GetFontSizeNumber(strFontNameSize)
            End If
        End If
        
        Select Case mintType
            Case Type_表头文本
                strFont = mstrTitleFont
            Case Type_表格文本
                strFont = mstrTextFont
            Case Type_条件文本
                strFont = mstrRecordFont
            Case Type_页眉页脚
                Call GetrtbObject
                If intType = 0 Then
                    rtbThis.SelFontName = strFontNameSize
                Else
                    rtbThis.SelFontSize = strFontNameSize
                End If
            Case Else
                strFont = mstrTextFont
        End Select
        
        If mintType <> Type_页眉页脚 Then
            strFont = SetFileFont(strFont, intType, strFontNameSize)
            If strFont <> "" Then Call SetEditorFont(strFont)
        End If
        
        
        
        mblnChanged = True
    Case conMenu_Edit_SummaryTime
    
        picSum.Visible = True
        picSum.Move picPane(1).Left, vfgThis.Top, vfgThis.Width / 5 * 2, vfgThis.Height / 2
        vsf.Move 15, 15, picSum.Width - 300, picSum.Height
        mclsVsf.AppendRows = True
        
        imgout(4).Move picSum.Width - imgout(4).Width, 0
    Case conMenu_Edit_Corresponding
        picColRelation.Visible = True
        picColRelation.Move picPane(1).Left, picPane(1).Top
    Case conMenu_Edit_Orders
        picColDoctor.Visible = True
        picColDoctor.Move picPane(1).Left, picPane(1).Top
    Case conMenu_Edit_PageSYN
        Call RtbSynchronous
    Case conMenu_Edit_SaveExit
        
        If SaveData Then
            DataChanged = False
            Unload Me
        End If
        
    Case conMenu_Edit_Transf_Save
        
        If SaveData Then
            DataChanged = False
        End If
    Case conMenu_Edit_NewCompute
        mintBlnReCulat = True
        If SaveData Then
            DataChanged = False
        End If
    Case conMenu_Edit_Transf_Cancle
                
        Call RefreshData
        DataChanged = False
        
    Case conMenu_File_Exit
    
        If picLabel.Visible = True Then
            picLabel.Visible = False
        ElseIf picCloumn.Visible = True Then
            picCloumn.Visible = False
        ElseIf picColDoctor.Visible = True Then
            picColDoctor.Visible = False
        ElseIf picColRelation.Visible = True Then
            picColRelation.Visible = False
        Else
            mblnOk = False
            Unload Me
        End If
        
    Case ID_TABLE_FORMATCOLWIDTH
        
        Set CbsEdit = cbsThis.FindControl(xtpControlEdit, ID_TABLE_FORMATCOLWIDTH)
        vfgThis.ColWidth(vfgThis.Col) = CbsEdit.Text
        mblnChanged = True
        
    Case ID_TABLE_FORMATROWHEIGHT
        Set CbsEdit = cbsThis.FindControl(xtpControlEdit, ID_TABLE_FORMATROWHEIGHT)
        If vfgThis.Row > 5 Then
            txtTabRowHeight.Text = CbsEdit.Text
        Else
            vfgThis.ROWHEIGHT(vfgThis.Row) = Val(CbsEdit.Text)
        End If
        
        mblnChanged = True
    Case conMenu_Help_Help
        
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
     
    End Select
End Sub


Private Function GetTab() As String
    Dim lngFixedRows As Long
    Dim strHead As String
    Dim strInput As String
    Dim strTmp As String
    
    '计算有效数据行(纸张|纸向|高|宽|上边距|下边距|左边距|右边距|行高|固定行数|标题栏字体名|标题栏字体大小|标题文本|表上项字体名|表上项字体大小|表上项文本)
    If optTabTiers(0).Value Then
        lngFixedRows = 1
    ElseIf optTabTiers(1).Value Then
        lngFixedRows = 2
    Else
        lngFixedRows = 3
    End If
    
    strHead = ""
    '获取表头信息
    With vfgThis
    For lngCount = .FixedCols To .Cols - 1
        If .RowHidden(3) = False Then strHead = strHead & "'" & lngCount - 1 & ",0," & Trim(.TextMatrix(3, lngCount)) & "," & .ColWidth(lngCount)
        If .RowHidden(4) = False Then strHead = strHead & "'" & lngCount - 1 & ",1," & Trim(.TextMatrix(4, lngCount)) & "," & .ColWidth(lngCount)
        If .RowHidden(5) = False Then strHead = strHead & "'" & lngCount - 1 & ",2," & Trim(.TextMatrix(5, lngCount)) & "," & .ColWidth(lngCount)
    Next
    strHead = Mid(strHead, 2)
    strTmp = .TextMatrix(2, 2)
    
    strInput = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex) & "|" & IIf(Me.optOrient(0).Value, 1, 2) & "|" & _
               Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips)) & "|" & _
               Val(txtTabRowHeight.Text) & "|" & lngFixedRows & "|" & Split(mstrTitleFont, ",")(0) & "|" & Split(mstrTitleFont, ",")(1) & "|" & _
               txtTitleText.Text & "|" & Split(mstrTextFont, ",")(0) & "|" & Split(mstrTextFont, ",")(1) & "|" & strTmp & "|" & strHead & "|" & _
               IIf(.ROWHEIGHT(3) < .RowHeightMin, .RowHeightMin, .ROWHEIGHT(3)) & "'" & IIf(.ROWHEIGHT(4) < .RowHeightMin, .RowHeightMin, .ROWHEIGHT(4)) & "'" & _
               IIf(.ROWHEIGHT(5) < .RowHeightMin, .RowHeightMin, .ROWHEIGHT(5))
    End With
    GetTab = strInput
               
               
               
End Function

Private Function SetFileFont(ByVal strTextFont As String, ByVal intType As Integer, ByVal strContent As String) As String
    Dim arrFont() As String
    Dim strFont As String
    Dim i As Integer
    arrFont = Split(strTextFont, ",")
    
    If UBound(arrFont) >= intType Then
        arrFont(intType) = strContent
    Else
        ReDim Preserve arrFont(UBound(arrFont) + 1)
        arrFont(UBound(arrFont)) = strContent
    End If
    For i = 0 To UBound(arrFont)
        strFont = strFont & "," & arrFont(i)
    Next
    strFont = Mid(strFont, 2)
    SetFileFont = strFont
End Function

Private Function SetEditorFont(ByVal strTextFont As String) As Boolean
    Set objFont = New StdFont
    With objFont
        .Name = Split(strTextFont, ",")(0)
        .Size = Val(Split(strTextFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strTextFont, "粗") > 0 Then .Bold = True
        If InStr(1, strTextFont, "斜") > 0 Then .Italic = True
    End With
    
    
    Select Case mintType
        Case Type_表头文本
            Set txtTitleText.Font = objFont
            vfgThis.ROWHEIGHT(1) = objFont.Size * 20 + 150
            With Me.vfgThis
                Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
                .ROWHEIGHT(1) = objFont.Size * 20 + 150
            End With
            Call SetTxtTitleSize
            mstrTitleFont = strTextFont
        Case Type_表格文本
            Set Me.vfgThis.Font = objFont
            mstrTextFont = strTextFont
        Case Type_条件文本
            With Me.vfgThis
                Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
            End With
            mstrRecordFont = strTextFont
        Case Else
            Set Me.vfgThis.Font = objFont
            mstrTextFont = strTextFont
    End Select
    SetEditorFont = True
End Function


Private Sub RtbSynchronous()
    Dim intPageAlign As Integer
    Dim strZIPHead As String, strZIPFoot As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    On Error GoTo errHand
    '将当前格式应用到所有护理文件
    
    gstrSQL = " Select 种类||'-'||编号 AS KEY From 病历文件列表 Where 种类=3 and ID<>[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件", mlngFileID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "当前只有一份护理文件，不需要执行同步功能！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("请再次确认：" & vbCrLf & "        执行该功能后，所有护理文件的页眉页脚格式将统一与当前文件设置保存一致！", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '先将当前文件的页眉页脚压缩包读出来
    If chk页码.Value = 1 Then
        If optPageAlign(0).Value Then
            intPageAlign = 1
        ElseIf optPageAlign(1).Value Then
            intPageAlign = 2
        Else
            intPageAlign = 3
        End If
    End If
    strZIPHead = ReadPageHeadFile(picFoot.Tag)
    strZIPFoot = ReadPageFootFile(picFoot.Tag)
    
    gcnOracle.BeginTrans
    blnTrans = True
    '循环写入数据库
    With rsTemp
        Do While Not .EOF
            If Not SavePageHead(!Key, strZIPHead) Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
            If Not SavePageFoot(!Key, strZIPFoot) Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
    Call zlDatabase.ExecuteProcedure("ZL_病历页面格式_页码(" & intPageAlign & ",'" & Split(picFoot.Tag, "-")(1) & "')", "更新页脚")
    gcnOracle.CommitTrans
    blnTrans = False
    '删除临时文件
    
    gobjFSO.DeleteFile strZIPHead, True
    gobjFSO.DeleteFile strZIPFoot, True
    
    MsgBox "同步成功！", vbInformation, gstrSysName
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Save
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Cancle
                
        Control.Enabled = DataChanged
    Case conMenu_Edit_Curve_AddUP, conMenu_Edit_Curve_AddBottom
        Control.Enabled = Not (optTabTiers(2).Value = True) And vfgThis.Row > 2 And vfgThis.Row < 6
    Case conMenu_Edit_Curve_DeleteRow
        Control.Enabled = Not (optTabTiers(0).Value = True) And vfgThis.Row > 2 And vfgThis.Row < 6
    Case conMenu_Edit_Curve_DeleteCol
        Control.Enabled = Not (vfgThis.Cols <= 4)
    Case ID_TABLE_FORMATCOLWIDTH, ID_TABLE_FORMATCOLWIDTH, ID_TABLE_FORMATROWHEIGHT, ID_TABLE_FORMATROWHEIGHT
        Control.Visible = False
    End Select
End Sub

Private Sub chk_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = chk.Value
    End With
    
    vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value) = Replace(vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value), IIf(chk.Value = 0, "`1", "`0"), "`" & lstColumnUsed.ItemData(lstColumnUsed.ListIndex))
    mblnChanged = True
End Sub

Private Sub chkLabelCrLf_Click()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = Me.chkLabelCrLf.Value
    End With
End Sub

Private Sub chk页码_Click()
    optPageAlign(0).Enabled = chk页码.Value
    optPageAlign(1).Enabled = chk页码.Value
    optPageAlign(2).Enabled = chk页码.Value
    lbl页码.Enabled = chk页码.Value
    cbo页码.Enabled = chk页码.Value
    If chk页码.Value = 1 Then
        If Not optPageAlign(0).Value Then
            If Not optPageAlign(1).Value Then
                If Not optPageAlign(2).Value Then optPageAlign(0).Value = True
            End If
        End If
        If cbo页码.ListIndex = -1 And cbo页码.ListCount > 0 Then cbo页码.ListIndex = cbo页码.ListCount - 1
    End If
    mblnChanged = True
End Sub

Private Sub chk重算数据行_Click()
    If chk重算数据行.Value = 1 Then
        mintBlnReCulat = True
    Else
        mintBlnReCulat = False
    End If
End Sub

Private Sub cmdOK_Click()
    Dim strTmp As String, strTmp1 As String
    Dim blnSplit As Boolean                         '多个项目时检查,如果前面的项目无后缀且后一个项目无前缀,则blnSplit=False,不允许继续
    Dim intType As Integer, intFace As Integer, intLen As Integer      '项目类型
    Dim strFaces As String                          '绑定两项目,只能是录入项目0与单选项目4;绑定两个以上项目,只能是录入项目0
    Dim strName As String                           '项目名称
    Dim intCount As Integer, arrTmp() As String
    Dim arrCol(), arrColValue()
    
    '当一列绑定2个项目时，项目之间必须存在前缀/或后缀符号加以区分
    '当一列绑定多个项目时，项目类型必须是录入型项目
    '单选与多选项目不能与其它项目一起绑定于某列
    '系统固定的项目，如签名人，日期，时间等，一列只能绑定一个
    strTemp = ""
    strTmp = ""
    strTmp1 = ""
    
    With lstColumnUsed
        If .ListCount = 1 Then
            strFaces = "0,1,2,3,4,5"
        ElseIf .ListCount = 2 Then
            strFaces = "0,4,5"
        Else
            strFaces = "0"
        End If
        
        '当前列
        For lngCount = 0 To .ListCount - 1
            strName = Mid(.List(lngCount), InStr(1, .List(lngCount), "{"))
            strName = Mid(strName, 1, InStr(1, strName, "}"))
            strTmp1 = strTmp1 & "'" & strName
        Next lngCount
        
        '所有列
        arrCol = Array()
        arrColValue = Array()
        For lngCount = 1 To vfgThis.Cols - 1
            If lngCount <> udColumnNo.Value And Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                arrTmp = Split(Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)), " ")
                For intCount = 0 To UBound(arrTmp)
                    If InStr(1, arrTmp(intCount), "`") > 0 Then
                        strName = CStr(Split(arrTmp(intCount), "`")(0))
                    Else
                        strName = CStr(arrTmp(intCount))
                    End If
                    strName = Mid(strName, InStr(1, strName, "{"))
                    strName = Mid(strName, 1, InStr(1, strName, "}"))
                    If strName <> "{}" Then
                        ReDim Preserve arrCol(UBound(arrCol) + 1)
                        arrCol(UBound(arrCol)) = lngCount
                        ReDim Preserve arrColValue(UBound(arrColValue) + 1)
                        arrColValue(UBound(arrColValue)) = strName
                    End If
                Next intCount
            End If
        Next lngCount
        
        For lngCount = 0 To .ListCount - 1
            strTemp = strTemp & Space(1) & .List(lngCount) & "`" & .ItemData(lngCount)
            strTmp = strTmp & Space(1) & .List(lngCount)
            strName = Mid(.List(lngCount), InStr(1, .List(lngCount), "{") + 1)
            strName = Mid(strName, 1, InStr(1, strName, "}") - 1)
            
            '每一列绑定的项目不能重复
            If .ListCount > 1 Then
                If UBound(Split(strTmp1, "'{" & strName & "}")) > 1 Then
                    MsgBox "当一列绑定多个项目时，项目名称不能重复！", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            '检查该列的项目是否与其他列项目重复
            For intCount = 0 To UBound(arrCol)
                If CStr(arrColValue(intCount)) = "{" & strName & "}" And Trim(strName) <> "" Then
                    MsgBox "{" & strName & "}已经在第" & CInt(arrCol(intCount)) & "存在,请检查！", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next intCount
            
            If lngCount > 0 Then
                '检查项目间是否存在分隔符
                If Not blnSplit Then
                    If Trim(Split(.List(lngCount), "{")(0) = "") Then
                        MsgBox "当一列绑定多个项目时，项目之间必须要存在前缀或后缀符号加以区分！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                '检查项目类型是否一致
                rsItems.Filter = "项目名称='" & strName & "'"
                If rsItems.RecordCount <> 0 Then
                    If Not (intType = rsItems!项目类型 And intFace = rsItems!项目表示) Then
                        MsgBox "当一列绑定多个项目时，项目的类型必须一致！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                    If rsItems!项目类型 = 1 And rsItems!项目表示 = 0 And NVL(rsItems!项目长度, 1) > 3 Then
                        MsgBox "一列最多只能绑定两个长度小于或等于3的文本项目！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                Else
                    If InStr(1, "日期,时间", strName) = 0 Then
                        MsgBox "固定项目不能与其它项目绑定在一起！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Else
                '只需要取第一个项目的属性即可
                rsItems.Filter = "项目名称='" & strName & "'"
                If rsItems.RecordCount <> 0 Then
                    intType = rsItems!项目类型
                    intFace = rsItems!项目表示
                    intLen = NVL(rsItems!项目长度, 1)
                    If .ListCount > 1 Then
                        If intType = 1 And intFace = 0 And intLen > 3 Then
                            MsgBox "一列最多只能绑定两个长度小于或等于3的文本项目！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        If intFace = 3 Then
                            MsgBox "多选项只能单独绑定！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    '一列做多允许保定两个长度<=3的文本项目
                    If .ListCount > 2 Then
                        If intType = 1 And intFace = 0 Then
                            MsgBox "一列最多只能绑定两个长度小于或等于3的文本项目！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    If .ListCount > 1 Then
                        If InStr(1, "日期,时间", strName) = 0 Then
                            MsgBox "固定项目不能与其它项目绑定在一起！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
            End If
            '固定取上一个项目的后缀
            blnSplit = Trim(Split(.List(lngCount), "}")(1) <> "")
            '检查项目表示(绑定两项目,只能是录入项目0与单选项目4;绑定两个以上项目,只能是录入项目0)
    '                If rsItems.RecordCount <> 0 Then
    '                    If InStr(1, strFaces, rsItems!项目表示) = 0 Then
    '                        If .ListCount > 2 Then
    '                            MsgBox "当一列绑定多个项目时，只能选择录入型的项目！", vbInformation, gstrSysName
    '                        Else
    '                            MsgBox "当一列绑定两个项目时，只能选择录入型或选择型的项目！", vbInformation, gstrSysName
    '                        End If
    '                        Exit Sub
    '                    End If
    '                End If
        Next
        strTemp = Trim(strTemp)
        strTmp = Trim(strTmp)
        rsItems.Filter = 0
        
        With vfgThis
            .TextMatrix(6, Me.udColumnNo.Value) = strTmp
            '根据对齐方式设置其内容
            Select Case .Cell(flexcpAlignment, 6, Me.udColumnNo.Value)
            Case 4
                .TextMatrix(7, Me.udColumnNo.Value) = " " & strTmp & " "
            Case 7
                .TextMatrix(7, Me.udColumnNo.Value) = " " & strTmp
            Case Else
                .TextMatrix(7, Me.udColumnNo.Value) = strTmp & " "
            End Select
            
            .Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value) = strTemp
            .Cell(flexcpData, 7, udColumnNo.Value, 7, udColumnNo.Value) = strTemp & " "
            
        End With
        
        '表列关系内容的设置
        Call InstallColRelation
        DataChanged = True
        picCloumn.Visible = False
    End With
    
End Sub

Private Sub ColorFillColor_pOK()

    mlngTabGridColor = ColorFillColor.COLOR
    
    With Me.vfgThis
        .GridColor = mlngTabGridColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, vbWhite, 0, 0, 1, 0, 0, 0
        .CellBorderRange 3, .FixedCols, 8, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
    End With
    
End Sub

Private Sub ColorForeColor_pOK()
    If mlngRecordColor = ColorForeColor.COLOR Then Exit Sub
    If mintType = Type_条件文本 Then
        mlngRecordColor = ColorForeColor.COLOR
        With Me.vfgThis
            .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = mlngRecordColor
        End With
    Else
        mlngTabTextColor = ColorForeColor.COLOR
        vfgThis.Cell(flexcpForeColor, 6, 1, 6, vfgThis.Cols - 1) = mlngTabTextColor
'        Me.vfgThis.ForeColor = mlngTabTextColor
       
    End If
    DataChanged = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picpaper.hWnd
    Case 2
        Item.Handle = picMain.hWnd
    End Select
End Sub

Private Sub Form_Load()
    Me.picBase(0).BackColor = Me.BackColor
    Me.picBase(1).BackColor = Me.BackColor
    Me.picBase(2).BackColor = Me.BackColor
    Me.picLabel.BackColor = Me.BackColor
    Me.picCloumn.BackColor = Me.BackColor
    Me.picpaper.BackColor = Me.BackColor
    Me.picFoot.BackColor = Me.BackColor
    CallHook cboLableSearch.hWnd
    CallHook cboItemSearch.hWnd
    
    With Me.vfgThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
    End With
'    With Me.picVBar
'        .BackColor = Me.BackColor
'        .Left = Me.ScaleWidth - .Width
'        .Top = 0: .Height = Me.Height
'    End With
    
    '---------------------------------------------------
    '默认样式设置
    Err = 0: On Error GoTo 0
    With Me.vfgThis
        .Rows = 9
        .TextMatrix(1, 0) = "标题文本"
        .TextMatrix(2, 0) = "表上标签"
        
        .TextMatrix(3, 0) = "表头单元"
        .TextMatrix(4, 0) = "表头单元"
        .TextMatrix(5, 0) = "表头单元"
        
        .TextMatrix(6, 0) = "表列内容"
        .TextMatrix(7, 0) = "条件格式"
        .TextMatrix(8, 0) = "表下标签"
        
        .ColWidth(0) = 1415
        .MergeCol(0) = True
        .RowHidden(3) = True
        .RowHidden(4) = True
        
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
'            .TextMatrix(.Rows - 1, lngCount) = "打印时间：[打印时间]                            第[页码]页"
        Next
        .MergeRow(1) = True
        .MergeRow(2) = True: .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(.Rows - 1) = True: .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(3) = True
        .MergeRow(4) = True
        .MergeRow(5) = True
        
        .Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        
        Call udTabCols_Change
        Call txtTabRowHeight_Change
'        strCurFont = Me.lblTabFont.Caption
        strCurFont = "宋体,9"
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
        End With
        
'        .ForeColor = Me.lblTabTextColor.ForeColor
        
'        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 8, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
            
        Call txtTitleText_Change
        
'        strCurFont = Me.lblTitleFont.Caption
        strCurFont = "宋体,9"
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
            
'        strCurFont = Me.lblRecordFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
'        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
        
        .RowHidden(3) = True
        .RowHidden(4) = True
        .RowHidden(.Rows - 1) = True
    End With
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    
    With Me.vfgColRelation
        .Rows = 3
        .Cols = 4
        .FixedRows = 2
        .FixedCols = 2
        .MergeCells = flexMergeNever
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeRow(0) = True
        .TextMatrix(0, 0) = "汇总项目"
        .TextMatrix(0, 1) = "汇总项目"
        .TextMatrix(0, 2) = "对照项目"
        .TextMatrix(0, 3) = "对照项目"
        .TextMatrix(1, 0) = "列号"
        .TextMatrix(1, 1) = "名称"
        .TextMatrix(1, 2) = "列号"
        .TextMatrix(1, 3) = "名称"
        .ColKey(0) = "汇总列号"
        .ColKey(1) = "汇总名称"
        .ColKey(2) = "对照列号"
        .ColKey(3) = "对照名称"
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
    End With
    
    
    With Me.vsfColDoctor
        .Rows = 2
        .Cols = 3
        .FixedRows = 2
        .FixedCols = 2
        .MergeCells = flexMergeNever
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeRow(0) = True
        
        .TextMatrix(0, 0) = "项目"
        .TextMatrix(0, 1) = "项目"
        .TextMatrix(0, 2) = "对照项目"
        .TextMatrix(1, 0) = "列号"
        .TextMatrix(1, 1) = "名称"
        .TextMatrix(1, 2) = "名称"
        .ColKey(0) = "项目列号"
        .ColKey(1) = "项目名称"
        .ColKey(2) = "对照名称"
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
    End With
    
    Set mclsVsf = New clsVsf
    With mclsVsf
        
        Call .Initialize(Me.Controls, vsf, True, True)
        Call .ClearColumn

        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[图标]", False)
        
        Call .AppendColumn("时段名称", 2100, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("开始时间", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("结束时间", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
        
        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(mclsVsf.ColIndex("时段名称"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("开始时间"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("结束时间"), True, vbVsfEditText)
                
        .AppendRows = True
    End With
    
    vsf.ColHidden(0) = True
    
    Dim objPane As Pane
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitMenuBar
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False '实时拖动
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis

    Set objPane = dkpMan.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "打印设置": objPane.Options = PaneNoCaption
'    dkpMan.cl
    Set objPane = dkpMan.CreatePane(2, 100, 200, DockBottomOf, objPane): objPane.Title = "设计": objPane.Options = PaneNoCaption
    
    dkpMan.FindPane(1).Hidden = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    rtbHead.Width = picPane(1) - picFoot.Width - 30
    rtbFoot.Width = picPane(1) - rtbHead.Width - 30
    
    Call SetPaneRange(dkpMan, 1, 15, 200, Me.ScaleWidth, 100)
    dkpMan.RecalcLayout
    SetpicBase
    Call SetTxtTitleSize
End Sub

Private Sub SetTxtTitleSize()
    Dim lngCol As Long
    Dim longWidth As Long
    
    longWidth = 0
    For lngCol = vfgThis.LeftCol To vfgThis.Cols - 1
'        If longWidth + vfgThis.ColWidth(lngCol) > picPane(1).Width - picHead.Width Then Exit For
        longWidth = vfgThis.ColWidth(lngCol) + longWidth
    Next
    
    If longWidth > rtbFoot.Width Then longWidth = rtbFoot.Width - 300
    With txtTitleText
        .Left = vfgThis.ColWidth(0) + 30
        .Top = IIf(vfgThis.ROWHEIGHT(0) > vfgThis.RowHeightMin, vfgThis.ROWHEIGHT(0), vfgThis.RowHeightMin) + vfgThis.Top + 20
        .Height = IIf(vfgThis.ROWHEIGHT(1) > vfgThis.RowHeightMin, vfgThis.ROWHEIGHT(1), vfgThis.RowHeightMin)
        .Width = longWidth
    End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
    If DataChanged Then
        Cancel = (MsgBox("更改后的设计必须保存后才生效，是否放弃保存？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    CallUnhook cboItemSearch.hWnd
    CallUnhook cboLableSearch.hWnd
    
    DataChanged = False
    
    Call SaveWinState(Me, App.ProductName)
End Sub


Private Sub imgColDelete_Click()
    If lstColumnUsed.ListIndex = -1 Then Exit Sub
    lstColumnUsed.RemoveItem lstColumnUsed.ListIndex
    If lstColumnUsed.ListCount > 0 Then
        lstColumnUsed.ListIndex = 0
    Else
        lstColumnUsed.ListIndex = -1
        Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
        Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
    End If
    
    chk.Enabled = True
End Sub

Private Sub imgLabDelete_Click()
    With lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .RemoveItem .ListIndex
        If .ListCount > 0 Then
            .ListIndex = 0
        Else
            .ListIndex = -1
            Me.imgLabDelete.Enabled = False
            Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
            Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
        End If
    End With
End Sub

Private Sub imgout_Click(Index As Integer)
    Select Case Index
    Case 0
        picCloumn.Visible = False
    Case 1
        picLabel.Visible = False
    Case 2
        picColDoctor.Visible = False
    Case 3
        picColRelation.Visible = False
    Case 4
        picSum.Visible = False
    End Select
End Sub

Private Sub ImgUpdown_Click(Index As Integer)
    If Val(ImgUpdown(Index).Tag) = 0 Then
        Set ImgUpdown(Index) = Img.ListImages("BaseUp").Picture
        picBase(Index).Visible = False
        ImgUpdown(Index).Tag = 1
        Call SetpicBase
    Else
        Set ImgUpdown(Index) = Img.ListImages("BaseDown").Picture
        picBase(Index).Visible = True
        ImgUpdown(Index).Tag = 0
        Call SetpicBase
    End If
End Sub

Private Sub SetpicBase()
    picSize(0).Move 15, 15
    picBase(0).Move 15, picSize(0).Top + picSize(0).Height + 15
    If picBase(0).Visible Then
        picSize(1).Move 15, picBase(0).Top + picBase(0).Height + 15
    Else
        picSize(1).Move 15, picSize(0).Top + picSize(0).Height + 15
    End If
    picBase(1).Move 15, picSize(1).Top + picSize(1).Height + 15
    If picBase(1).Visible Then
        picSize(2).Move 15, picBase(1).Top + picBase(1).Height + 15
    Else
        picSize(2).Move 15, picSize(1).Top + picSize(1).Height + 15
    End If
    picBase(2).Move 15, picSize(2).Top + picSize(2).Height + 15
    
End Sub

Private Sub lstColumnUsed_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        Me.txtColumnPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
        Me.txtColumnPostfix.Text = Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "}") + 1)
        chk.Value = .ItemData(.ListIndex)
        imgColDelete.Move 2745, .ListIndex * (.FontSize * 20) + .Top
    End With
End Sub

Private Sub lstLabelUsed_Click()
    Dim lngHeight As Long
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        Me.chkLabelCrLf.Value = IIf(.ItemData(.ListIndex) = 0, vbUnchecked, vbChecked)
        Me.txtLabelPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
        imgLabDelete.Move 1980, (.ListIndex - .TopIndex) * (lstLabelUsed.FontSize * 20) + lstLabelUsed.Top
    End With
End Sub

Private Sub lstLabelUsed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y > 0 Then dkpMan.FindPane(1).Hidden = True
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf
        Cancel = (.TextMatrix(Row, 1) = "" Or .TextMatrix(Row, 2) = "" Or .TextMatrix(Row, 3) = "")
    End With
End Sub

Private Sub optOrient_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub optPageAlign_Click(Index As Integer)
    mblnChanged = True
End Sub

Private Sub optTabTiers_Click(Index As Integer)
    With vfgThis
        If optTabTiers(0).Value Then
            If .Row = 5 Or .Row = 4 Then .Row = 3
            .RowHidden(3) = False
            .RowHidden(4) = True
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 1
        ElseIf optTabTiers(1).Value Then
            If .Row = 5 Then .Row = 4
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 2
        Else
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = False
            udHeadRow.Min = 1
            udHeadRow.Max = 3
        End If

    End With
    DataChanged = True
End Sub

Private Sub optTabTiers_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picBase_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Y > 0 Then dkpMan.FindPane(1).Hidden = True
End Sub

Private Sub picDown_Click()
    Dim lngIndex As Long
    lngIndex = lstLabelUsed.ListIndex
        If lngIndex + 1 < lstLabelUsed.ListCount Then
        lstLabelUsed.AddItem lstLabelUsed.Text, lngIndex + 2
        lstLabelUsed.RemoveItem lngIndex
        lstLabelUsed.Selected(lngIndex + 1) = True
    End If
    mblnChanged = True
End Sub

Private Sub picImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    If Index = 0 Then
        strInfo = "温馨提示："
        strInfo = strInfo & vbCrLf & "1.只针对绑定了一个汇总项目的列。"
        strInfo = strInfo & vbCrLf & "2.对照项目只能选择绑定了一个项目的列，并且绑定的项目为长度小于100的文本项目或单选项目。"
        strInfo = strInfo & vbCrLf & "3.点击对照项目的名称列可选择项目。"
    ElseIf Index = 1 Then
        strInfo = "温馨提示："
        strInfo = strInfo & vbCrLf & "1.只针对绑定一个非固定项目的列"
        strInfo = strInfo & vbCrLf & "2.对照项目绑定之后通过入量导入,自动添加所有绑定医嘱的项目,绑定的医嘱项目要与设定的项目类型相同(用量和总给予量需绑定数值项目)"
        strInfo = strInfo & vbCrLf & "3.点击对照项目的名称列可选择项目。"
    End If
    Call ZLCommFun.ShowTipInfo(picImg(Index).hWnd, strInfo, True)
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    picPane(0).Move 30, 0, picPane(0).Width, picMain.Height
    picPane(1).Move picPane(0).Left + picPane(0).Width, picPane(0).Top, picMain.Width - picPane(0).Width - 30, picMain.Height
    
End Sub

Private Sub cmdLabOK_Click()
    With lstLabelUsed
        Dim intCrLf As Integer
        intCrLf = 0
        strTemp = ""
        For lngCount = 0 To .ListCount - 1
            If .ItemData(lngCount) <> 0 Then intCrLf = intCrLf + 1
            strTemp = strTemp & Space(1) & IIf(.ItemData(lngCount) = 0, "", vbCrLf) & .List(lngCount)
        Next
        strTemp = Trim(strTemp)
        For lngCount = Me.vfgThis.FixedCols To Me.vfgThis.Cols - 1
            Me.vfgThis.TextMatrix(2, lngCount) = strTemp
        Next
        Me.vfgThis.ROWHEIGHT(2) = Me.vfgThis.FontSize * 20 * (intCrLf + 1) + 150
        
        DataChanged = True
        picLabel.Visible = False
    End With
End Sub


Private Sub picPane_Resize(Index As Integer)
    Select Case Index
    Case 1
        picHead.Move 15, 15
        vfgThis.Move 15, picHead.Height, picPane(Index).Width - 30, picPane(Index).Height - picFoot.Height * 3 - 100
        picFoot.Move 15, vfgThis.Top + vfgThis.Height, picFoot.Width, picFoot.Height
        rtbHead.Move picFoot.Left + picFoot.Width, picHead.Top, vfgThis.Width - picFoot.Width - 30, picHead.Height
        rtbFoot.Move picFoot.Left + picFoot.Width, picFoot.Top, vfgThis.Width - picFoot.Width - 30, picFoot.Height
        picFoot页码.Move 15, picFoot.Top + picFoot.Height, picPane(1).Width, picFoot页码.Height
        Call SetTxtTitleSize
    
    End Select
End Sub

Private Sub picpaper_GotFocus()
    On Error Resume Next
        
        Printer.PaperSize = cboPaperKind.ItemData(cboPaperKind.ListIndex)
        Printer.Orientation = IIf(optOrient(0).Value, 1, 2)
        If Printer.PaperSize = 256 Then
            Call SetCustonPager(Me.hWnd, Int(Me.ScaleY(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)), Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)))
        End If
       
End Sub

Private Sub picUP_Click()
    Dim lngIndex As Long
    lngIndex = lstLabelUsed.ListIndex '记录当前序号
    If lngIndex = -1 Then Exit Sub '加这句，判断是否选择
    If lngIndex - 1 > -1 Then '判断是否最前
        lstLabelUsed.AddItem lstLabelUsed.Text, lngIndex - 1 '在当前项的上一条记录前插入一条记录
        lstLabelUsed.RemoveItem lngIndex + 1 '删除原记录
        lstLabelUsed.Selected(lngIndex - 1) = True '选中移动后记录
    End If
    mblnChanged = True
End Sub

Private Sub rtbFoot_Change()
    mblnChanged = True
End Sub


Private Sub rtbFoot_GotFocus()
    mblnRTBFoot = True
    mintType = Type_页眉页脚
End Sub

Private Sub rtbFoot_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button = 2 Then
        vfgThis.Tag = vfgThis.Col
        Set cbrPopupBar = cbsThis.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_MarkMap, "插入图片")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_PageSYN, "页面格式同步")
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub rtbHead_Change()
    mblnChanged = True
End Sub

Private Sub rtbHead_GotFocus()
    mblnRTBFoot = False
    mintType = Type_页眉页脚
End Sub

Private Sub rtbHead_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button = 2 Then
        vfgThis.Tag = vfgThis.Col
        Set cbrPopupBar = cbsThis.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_MarkMap, "插入图片")
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_PageSYN, "页面格式同步")
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub txtColumnPostfix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "}")) & Me.txtColumnPostfix.Text
    End With
End Sub

Private Sub txtColumnPostfix_GotFocus()
    Me.txtColumnPostfix.SelStart = 0: Me.txtColumnPostfix.SelLength = 4000
    Call ZLCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPostfix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtColumnPrefix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtColumnPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtColumnPrefix_GotFocus()
    Me.txtColumnPrefix.SelStart = 0: Me.txtColumnPrefix.SelLength = 4000
    Call ZLCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtHeadRow_Change()
    DataChanged = True
End Sub

Private Sub txtHeadText_Change()
    Dim strInput As String
    Dim lngRow As Long, lngCol As Long
    Dim blnExist As Boolean
    
    strInput = Trim(Me.txtHeadText.Text)
    '检查,如果有相邻的四个单元格的值相同,则不允许设置(有可能需要和左上,左下,右上,右下进行检查)
    lngRow = udHeadRow.Value + 2
    lngCol = Me.udHeadCol.Value
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCol) = strInput
    
    If lngRow <= 4 Then
        If (vfgThis.TextMatrix(3, lngCol) = vfgThis.TextMatrix(4, lngCol) And vfgThis.TextMatrix(3, lngCol) <> "") Then
            If lngCol > 1 Then
                '左上
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '右上
                If vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow, lngCol + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCol + 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
    If lngRow >= 4 Then
        If (vfgThis.TextMatrix(4, lngCol) = vfgThis.TextMatrix(5, lngCol) And vfgThis.TextMatrix(4, lngCol) <> "") Then
            If lngCol > 1 Then
                '左下
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '右下
                If vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow, lngCol + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCol + 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
Limit:
    If blnExist Then strInput = strInput & "_1"

WriteIt:
    Me.txtHeadText.Text = strInput
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCol) = strInput
'    vfgThis.AutoSize 0, vfgThis.Cols - 1
    Call cmdLabOK_Click
    DataChanged = True
End Sub

Private Sub txtLabelPrefix_Change()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtLabelPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtLabelPrefix_GotFocus()
    Me.txtLabelPrefix.SelStart = 0: Me.txtLabelPrefix.SelLength = 4000
    Call ZLCommFun.OpenIme(True)
End Sub

Private Sub txtLabelPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtMarjin_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txtRecordFrom_Change()
    DataChanged = True
End Sub

Private Sub txtRecordTo_Change()
    DataChanged = True
End Sub

Private Sub txtTabCols_Change()
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_Change()
    Me.vfgThis.RowHeightMin = Val(Me.txtTabRowHeight.Text)
    Call SetTxtTitleSize
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_GotFocus()
    mintType = Type_表头文本
    Me.txtTabRowHeight.SelStart = 0: Me.txtTabRowHeight.SelLength = 100
    Call ZLCommFun.OpenIme(False)
End Sub

Private Sub txtTabRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtTitleText_Change()
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(1, lngCount) = Trim(Me.txtTitleText.Text)
        Next
    End With
    With Me.vfgThis
        .ROWHEIGHT(1) = txtTitleText.FontSize * 20 + 150
        vfgThis.AutoSize 0, vfgThis.Cols - 1
    End With
    
    DataChanged = True
End Sub

Private Sub txtTitleText_GotFocus()
    Dim CbsCombo As CommandBarComboBox
    Me.txtTitleText.SelStart = 0: Me.txtTitleText.SelLength = 4000
    Call ZLCommFun.OpenIme(True)
    Set CbsCombo = cbsThis.FindControl(xtpControlComboBox, ID_FORMAT_FONTNAME)
    CbsCombo.Text = txtTitleText.FontName
    Set CbsCombo = cbsThis.FindControl(xtpControlComboBox, ID_FORMAT_FONTSIZE)
    CbsCombo.Text = IIf(InStr(1, txtTitleText.FontSize, ".5") > 0, Val(txtTitleText.FontSize), Round(Val(txtTitleText.FontSize)))
    mintType = 0
End Sub

Private Sub txtTitleText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call ZLCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtTitleText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtTitleText.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtTitleText.hWnd, GWL_WNDPROC, AddressOf WndMessage)
        Set cbrPopupBar = cbsThis.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddLeft, "在左侧新增列"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddRight, "在右侧新增列"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_DeleteCol, "删除列"): cbrPopupItem.IconId = conMenu_Edit_Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddUP, "在上方新增行"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddBottom, "在下方新增行"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_DeleteRow, "删除行"): cbrPopupItem.IconId = conMenu_Edit_Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_BuddyDouble, "绑定列"): cbrPopupItem.IconId = conMenu_Edit_Curve_BuddyDouble
        cbrPopupBar.ShowPopup
        Call SetWindowLong(txtTitleText.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub txtTitleText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And glngTXTProc <> 0 Then
        glngTXTProc = GetWindowLong(txtTitleText.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtTitleText.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub udColumnNo_Change()
     Me.lstColumnUsed.Clear
'    strTemp = Trim(Me.vfgThis.TextMatrix(5, Me.udColumnNo.Value))
    strTemp = vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value)
    
    If strTemp = "" Or strTemp = "{}`0" Then
        Me.lstColumnUsed.ListIndex = -1
        Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
        Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
    Else
        Dim aryCol() As String
        aryCol = Split(strTemp, Space(1))
        For lngCount = 0 To UBound(aryCol)
            If InStr(aryCol(lngCount), "`") > 0 Then
                Me.lstColumnUsed.AddItem Mid(aryCol(lngCount), 1, InStr(aryCol(lngCount), "`") - 1)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Mid(aryCol(lngCount), InStr(aryCol(lngCount), "`") + 1))
            Else
                Me.lstColumnUsed.AddItem aryCol(lngCount)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = 0
            End If
        Next
        Me.lstColumnUsed.ListIndex = 0
        Me.imgColDelete.Enabled = True
        Me.txtColumnPrefix.Enabled = True
        Me.txtColumnPostfix.Enabled = True
        
        chk.Enabled = True
        
    End If
End Sub

Private Sub udHeadCol_Change()
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    txtHeadText.Text = vfgThis.TextMatrix(udHeadRow.Value + 2, udHeadCol.Value)

    DataChanged = blnSvrChanged
End Sub

Private Sub udHeadRow_Change()
    Call udHeadCol_Change
End Sub

Private Sub udRecordFrom_Change()
    If Me.udRecordFrom.Value > Me.udRecordTo.Value Then
        Me.lblRecordTo.Caption = "次日" & Space(7) & "点"
    Else
        Me.lblRecordTo.Caption = "当日" & Space(7) & "点"
    End If
End Sub

Private Sub udRecordTo_Change()
     Call udRecordFrom_Change
End Sub

Private Sub udTabCols_Change()
    If vfgThis.Cols = Me.udTabCols.Value + 1 Then Exit Sub
    Me.vfgThis.Cols = Me.udTabCols.Value + 1
    Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
    Me.udHeadCol.Max = Me.udTabCols.Value
    If Val(Me.txtHeadCol.Text) > Me.udHeadCol.Max Then Me.txtHeadCol.Text = Me.udHeadCol.Max
    Me.udColumnNo.Max = Me.udTabCols.Value
    If Val(Me.txtColumnNo.Text) > Me.udColumnNo.Max Then Me.txtColumnNo.Text = Me.udColumnNo.Max
    
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .TextMatrix(1, lngCount) = .TextMatrix(1, .FixedCols)
            .TextMatrix(2, lngCount) = .TextMatrix(2, .FixedCols)
            .TextMatrix(.Rows - 1, lngCount) = .TextMatrix(.Rows - 1, .FixedCols)
        Next
        '.Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, mlngTabGridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 8, .Cols - 1, mlngTabGridColor, 1, 1, 1, 1, 1, 1
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = mlngRecordColor
    End With
End Sub

Private Sub vfgColRelation_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = vfgColRelation.ColIndex("对照名称") Then
        vfgColRelation.TextMatrix(Row, vfgColRelation.ColIndex("对照列号")) = vfgColRelation.Cell(flexcpText, Row, Col)
    End If
    vfgColRelation.TextMatrix(Row, Col) = vfgColRelation.EditText
End Sub

Private Sub vfgColRelation_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    With vfgColRelation
        If Row >= .FixedRows And Col = .ColIndex("对照名称") And .ComboIndex >= 0 Then
            FinishEdit = True
        End If
    End With
End Sub

Private Sub vfgColRelation_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgColRelation
        If Row >= .FixedRows And Col = .ColIndex("对照名称") Then
            If .ColComboList(Col) = "" Then
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vfgColRelation_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long
    Dim blnCancle As Boolean
    
    With vfgColRelation
        If Row >= .FixedRows And Col = .ColIndex("对照名称") And .ColComboList(.ColIndex("对照名称")) <> "" Then
            If Trim(vfgColRelation.EditText) = Trim(.TextMatrix(Row, .ColIndex("对照名称"))) Then Exit Sub
            If .ComboIndex < 0 Or Trim(vfgColRelation.EditText) = "" Then
                .TextMatrix(Row, .ColIndex("对照列号")) = ""
                DataChanged = True
                Exit Sub
            End If
            For lngRow = .FixedRows To .Rows - 1
                If lngRow <> Row Then
                    If .TextMatrix(lngRow, .ColIndex("对照名称")) = vfgColRelation.EditText Then
                        blnCancle = True
                        Exit For
                    End If
                End If
            Next lngRow
            If blnCancle = True Then
                MsgBox "您选择的对照项目已经与其它汇总项目设置了对照关系，请重新选择!", vbInformation, gstrSignName
                Cancel = True
            Else
                .TextMatrix(Row, .ColIndex("对照名称")) = vfgColRelation.EditText
                DataChanged = True
            End If
        End If
    End With
End Sub

Private Sub vfgThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnSvrChanged As Boolean
    Dim CbsEdit As CommandBarEdit
    Dim cbrControl As CommandBarControl
    Dim CbsCombo As CommandBarComboBox
    Dim lngFontSize As Long
    
    
    If vfgThis.Row = 7 Then
        mintType = Type_条件文本
        lngFontSize = vfgThis.Cell(flexcpFontSize, 7, vfgThis.FixedCols, 7, vfgThis.Cols - 1)
    Else
        mintType = Type_表格文本
        lngFontSize = vfgThis.FontSize
    End If
    vfgThis.Tag = Val(NewCol)
    Set CbsCombo = cbsThis.FindControl(xtpControlComboBox, ID_FORMAT_FONTNAME)
    CbsCombo.Text = vfgThis.FontName
    Set CbsCombo = cbsThis.FindControl(xtpControlComboBox, ID_FORMAT_FONTSIZE)
    CbsCombo.Text = IIf(InStr(1, lngFontSize, ".5") > 0, lngFontSize, Round(lngFontSize))
    
    blnSvrChanged = DataChanged
    
    Me.udHeadCol.Value = NewCol
    Me.udColumnNo.Value = NewCol
    If NewRow >= 3 And NewRow <= 5 Then udHeadRow.Value = NewRow - 2
    
    Dim intAlign As Integer
    Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Right_Alignment)
    cbrControl.Checked = False
    Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Center_Alignment)
    cbrControl.Checked = False
    Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Left_Alignment)
    cbrControl.Checked = False
    
    
    
    
    Select Case vfgThis.Cell(flexcpAlignment, 6, NewCol)
    Case Is >= 9
        intAlign = 0
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Left_Alignment)
    Case Is >= 6    '右
        intAlign = 2
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Right_Alignment)
    Case Is >= 2    '中
        intAlign = 1
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Center_Alignment)
    Case Is <= 2    '左
        intAlign = 0
        Set cbrControl = cbsThis.FindControl(xtpControlButton, conMenu_Edit_Left_Alignment)
    End Select
    cbrControl.Checked = True
    mintTableAlign = intAlign
    
    
    If NewRow > 2 Then
        Set CbsEdit = cbsThis.FindControl(xtpControlEdit, ID_TABLE_FORMATROWHEIGHT)
        CbsEdit.Text = vfgThis.CellHeight
    End If
    Set CbsEdit = cbsThis.FindControl(xtpControlEdit, ID_TABLE_FORMATCOLWIDTH)
    CbsEdit.Text = vfgThis.CellWidth
    DataChanged = blnSvrChanged
    If picCloumn.Visible = True And vfgThis.Row <> 2 Then
        picCloumn.Left = vfgThis.Left + vfgThis.CellLeft + vfgThis.CellWidth
        If picCloumn.Left + picCloumn.Width > picPane(1).Width Then
            picCloumn.Left = vfgThis.Left + vfgThis.CellLeft - picCloumn.Width
        End If
    End If
End Sub

Private Sub vfgThis_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)

    If NewTopRow = 1 Then
        txtTitleText.Visible = True
    Else
        txtTitleText.Visible = False
    End If
    
    Call SetTxtTitleSize
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
     Call SetTxtTitleSize
    DataChanged = True
End Sub

Private Sub GetrtbObject()
    If mblnRTBFoot Then
        Set rtbThis = rtbFoot
    Else
        Set rtbThis = rtbHead
    End If
End Sub

Private Sub vfgThis_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    Dim lngCol As Long
    If Col >= Position Then
        For lngCol = Position To Col
            vfgThis.TextMatrix(0, lngCol) = lngCol
        Next lngCol
    Else
        For lngCol = Col To Position
            vfgThis.TextMatrix(0, lngCol) = lngCol
        Next lngCol
    End If
    Call InstallColRelation
    DataChanged = True
End Sub

Private Sub vfgThis_DblClick()
    If vfgThis.Row = 6 Or vfgThis.Row = 7 Then
        picCloumn.Visible = True
        cboItemSearch.Text = ""
        If InStr(2, vfgThis.TextMatrix(vfgThis.Row, vfgThis.Col), "{") > 0 Then
            picCloumn.Tag = 2
        Else
            picCloumn.Tag = 1
        End If
        picCloumn.Top = vfgThis.Top + vfgThis.CellTop
        picCloumn.Left = vfgThis.Left + vfgThis.CellLeft + vfgThis.CellWidth
        If picCloumn.Left + picCloumn.Width > picPane(1).Width Then
            picCloumn.Left = vfgThis.Left + vfgThis.CellLeft - picCloumn.Width
        End If
        udColumnNo.Value = vfgThis.Col
        txtColumnNo.Text = vfgThis.Col
    ElseIf vfgThis.Row = 2 Then
        picLabel.Visible = True
        picLabel.Top = vfgThis.Top + vfgThis.CellTop + vfgThis.ROWHEIGHT(2)
        picLabel.Left = vfgThis.Left + vfgThis.CellLeft
    End If
End Sub

Private Sub vfgThis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call imgLabDelete_Click
    End If
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    
    If Button = 2 Then
        vfgThis.Tag = vfgThis.Col
        Set cbrPopupBar = cbsThis.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddLeft, "在左侧新增列"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddRight, "在右侧新增列"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_DeleteCol, "删除列"): cbrPopupItem.IconId = conMenu_Edit_Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddUP, "在上方新增行"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_AddBottom, "在下方新增行"): cbrPopupItem.IconId = conMenu_Edit_Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_DeleteRow, "删除行"): cbrPopupItem.IconId = conMenu_Edit_Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu_Edit_Curve_BuddyDouble, "绑定列"): cbrPopupItem.IconId = conMenu_Edit_Curve_BuddyDouble
        cbrPopupBar.ShowPopup
    End If
End Sub

Private Sub LoadPaper()
    '---------------------------------------------------
    '功能：装入当前打印机可用的纸张
    '---------------------------------------------------
    Dim intCurPaper As Integer
    Dim strDevice As String
    Dim i As Integer, j As Integer
    Dim strPaper As String * 1000
    
    
     With Me.cboPaperKind
        .AddItem GetPaperName(256)
        .ItemData(.NewIndex) = 256
        If Not ExistsPrinter Then
            .ListIndex = 0
            .Enabled = False
            Exit Sub
        End If
        
        strDevice = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", Printer.DeviceName)
        For lngCount = 0 To Printers.Count - 1
            If Printers(lngCount).DeviceName = strDevice Then
                Set Printer = Printers(lngCount)
                Exit For
            End If
        Next
        Me.lblPrinter.Caption = "当前打印机: " & Printer.DeviceName
        lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaper, 0)
        intCurPaper = Printer.PaperSize
        .Enabled = True
'        For lngCount = 1 To 41
'            Err = 0: On Error Resume Next
'            Printer.PaperSize = lngCount
'            Err = 0: On Error GoTo 0
'            If Printer.PaperSize = lngCount Then
'                .AddItem GetPaperName(lngCount)
'                .ItemData(.NewIndex) = lngCount
'                If lngCount = intCurPaper Then .ListIndex = .NewIndex
'            End If
'        Next
        
        For i = 1 To lngCount
            j = Asc(Mid(strPaper, i * 2, 1)) * 256# + Asc(Mid(strPaper, i * 2 - 1, 1))
            If j >= 1 And j <= 41 Then '只列出标准支持的纸张
                .AddItem GetPaperName(j)
                .ItemData(.NewIndex) = j
                If j = intCurPaper Then .ListIndex = .NewIndex
            End If
        Next
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub



Private Function OverRun() As Boolean
    Dim lngPageMargin As Long    '边距
    Dim lngPageWidth As Long     '纸张宽度
    Dim lngTrimSize As Long      '纸张实际宽度
    Dim dblTableWidth As Double     '表格宽度
    Dim lngCol As Long, lngCols As Long
    '检查表格的总体宽度是否超过了右边距
    
    lngCols = Me.vfgThis.Cols - 1
    For lngCol = 1 To lngCols
        dblTableWidth = dblTableWidth + Val(vfgThis.ColWidth(lngCol))
    Next
    
    '检查表格宽度是否超出页面有效打印范围
    Printer.Orientation = IIf(optOrient(0).Value, 1, 2)
    If cboPaperKind.ItemData(cboPaperKind.ListIndex) = 256 Then
        Call SetCustonPager(Me.hWnd, Int(Me.ScaleY(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)), Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)))
    Else
        Printer.PaperSize = cboPaperKind.ItemData(cboPaperKind.ListIndex)
    End If
    lngPageWidth = Printer.ScaleWidth
    lngPageMargin = Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips)) + Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips))
    lngTrimSize = lngPageWidth - lngPageMargin - 100
    If dblTableWidth > lngTrimSize Then
        OverRun = True
        Exit Function
    End If
End Function

Private Function PageHeadTest() As Boolean
    '超过上边距返回假
    Dim fr As FORMATRANGE           '格式化的文本范围
    Dim rcDrawTo As RECT            '目标文字区域
    Dim rcPage As RECT              '目标页面区域
    Dim gTargetDC As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
'    Dim lngOffsetWidth As Long
'    Dim lngOffsetHeight As Long
    Dim lngNextPos As Long, lngLen As Long, lngTMP As Long, lngPageCount As Long
    Dim blnTrue As Boolean
    
    lngLen = lstrlen(rtbHead.Text)
    'printer.Duplex = vbPRDPHorizontal
    'printer.ScaleMode = vbTwips
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
'    lngOffsetWidth = Printer.ScaleWidth
'    lngOffsetHeight = Printer.ScaleHeight
    
    gTargetDC = hDC
    With rcPage
        .Left = 0
        .Top = 0
        .Right = Printer.Width
        .Bottom = Printer.Height
    End With
    With rcDrawTo
        .Left = lngOffsetLeft
        .Top = lngOffsetTop
        .Right = Printer.Width - lngOffsetLeft
        .Bottom = Printer.ScaleX(txtMarjin(0).Text, vbMillimeters, vbTwips)
        '46251,刘鹏飞,2012-09-11
        If Val(cbo页码.ItemData(cbo页码.ListIndex)) = 1 Or Val(cbo页码.ItemData(cbo页码.ListIndex)) = 2 Then
            .Bottom = .Bottom - Printer.TextHeight("1")
            blnTrue = True
        End If
    End With
    With fr
        .hDC = Printer.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    
    Do
        lngNextPos = SendMessage(rtbHead.hWnd, EM_FORMATRANGE, 0, fr)
        
        lngPageCount = lngPageCount + 1             ' 页数＋1
        '记录分页信息
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          '实际打印高度
        AllPages(lngPageCount).Start = lngTMP
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTMP Or lngNextPos >= lngLen Then Exit Do      ' 完成所有页面的分页
        lngTMP = lngNextPos
    Loop
    Call SendMessage(rtbHead.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    If fr.rc.Bottom > rcDrawTo.Bottom Or lngPageCount > 1 Then
        MsgBox "设计的页眉内容" & IIf(blnTrue = True, "加页码高度", "") & "超过了上边距！", vbInformation, gstrSysName
        Exit Function
    End If
    PageHeadTest = True
End Function

Private Function ReadPageHead(objHead As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Private Function ReadPageFoot(objFoot As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '读取文件
        gobjFSO.DeleteFile strFile, True      '删除临时文件
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

Private Function ReadPageHeadFile(ByVal StrKey As String) As String
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageHeadFile = strZip
    End If
End Function

Private Function ReadPageFootFile(ByVal StrKey As String) As String
'################################################################################################################
'## 功能：  读取页面图片
'## 参数：  病历种类-页面编号
'## 返回：  返回获得的图片变量。
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageFootFile = strZip
    End If
End Function


'################################################################################################################
'## 功能：  在压缩文件相同目录释放产生解压文件
'## 参数：  strZipFile     :压缩文件
'## 返回：  解压文件名，失败则返回零长度""
'################################################################################################################
Private Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function

Private Sub vfgThis_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y > 0 Then dkpMan.FindPane(1).Hidden = True
End Sub

Private Sub vfgThis_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = vfgThis.Rows - 1 Or Row > 5 Or Row < 3 Then
        Cancel = True
    Else
        vfgThis.Editable = flexEDKbdMouse
    End If
End Sub

Private Sub vfgThis_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vfgThis.EditText <> vfgThis.TextMatrix(Row, Col) Then
        vfgThis.TextMatrix(Row, Col) = vfgThis.EditText
        vfgThis.Cell(flexcpAlignment, 0, vfgThis.FixedCols, vfgThis.FixedRows, vfgThis.Cols - 1) = flexAlignCenterCenter
        DataChanged = True
    End If
End Sub


Public Sub CallHook(ByVal hWnd As Long)
    glngTXTProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndMessage)
End Sub

Public Sub CallUnhook(ByVal hWnd As Long)
    Dim lngReturnValue As Long
    lngReturnValue = SetWindowLong(hWnd, GWL_WNDPROC, glngTXTProc)
End Sub
