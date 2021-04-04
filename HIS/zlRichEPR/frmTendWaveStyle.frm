VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendWaveStyle 
   AutoRedraw      =   -1  'True
   Caption         =   "体温单设置"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16200
   Icon            =   "frmTendWaveStyle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   16200
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picTable 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   8910
      Left            =   645
      ScaleHeight     =   8910
      ScaleWidth      =   6885
      TabIndex        =   84
      Top             =   420
      Width           =   6885
      Begin MSComCtl2.FlatScrollBar vsbTab 
         Height          =   8910
         Left            =   6630
         TabIndex        =   175
         Top             =   0
         Visible         =   0   'False
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   15716
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin VB.PictureBox picBabyTable 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4515
         Left            =   30
         ScaleHeight     =   4515
         ScaleWidth      =   6885
         TabIndex        =   85
         Top             =   4080
         Visible         =   0   'False
         Width           =   6885
         Begin MSComCtl2.UpDown udBabyCol 
            Height          =   300
            Left            =   1620
            TabIndex        =   174
            Top             =   555
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   3
            BuddyControl    =   "txtBabyCol"
            BuddyDispid     =   196623
            OrigLeft        =   1665
            OrigTop         =   495
            OrigRight       =   1920
            OrigBottom      =   915
            Max             =   60
            Min             =   3
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtHeadRow 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3795
            MaxLength       =   1
            TabIndex        =   168
            Text            =   "2"
            Top             =   585
            Width           =   360
         End
         Begin VB.TextBox txtHeadCol 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3795
            MaxLength       =   2
            TabIndex        =   167
            Text            =   "1"
            Top             =   1020
            Width           =   345
         End
         Begin VB.TextBox txtHeadText 
            Height          =   885
            Left            =   4755
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   165
            Top             =   480
            Width           =   1710
         End
         Begin VB.Frame fraSplit 
            Height          =   30
            Index           =   12
            Left            =   990
            TabIndex        =   164
            Top             =   315
            Width           =   2235
         End
         Begin VB.TextBox txtBabyLeft 
            Height          =   300
            Left            =   2940
            MaxLength       =   3
            TabIndex        =   100
            Text            =   "20"
            Top             =   1965
            Width           =   630
         End
         Begin VB.Frame fraSplit 
            Height          =   30
            Index           =   10
            Left            =   945
            TabIndex        =   99
            Top             =   3180
            Width           =   2235
         End
         Begin VB.Frame fraSplit 
            Height          =   30
            Index           =   9
            Left            =   2370
            TabIndex        =   98
            Top             =   105
            Width           =   4245
         End
         Begin VB.CommandButton cmdTabFont 
            Caption         =   "文本字体(&F)"
            Height          =   300
            Left            =   390
            TabIndex        =   97
            Top             =   3450
            Width           =   1215
         End
         Begin VB.CommandButton cmdTabTextColor 
            Caption         =   "文本颜色(&R)"
            Height          =   300
            Left            =   390
            TabIndex        =   96
            Top             =   3810
            Width           =   1215
         End
         Begin VB.Frame fraSplit 
            Height          =   30
            Index           =   8
            Left            =   930
            TabIndex        =   95
            Top             =   2445
            Width           =   2235
         End
         Begin VB.Frame fraSplit 
            Height          =   30
            Index           =   7
            Left            =   930
            TabIndex        =   94
            Top             =   1530
            Width           =   2235
         End
         Begin VB.TextBox txtBabyTitleText 
            Height          =   300
            Left            =   390
            TabIndex        =   93
            Text            =   "婴儿每日记录"
            Top             =   2655
            Width           =   2805
         End
         Begin VB.CommandButton cmdBybyTitleFont 
            Caption         =   "标题字体(&T)"
            Height          =   300
            Left            =   3270
            TabIndex        =   92
            Top             =   2655
            Width           =   1260
         End
         Begin VB.OptionButton optTabTiers 
            Caption         =   "单(&1)"
            Height          =   180
            Index           =   0
            Left            =   1185
            TabIndex        =   91
            Top             =   1710
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton optTabTiers 
            Caption         =   "双(&2)"
            Height          =   180
            Index           =   1
            Left            =   1980
            TabIndex        =   90
            Top             =   1710
            Width           =   780
         End
         Begin VB.OptionButton optTabTiers 
            Caption         =   "三(&3)"
            Height          =   180
            Index           =   2
            Left            =   2775
            TabIndex        =   89
            Top             =   1710
            Width           =   780
         End
         Begin VB.TextBox txtBabyTabRowHeight 
            Height          =   300
            Left            =   1245
            MaxLength       =   3
            TabIndex        =   88
            Text            =   "300"
            Top             =   990
            Width           =   615
         End
         Begin VB.CommandButton cmdTabGridColor 
            Caption         =   "表格颜色(&G)"
            Height          =   300
            Left            =   390
            TabIndex        =   87
            Top             =   4155
            Width           =   1215
         End
         Begin VB.TextBox txtBabyCol 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   86
            Text            =   "3"
            Top             =   570
            Width           =   375
         End
         Begin MSComCtl2.UpDown udBabyLeft 
            Height          =   300
            Left            =   3570
            TabIndex        =   101
            Top             =   1950
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtBabyLeft"
            BuddyDispid     =   196615
            OrigLeft        =   2040
            OrigTop         =   765
            OrigRight       =   2295
            OrigBottom      =   1065
            Max             =   300
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udHeadCol 
            Height          =   300
            Left            =   4140
            TabIndex        =   166
            Top             =   1005
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtHeadCol"
            BuddyDispid     =   196612
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
            Height          =   300
            Left            =   4155
            TabIndex        =   169
            Top             =   570
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   2
            BuddyControl    =   "txtHeadRow"
            BuddyDispid     =   196611
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
            AutoSize        =   -1  'True
            Caption         =   "表头单元"
            Height          =   180
            Left            =   210
            TabIndex        =   173
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblHeadRow 
            AutoSize        =   -1  'True
            Caption         =   "层号"
            Height          =   180
            Left            =   3390
            TabIndex        =   172
            Top             =   690
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "列号"
            Height          =   180
            Left            =   3390
            TabIndex        =   171
            Top             =   1095
            Width           =   360
         End
         Begin VB.Label lblHeadText 
            AutoSize        =   -1  'True
            Caption         =   "文本"
            Height          =   180
            Left            =   4395
            TabIndex        =   170
            Top             =   825
            Width           =   360
         End
         Begin VB.Label lblleft 
            AutoSize        =   -1  'True
            Caption         =   "表格相对于曲线部分的相对距离          mm"
            Height          =   180
            Left            =   420
            TabIndex        =   112
            Top             =   2070
            Width           =   3600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "表列设置"
            Height          =   180
            Left            =   210
            TabIndex        =   111
            Top             =   3090
            Width           =   720
         End
         Begin VB.Label lblBabyStyle 
            AutoSize        =   -1  'True
            Caption         =   "婴儿体温单设置(下表格)"
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
            Left            =   225
            TabIndex        =   110
            Top             =   15
            Width           =   2415
         End
         Begin VB.Label lblBabyFont 
            Caption         =   "宋体,9"
            Height          =   180
            Left            =   1650
            TabIndex        =   109
            Top             =   3555
            Width           =   1875
         End
         Begin VB.Label lblTabTextColor 
            Caption         =   "文本颜色"
            Height          =   180
            Left            =   1650
            TabIndex        =   108
            Top             =   3915
            Width           =   1635
         End
         Begin VB.Label lblBabyBasic 
            AutoSize        =   -1  'True
            Caption         =   "基本形态"
            Height          =   180
            Left            =   210
            TabIndex        =   107
            Top             =   1455
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "标题文本"
            Height          =   180
            Left            =   210
            TabIndex        =   106
            Top             =   2355
            Width           =   720
         End
         Begin VB.Label lblBabyTitleFont 
            Caption         =   "宋体,20"
            Height          =   180
            Left            =   4590
            TabIndex        =   105
            Top             =   2730
            Width           =   1695
         End
         Begin VB.Label lblTabTiers 
            AutoSize        =   -1  'True
            Caption         =   "表头层数"
            Height          =   180
            Left            =   420
            TabIndex        =   104
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label lblBabyTabRowHeight 
            AutoSize        =   -1  'True
            Caption         =   "最小行高"
            Height          =   180
            Left            =   450
            TabIndex        =   103
            Top             =   1095
            Width           =   720
         End
         Begin VB.Shape shpTabGridColor 
            Height          =   180
            Left            =   1665
            Top             =   4290
            Width           =   1605
         End
         Begin VB.Label lblHeadCol 
            AutoSize        =   -1  'True
            Caption         =   "列号"
            Height          =   180
            Left            =   420
            TabIndex        =   102
            Top             =   690
            Width           =   360
         End
      End
      Begin VB.CheckBox chkBaby 
         Caption         =   "婴儿体温单"
         Height          =   270
         Left            =   435
         TabIndex        =   143
         Top             =   3720
         Width           =   1290
      End
      Begin VB.PictureBox picSpecial 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3585
         Left            =   15
         ScaleHeight     =   3585
         ScaleWidth      =   6885
         TabIndex        =   132
         Top             =   4035
         Width           =   6885
         Begin VB.Frame fraSplit 
            Height          =   30
            Index           =   11
            Left            =   1965
            TabIndex        =   135
            Top             =   135
            Width           =   4665
         End
         Begin VB.TextBox txtDownTabRowHeight 
            Height          =   300
            Left            =   3495
            MaxLength       =   3
            TabIndex        =   134
            Text            =   "255"
            Top             =   3210
            Width           =   480
         End
         Begin VB.TextBox txtAddNullTab 
            Height          =   300
            Left            =   5745
            MaxLength       =   2
            TabIndex        =   133
            Text            =   "0"
            Top             =   3195
            Width           =   480
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgTab 
            Height          =   2475
            Left            =   2700
            TabIndex        =   136
            Top             =   630
            Width           =   3525
            _cx             =   6218
            _cy             =   4366
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
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
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
         Begin MSComctlLib.ListView lvwTabItem 
            Height          =   2850
            Left            =   345
            TabIndex        =   137
            Tag             =   "10"
            Top             =   630
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   5027
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "特殊项目栏(下表格)"
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
            Left            =   240
            TabIndex        =   142
            Top             =   45
            Width           =   1770
         End
         Begin VB.Label lblColumnItems 
            AutoSize        =   -1  'True
            Caption         =   "可选体温项目:"
            Height          =   180
            Left            =   375
            TabIndex        =   141
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label lblSelectColumnItems 
            AutoSize        =   -1  'True
            Caption         =   "已选体温项目:"
            Height          =   180
            Left            =   2715
            TabIndex        =   140
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label lblDownTabRowHeight 
            AutoSize        =   -1  'True
            Caption         =   "最小高度"
            Height          =   180
            Left            =   2700
            TabIndex        =   139
            Top             =   3255
            Width           =   720
         End
         Begin VB.Label lblAddNullTab 
            AutoSize        =   -1  'True
            Caption         =   "表格空行"
            Height          =   180
            Left            =   4980
            TabIndex        =   138
            Top             =   3255
            Width           =   720
         End
      End
      Begin VB.TextBox txtAddCurveNull 
         Height          =   300
         Left            =   1920
         MaxLength       =   2
         TabIndex        =   131
         Text            =   "0"
         Top             =   3375
         Width           =   600
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   3
         Left            =   3570
         TabIndex        =   130
         Top             =   2145
         Width           =   3120
      End
      Begin VB.TextBox txtTabRowName 
         Height          =   300
         Index           =   3
         Left            =   5040
         MaxLength       =   10
         TabIndex        =   129
         Text            =   "时       间"
         Top             =   1365
         Width           =   1290
      End
      Begin VB.TextBox txtCurveRowHeight 
         Height          =   300
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   128
         Text            =   "90"
         Top             =   3030
         Width           =   600
      End
      Begin VB.TextBox txtCurveColWidth 
         Height          =   300
         Left            =   1920
         MaxLength       =   3
         TabIndex        =   127
         Text            =   "180"
         Top             =   2685
         Width           =   600
      End
      Begin VB.TextBox txtScaleColWidth 
         Height          =   300
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   126
         Text            =   "1350"
         Top             =   2340
         Width           =   600
      End
      Begin VB.TextBox txtTabRowName 
         Height          =   300
         Index           =   2
         Left            =   3765
         MaxLength       =   10
         TabIndex        =   125
         Text            =   "手术后天数"
         Top             =   1365
         Width           =   1290
      End
      Begin VB.TextBox txtTabRowName 
         Height          =   300
         Index           =   1
         Left            =   2490
         MaxLength       =   10
         TabIndex        =   124
         Text            =   "住院天数"
         Top             =   1365
         Width           =   1290
      End
      Begin VB.TextBox txtTabRowName 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   123
         Text            =   "日       期"
         Top             =   1365
         Width           =   1290
      End
      Begin VB.TextBox txtTabRowHeight 
         Height          =   300
         Left            =   1215
         MaxLength       =   3
         TabIndex        =   122
         Text            =   "255"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox txtTabTimeSplit 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5730
         MaxLength       =   2
         TabIndex        =   121
         Text            =   "4"
         Top             =   1005
         Width           =   345
      End
      Begin VB.TextBox txtTabBeginTime 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4230
         MaxLength       =   2
         TabIndex        =   120
         Text            =   "4"
         Top             =   1005
         Width           =   345
      End
      Begin VB.TextBox txtTabDayTime 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2700
         MaxLength       =   2
         TabIndex        =   119
         Text            =   "6"
         Top             =   1005
         Width           =   345
      End
      Begin VB.TextBox txtTabDays 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1215
         MaxLength       =   2
         TabIndex        =   118
         Text            =   "7"
         Top             =   1005
         Width           =   345
      End
      Begin VB.TextBox txtTitleText 
         Height          =   300
         Left            =   420
         TabIndex        =   117
         Text            =   "精神病专科体温单"
         Top             =   420
         Width           =   2895
      End
      Begin VB.CommandButton cmdTitleFont 
         Caption         =   "标题字体(&T)"
         Height          =   300
         Left            =   3450
         TabIndex        =   116
         Top             =   420
         Width           =   1185
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   0
         Left            =   1995
         TabIndex        =   115
         Top             =   885
         Width           =   4695
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   1
         Left            =   1020
         TabIndex        =   114
         Top             =   210
         Width           =   5670
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   2
         Left            =   1770
         TabIndex        =   113
         Top             =   2145
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgCurve 
         Height          =   1695
         Left            =   2760
         TabIndex        =   144
         Top             =   2340
         Width           =   3525
         _cx             =   6218
         _cy             =   2990
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
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
      Begin MSComCtl2.UpDown udTabTimeSplit 
         Height          =   300
         Left            =   6075
         TabIndex        =   145
         Top             =   1005
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   4
         BuddyControl    =   "txtTabTimeSplit"
         BuddyDispid     =   196655
         OrigLeft        =   1530
         OrigTop         =   105
         OrigRight       =   1785
         OrigBottom      =   360
         Max             =   4
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTabDayTime 
         Height          =   300
         Left            =   3045
         TabIndex        =   146
         Top             =   990
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   6
         BuddyControl    =   "txtTabDayTime"
         BuddyDispid     =   196657
         OrigLeft        =   1530
         OrigTop         =   105
         OrigRight       =   1785
         OrigBottom      =   360
         Increment       =   2
         Max             =   24
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTabDays 
         Height          =   300
         Left            =   1560
         TabIndex        =   147
         Top             =   1005
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   7
         BuddyControl    =   "txtTabDays"
         BuddyDispid     =   196658
         OrigLeft        =   1530
         OrigTop         =   105
         OrigRight       =   1785
         OrigBottom      =   360
         Max             =   40
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTabBeginTime 
         Height          =   300
         Left            =   4590
         TabIndex        =   148
         Top             =   1005
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   4
         BuddyControl    =   "txtTabBeginTime"
         BuddyDispid     =   196656
         OrigLeft        =   1530
         OrigTop         =   105
         OrigRight       =   1785
         OrigBottom      =   360
         Max             =   4
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblAddCurveNull 
         AutoSize        =   -1  'True
         Caption         =   "曲线表格添加行数"
         Height          =   180
         Left            =   420
         TabIndex        =   163
         Top             =   3435
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "项目选择:"
         Height          =   180
         Left            =   2760
         TabIndex        =   162
         Top             =   2070
         Width           =   810
      End
      Begin VB.Label lblCurveRowHeight 
         AutoSize        =   -1  'True
         Caption         =   "曲线表格最小高度"
         Height          =   180
         Left            =   420
         TabIndex        =   161
         Top             =   3090
         Width           =   1440
      End
      Begin VB.Label lblCurveColWidth 
         AutoSize        =   -1  'True
         Caption         =   "曲线表格最小宽度"
         Height          =   180
         Left            =   420
         TabIndex        =   160
         Top             =   2745
         Width           =   1440
      End
      Begin VB.Label lblScaleColWidth 
         AutoSize        =   -1  'True
         Caption         =   "刻度列最小总宽度"
         Height          =   180
         Left            =   420
         TabIndex        =   159
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label lblTabRowName 
         AutoSize        =   -1  'True
         Caption         =   "列头名称"
         Height          =   180
         Left            =   420
         TabIndex        =   158
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label lblTitleText 
         AutoSize        =   -1  'True
         Caption         =   "标题文本"
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
         Left            =   225
         TabIndex        =   157
         Top             =   135
         Width           =   780
      End
      Begin VB.Label lblBasic 
         AutoSize        =   -1  'True
         Caption         =   "一般项目栏(上表格)"
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
         Left            =   210
         TabIndex        =   156
         Top             =   795
         Width           =   1770
      End
      Begin VB.Label lblTabTimeSplit 
         AutoSize        =   -1  'True
         Caption         =   "时间间隔"
         Height          =   180
         Left            =   4950
         TabIndex        =   155
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTabBeginTime 
         AutoSize        =   -1  'True
         Caption         =   "开始时点"
         Height          =   180
         Left            =   3450
         TabIndex        =   154
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTabDays 
         AutoSize        =   -1  'True
         Caption         =   "监测天数"
         Height          =   180
         Left            =   420
         TabIndex        =   153
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTabRowHeight 
         AutoSize        =   -1  'True
         Caption         =   "最小高度"
         Height          =   180
         Left            =   420
         TabIndex        =   152
         Top             =   1755
         Width           =   720
      End
      Begin VB.Label lblTabDayTime 
         AutoSize        =   -1  'True
         Caption         =   "监测次数"
         Height          =   180
         Left            =   1920
         TabIndex        =   151
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTitleFont 
         Caption         =   "宋体,20"
         Height          =   180
         Left            =   4650
         TabIndex        =   150
         Top             =   525
         Width           =   1605
      End
      Begin VB.Label lblRecordStyle 
         AutoSize        =   -1  'True
         Caption         =   "体征绘制栏(曲线)"
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
         Left            =   210
         TabIndex        =   149
         Top             =   2070
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList imgSize 
      Left            =   1125
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   9
      ImageHeight     =   9
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendWaveStyle.frx":1CCA
            Key             =   "-"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTendWaveStyle.frx":21B4
            Key             =   "+"
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimDraw 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   7590
      Top             =   5160
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   6930
      Index           =   1
      Left            =   8040
      ScaleHeight     =   6930
      ScaleWidth      =   6225
      TabIndex        =   52
      Top             =   3075
      Width           =   6225
      Begin VB.PictureBox picCloumn 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   3705
         Left            =   255
         ScaleHeight     =   3705
         ScaleWidth      =   6105
         TabIndex        =   59
         Top             =   3015
         Visible         =   0   'False
         Width           =   6105
         Begin VB.TextBox txtSeachColumnItems 
            Height          =   315
            Left            =   720
            TabIndex        =   81
            Top             =   705
            Width           =   1890
         End
         Begin VB.ListBox lstColumnItems 
            Height          =   2400
            Left            =   225
            TabIndex        =   74
            Top             =   1125
            Width           =   2370
         End
         Begin VB.CommandButton cmdColumn 
            Caption         =   "选用(&S)"
            Height          =   300
            Index           =   0
            Left            =   2760
            TabIndex        =   73
            Top             =   1125
            Width           =   1100
         End
         Begin VB.CommandButton cmdColumn 
            Caption         =   "删除(&E)"
            Enabled         =   0   'False
            Height          =   300
            Index           =   1
            Left            =   2760
            TabIndex        =   72
            Top             =   1410
            Width           =   1100
         End
         Begin VB.TextBox txtColumnNo 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4215
            MaxLength       =   2
            TabIndex        =   71
            Text            =   "1"
            Top             =   360
            Width           =   390
         End
         Begin VB.ListBox lstColumnUsed 
            Height          =   1680
            Left            =   3990
            TabIndex        =   69
            Top             =   690
            Width           =   2730
         End
         Begin VB.TextBox txtColumnPostfix 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4425
            TabIndex        =   68
            Top             =   2805
            Width           =   2295
         End
         Begin VB.CommandButton cmdColumn 
            Caption         =   "应用(&Y)"
            Height          =   300
            Index           =   2
            Left            =   2760
            TabIndex        =   67
            Top             =   1875
            Width           =   1100
         End
         Begin VB.TextBox txtColumnPrefix 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4425
            TabIndex        =   66
            Top             =   2430
            Width           =   2295
         End
         Begin VB.CheckBox chk 
            Caption         =   "对角线"
            Height          =   210
            Left            =   5610
            TabIndex        =   65
            Top             =   3240
            Width           =   1020
         End
         Begin VB.PictureBox picAlign 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4425
            ScaleHeight     =   315
            ScaleWidth      =   1005
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   3180
            Width           =   1005
            Begin VB.OptionButton optAlign 
               Height          =   315
               Index           =   2
               Left            =   660
               Picture         =   "frmTendWaveStyle.frx":269E
               Style           =   1  'Graphical
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   0
               Width           =   345
            End
            Begin VB.OptionButton optAlign 
               Height          =   315
               Index           =   1
               Left            =   330
               Picture         =   "frmTendWaveStyle.frx":29F7
               Style           =   1  'Graphical
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   0
               Width           =   345
            End
            Begin VB.OptionButton optAlign 
               Height          =   315
               Index           =   0
               Left            =   0
               Picture         =   "frmTendWaveStyle.frx":2D87
               Style           =   1  'Graphical
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   0
               Width           =   345
            End
         End
         Begin VB.Frame fraSplit 
            Height          =   30
            Index           =   4
            Left            =   1035
            TabIndex        =   60
            Top             =   255
            Width           =   1590
         End
         Begin MSComCtl2.UpDown udColumnNo 
            Height          =   300
            Left            =   4605
            TabIndex        =   70
            Top             =   330
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtColumnNo"
            BuddyDispid     =   196682
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "查找"
            Height          =   180
            Left            =   270
            TabIndex        =   82
            Top             =   780
            Width           =   360
         End
         Begin VB.Label lblBaby 
            AutoSize        =   -1  'True
            Caption         =   "可选护理记录项目:"
            Height          =   180
            Left            =   240
            TabIndex        =   80
            Top             =   420
            Width           =   1530
         End
         Begin VB.Label lblColumnNo 
            AutoSize        =   -1  'True
            Caption         =   "第        列内容项目:"
            Height          =   180
            Left            =   4005
            TabIndex        =   79
            Top             =   420
            Width           =   1890
         End
         Begin VB.Label lblColumnPrefix 
            AutoSize        =   -1  'True
            Caption         =   "前缀"
            Height          =   180
            Left            =   3990
            TabIndex        =   78
            Top             =   2475
            Width           =   360
         End
         Begin VB.Label lblColumnPostfix 
            AutoSize        =   -1  'True
            Caption         =   "后缀"
            Height          =   180
            Left            =   3990
            TabIndex        =   77
            Top             =   2850
            Width           =   360
         End
         Begin VB.Label lbl列对齐 
            AutoSize        =   -1  'True
            Caption         =   "对齐"
            Height          =   180
            Left            =   3990
            TabIndex        =   76
            Top             =   3240
            Width           =   360
         End
         Begin VB.Label lblColText 
            AutoSize        =   -1  'True
            Caption         =   "表列内容"
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
            Left            =   240
            TabIndex        =   75
            Top             =   180
            Width           =   780
         End
      End
      Begin MSComCtl2.FlatScrollBar hsb 
         Height          =   255
         Left            =   60
         TabIndex        =   56
         Top             =   1050
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Arrows          =   65536
         Max             =   100
         Orientation     =   1179649
      End
      Begin MSComCtl2.FlatScrollBar vsb 
         Height          =   1155
         Left            =   1935
         TabIndex        =   55
         Top             =   0
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   2037
         _Version        =   393216
         Appearance      =   0
         Max             =   100
         Orientation     =   1179648
      End
      Begin VB.PictureBox picDraw 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   0
         ScaleHeight     =   900
         ScaleWidth      =   1575
         TabIndex        =   53
         Top             =   0
         Width           =   1575
      End
      Begin VB.PictureBox Picbaby 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   1665
         Left            =   0
         ScaleHeight     =   1665
         ScaleWidth      =   7860
         TabIndex        =   57
         Top             =   1320
         Width           =   7860
         Begin VB.TextBox txtBabyTitle 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   915
            TabIndex        =   83
            Text            =   "婴儿每日记录"
            Top             =   240
            Width           =   4425
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgThis 
            Height          =   1560
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   5505
            _cx             =   9710
            _cy             =   2752
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
            FormatString    =   $"frmTendWaveStyle.frx":310D
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
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4005
      Index           =   0
      Left            =   8715
      ScaleHeight     =   4005
      ScaleWidth      =   6645
      TabIndex        =   35
      Top             =   285
      Width           =   6645
      Begin XtremeSuiteControls.TabControl tbcStyle 
         Height          =   3930
         Left            =   420
         TabIndex        =   50
         Top             =   210
         Width           =   5460
         _Version        =   589884
         _ExtentX        =   9631
         _ExtentY        =   6932
         _StockProps     =   64
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   7590
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   51
      Top             =   9990
      Width           =   16200
      _ExtentX        =   28575
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendWaveStyle.frx":3195
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24686
            Text            =   "可以根据医院实际情况，设置调整体温单格式、打印和页眉页脚。"
            TextSave        =   "可以根据医院实际情况，设置调整体温单格式、打印和页眉页脚。"
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
   Begin VB.PictureBox picOutput 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   7875
      Left            =   390
      ScaleHeight     =   7875
      ScaleWidth      =   6930
      TabIndex        =   0
      Top             =   270
      Width           =   6930
      Begin VB.PictureBox picPrint 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         Height          =   3225
         Left            =   0
         ScaleHeight     =   3225
         ScaleWidth      =   6885
         TabIndex        =   3
         Top             =   330
         Width           =   6885
         Begin VB.ComboBox cboPrinter 
            Height          =   300
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   315
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.Frame Frame5 
            Caption         =   "边距(mm)"
            Height          =   1065
            Left            =   120
            TabIndex        =   17
            Top             =   1755
            Width           =   2805
            Begin MSComCtl2.UpDown UDRight 
               Height          =   300
               Left            =   2460
               TabIndex        =   29
               Top             =   600
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   25
               BuddyControl    =   "txtRight"
               BuddyDispid     =   196706
               OrigLeft        =   2010
               OrigTop         =   600
               OrigRight       =   2250
               OrigBottom      =   900
               Max             =   100
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UDDown 
               Height          =   300
               Left            =   2460
               TabIndex        =   23
               Top             =   270
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   25
               BuddyControl    =   "txtDown"
               BuddyDispid     =   196704
               OrigLeft        =   2010
               OrigTop         =   270
               OrigRight       =   2250
               OrigBottom      =   585
               Max             =   100
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UDUp 
               Height          =   300
               Left            =   1260
               TabIndex        =   20
               Top             =   270
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   25
               BuddyControl    =   "txtUP"
               BuddyDispid     =   196703
               OrigLeft        =   915
               OrigTop         =   270
               OrigRight       =   1155
               OrigBottom      =   585
               Max             =   100
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UDLeft 
               Height          =   300
               Left            =   1260
               TabIndex        =   26
               Top             =   615
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   529
               _Version        =   393216
               Value           =   25
               BuddyControl    =   "txtLeft"
               BuddyDispid     =   196705
               OrigLeft        =   915
               OrigTop         =   615
               OrigRight       =   1155
               OrigBottom      =   930
               Max             =   100
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtUP 
               Height          =   300
               Left            =   720
               MaxLength       =   3
               TabIndex        =   19
               TabStop         =   0   'False
               Text            =   "25"
               Top             =   270
               Width           =   540
            End
            Begin VB.TextBox txtDown 
               Height          =   300
               Left            =   1905
               MaxLength       =   3
               TabIndex        =   22
               TabStop         =   0   'False
               Text            =   "25"
               Top             =   270
               Width           =   540
            End
            Begin VB.TextBox txtLeft 
               Height          =   300
               Left            =   720
               MaxLength       =   3
               TabIndex        =   25
               TabStop         =   0   'False
               Text            =   "25"
               Top             =   615
               Width           =   540
            End
            Begin VB.TextBox txtRight 
               Height          =   300
               Left            =   1905
               MaxLength       =   3
               TabIndex        =   28
               TabStop         =   0   'False
               Text            =   "25"
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "右"
               Height          =   180
               Left            =   1695
               TabIndex        =   27
               Top             =   660
               Width           =   180
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "下"
               Height          =   180
               Left            =   1695
               TabIndex        =   21
               Top             =   330
               Width           =   180
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "上"
               Height          =   180
               Left            =   510
               TabIndex        =   18
               Top             =   330
               Width           =   180
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "左"
               Height          =   180
               Left            =   510
               TabIndex        =   24
               Top             =   675
               Width           =   180
            End
         End
         Begin VB.Frame fraOrient 
            Caption         =   "纸向"
            Height          =   1065
            Left            =   2925
            TabIndex        =   30
            Top             =   1755
            Width           =   1425
            Begin VB.OptionButton optPortrait 
               Caption         =   "纵向"
               Height          =   285
               Left            =   675
               TabIndex        =   31
               Top             =   315
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optCross 
               Caption         =   "横向"
               Height          =   285
               Left            =   675
               TabIndex        =   32
               Top             =   600
               Width           =   660
            End
            Begin VB.Image imgPortrait 
               Height          =   480
               Left            =   120
               Picture         =   "frmTendWaveStyle.frx":3A29
               Top             =   330
               Width           =   480
            End
            Begin VB.Image imgCross 
               Height          =   480
               Left            =   120
               Picture         =   "frmTendWaveStyle.frx":42F3
               Top             =   330
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "纸张"
            Height          =   1065
            Left            =   120
            TabIndex        =   6
            Top             =   675
            Width           =   4230
            Begin VB.ComboBox cboPage 
               Height          =   300
               Left            =   720
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   225
               Width           =   3300
            End
            Begin MSComCtl2.UpDown UDHeight 
               Height          =   285
               Left            =   3495
               TabIndex        =   15
               Top             =   630
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtHeight"
               BuddyDispid     =   196718
               OrigLeft        =   2985
               OrigTop         =   630
               OrigRight       =   3225
               OrigBottom      =   930
               Max             =   460
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UDWidth 
               Height          =   285
               Left            =   1275
               TabIndex        =   11
               Top             =   615
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   503
               _Version        =   393216
               BuddyControl    =   "txtWidth"
               BuddyDispid     =   196719
               OrigLeft        =   1200
               OrigTop         =   645
               OrigRight       =   1440
               OrigBottom      =   945
               Max             =   460
               SyncBuddy       =   -1  'True
               BuddyProperty   =   65547
               Enabled         =   -1  'True
            End
            Begin VB.TextBox txtHeight 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   2955
               MaxLength       =   3
               TabIndex        =   14
               Top             =   630
               Width           =   540
            End
            Begin VB.TextBox txtWidth 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   720
               MaxLength       =   3
               TabIndex        =   10
               Top             =   615
               Width           =   540
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "大小"
               Height          =   180
               Left            =   285
               TabIndex        =   7
               Top             =   300
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "宽度"
               Height          =   180
               Left            =   300
               TabIndex        =   9
               Top             =   675
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "高度"
               Height          =   180
               Left            =   2550
               TabIndex        =   13
               Top             =   675
               Width           =   360
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "mm"
               Height          =   180
               Left            =   1560
               TabIndex        =   12
               Top             =   675
               Width           =   180
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "mm"
               Height          =   180
               Left            =   3795
               TabIndex        =   16
               Top             =   690
               Width           =   180
            End
         End
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   2145
            Left            =   4410
            ScaleHeight     =   494.587
            ScaleMode       =   0  'User
            ScaleWidth      =   460
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   675
            Width           =   1995
            Begin VB.PictureBox picPaper 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   1485
               Left            =   405
               ScaleHeight     =   1455
               ScaleMode       =   0  'User
               ScaleWidth      =   1140
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   270
               Width           =   1170
            End
            Begin VB.PictureBox picShadow 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00808080&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   1485
               Left            =   450
               ScaleHeight     =   1485
               ScaleWidth      =   1170
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   315
               Width           =   1170
            End
         End
         Begin VB.Label lblPaperHint 
            AutoSize        =   -1  'True
            Caption         =   "注意:  如果实际打印机和当前打印机不符，可能导致纸张设置失效！"
            Height          =   180
            Left            =   135
            TabIndex        =   37
            Top             =   2985
            Width           =   5490
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "体温单：打印机设置"
            Height          =   180
            Left            =   720
            TabIndex        =   4
            Top             =   315
            Width           =   1620
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "frmTendWaveStyle.frx":4BBD
            Top             =   75
            Width           =   480
         End
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   6
         Left            =   930
         TabIndex        =   39
         Top             =   3705
         Width           =   5895
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   5
         Left            =   915
         TabIndex        =   2
         Top             =   210
         Width           =   5910
      End
      Begin VB.PictureBox picFoot 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3405
         Left            =   0
         ScaleHeight     =   3405
         ScaleWidth      =   6885
         TabIndex        =   40
         Top             =   3870
         Width           =   6885
         Begin VB.CommandButton cmdSync 
            Caption         =   "同步(&G)"
            Height          =   350
            Left            =   5730
            TabIndex        =   48
            ToolTipText     =   "所有护理文件的页眉页脚与当前文件的页眉页脚格式一致"
            Top             =   1530
            Width           =   1100
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "插图(&I)"
            Height          =   350
            Left            =   135
            TabIndex        =   42
            Top             =   1530
            Width           =   1710
         End
         Begin VB.CheckBox chkI 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   5250
            Picture         =   "frmTendWaveStyle.frx":5487
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "斜体(Alt+I)"
            Top             =   1530
            Width           =   345
         End
         Begin VB.CheckBox chkU 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   4920
            Picture         =   "frmTendWaveStyle.frx":BCD9
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "下划线(Alt+U)"
            Top             =   1530
            Width           =   345
         End
         Begin VB.CheckBox chkB 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   4590
            Picture         =   "frmTendWaveStyle.frx":1252B
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "粗体(Alt+B)"
            Top             =   1530
            Width           =   345
         End
         Begin VB.ComboBox cboFSize 
            Height          =   300
            Left            =   3780
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   1560
            Width           =   795
         End
         Begin VB.ComboBox cboFont 
            Height          =   300
            Left            =   1875
            Style           =   2  'Dropdown List
            TabIndex        =   43
            Top             =   1560
            Width           =   1905
         End
         Begin RichTextLib.RichTextBox rtbHead 
            Height          =   1425
            Left            =   135
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   30
            Width           =   6720
            _ExtentX        =   11853
            _ExtentY        =   2514
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   2
            OLEDragMode     =   0
            OLEDropMode     =   0
            TextRTF         =   $"frmTendWaveStyle.frx":18D7D
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
            Height          =   1425
            Left            =   135
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1950
            Width           =   6720
            _ExtentX        =   11853
            _ExtentY        =   2514
            _Version        =   393217
            BorderStyle     =   0
            ScrollBars      =   2
            OLEDragMode     =   0
            OLEDropMode     =   0
            TextRTF         =   $"frmTendWaveStyle.frx":18E1A
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
         Begin VB.Label lblFont 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "字体"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   1515
            TabIndex        =   54
            Top             =   1620
            Width           =   360
         End
      End
      Begin VB.Label lblFoot 
         AutoSize        =   -1  'True
         Caption         =   "页眉页脚"
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
         Left            =   135
         TabIndex        =   38
         Top             =   3630
         Width           =   780
      End
      Begin VB.Label lblPrinter 
         AutoSize        =   -1  'True
         Caption         =   "打印设置"
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
         Left            =   120
         TabIndex        =   1
         Top             =   135
         Width           =   780
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmTendWaveStyle.frx":18EB7
      Left            =   750
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendWaveStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
Private Const PHYSICALOFFSETX = 112  '对于打印设备而言，表示从物理页的左边缘到可打印区域的左边缘的距离，采用设备单位。
Private Const PHYSICALOFFSETY = 113  '对于打印设备而言，表示从物理页的上边缘到可打印区域的上边缘的距离，采用设备单位。
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '获取中英文混合字符串长度
'######################################################################################################

Private mdblW As Double  '左边不可打印比例
Private mdblH As Double  '上边不可打印比例
Private msinVStep As Single      '滚动条的步长
Private msinHStep As Single      '滚动条的步长
'打印参数变量
Private mintPage As Integer '纸张
Private mlngWidth As Long '自定义纸张宽度,Twip
Private mlngHeight As Long '自定义纸张高度'Twip
Private mintOrient As Integer   '纸向
Private mlngLeft As Long '左边距'mm
Private mlngRight As Long '右边距'mm
Private mlngTop As Long '上边距'mm
Private mlngBottom As Long '下边距'mm
Private mblnRTBFoot As Boolean
'事件控制
Private mblnChange As Boolean  '控制打印设置
Private mblnChanged As Boolean '记录数据是否发生变化
Private mblnRedraw As Boolean '记录是否需要从新画图
Private rtbThis As Object
Public mbytMode As Byte
Public mlngFileID As Long  '病历文件列表的ID

Private Type TabItemCol
    ItemNO As String '项目序号
    ItemName As String '项目名称
    ItemUnit As String '项目单位
    ItemShow As Integer  '项目表示
    ItemFrequency As String '记录频次
End Type

Private strCurFont As String
Private objFont As StdFont
Private mbln心率应用方式 As Boolean
Private mrsItems As New ADODB.Recordset
'--修改说明：50182,刘鹏飞,2012-08-24,新增体温单设置页眉页脚功能
Private WithEvents mfrmTendWavePrint As frmTendWavePrint
Attribute mfrmTendWavePrint.VB_VarHelpID = -1


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

Public Function ShowMe(ByVal frmParent As Object, ByVal lngFileID As Long) As Boolean
    mlngFileID = lngFileID
    gblnOK = False
    DataChanged = False
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    ShowMe = gblnOK
End Function

Private Sub cboFont_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboFSize_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboPage_Click()
    Dim blnOK As Boolean
    Dim dblRight As Double
    Dim dblDown As Double
    
    '纸张
    Select Case cboPage.ItemData(cboPage.ListIndex)
    Case 256
        '强行设置自定义纸张可用,不检查
        mintPage = 256
    Case Else
        Printer.PaperSize = cboPage.ItemData(cboPage.ListIndex)
        mintPage = Printer.PaperSize
    End Select
        
    optPortrait.Enabled = True
    optCross.Enabled = True
    Err = 0
    On Error Resume Next
    optCross.Tag = Printer.Orientation
    Printer.Orientation = 1
    If Printer.Orientation <> 1 Then optPortrait.Enabled = False
    Printer.Orientation = 2
    If Printer.Orientation <> 2 Then optCross.Enabled = False
    
    If optCross.Enabled = False Then
        optPortrait.Value = True
        imgPortrait.Visible = True
        imgCross.Visible = False
    End If
    If Printer.Orientation <> mintOrient Then Printer.Orientation = mintOrient
    mintOrient = Printer.Orientation
    '最后实际设置纸张大小(纸向影响之后)
    Select Case mintPage
    Case 256
        '自定义纸张认为全部可以打印
        mdblW = 0
        mdblH = 0
        
'        If cboPage.Text = "B5, 182 x 257 毫米" Then
'            mlngWidth = 182 * conRatemmToTwip
'            mlngHeight = 257 * conRatemmToTwip
'        End If
        If Val(optCross.Tag) <> mintOrient Then
            Call SetCustonPager(Me.hWnd, mlngWidth, mlngHeight)
            mlngWidth = Printer.Width
            mlngHeight = Printer.Height
        End If
        
        txtWidth.Enabled = True
        txtHeight.Enabled = True
        UDWidth.Enabled = True
        UDHeight.Enabled = True
    Case Else
        '取该打印机支持该幅面的真实尺寸
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
        
        '不可打印区域比例
        mdblW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
        mdblH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
        
        txtWidth.Enabled = False
        txtHeight.Enabled = False
        UDWidth.Enabled = False
        UDHeight.Enabled = False
    
    End Select
        
    '显示纸张尺寸
    mblnChange = False
    txtWidth.Tag = mlngWidth
    txtWidth.Text = CLng(mlngWidth / conRatemmToTwip)
    txtHeight.Tag = mlngHeight
    txtHeight.Text = CLng(mlngHeight / conRatemmToTwip)
    mblnChange = True
    
    '显示可用边距
    '最小在可打印区域之内
    '最大不超过宽高的1/4
'    If cboPage.Text = "B5, 182 x 257 毫米" Then
'        UDLeft.Min = 0
'        UDLeft.Max = 5
'    Else
    UDLeft.Min = mlngWidth / conRatemmToTwip * mdblW
    UDLeft.Max = mlngWidth / conRatemmToTwip / 4
'    End If
    UDRight.Min = UDLeft.Min
    UDRight.Max = UDLeft.Max
    
    UDUp.Min = mlngHeight / conRatemmToTwip * mdblH
    UDUp.Max = mlngHeight / conRatemmToTwip / 4
    UDDown.Min = UDUp.Min
    UDDown.Max = UDUp.Max
    
    If mlngLeft >= UDLeft.Min And mlngLeft <= UDLeft.Max Then
        UDLeft.Value = mlngLeft
    Else
        UDLeft.Value = UDLeft.Min
    End If
    If mlngRight >= UDRight.Min And mlngRight <= UDRight.Max Then
        UDRight.Value = mlngRight
    Else
        UDRight.Value = UDRight.Min
    End If
    If mlngTop >= UDUp.Min And mlngTop <= UDUp.Max Then
        UDUp.Value = mlngTop
    Else
        UDUp.Value = UDUp.Min
    End If
    If mlngBottom >= UDDown.Min And mlngBottom <= UDDown.Max Then
        UDDown.Value = mlngBottom
    Else
        UDDown.Value = UDDown.Min
    End If
    
    mlngLeft = UDLeft.Value
    mlngRight = UDRight.Value
    mlngTop = UDUp.Value
    mlngBottom = UDDown.Value
    
    '显示纸向
    mblnChange = False
    If mintOrient = 1 Then
        optPortrait.Value = True: optPortrait_Click
    Else
        optCross.Value = True: optCross_Click
    End If
    mblnChange = True
    
    '显示预览纸张
    Call ShowPaper
    '页眉页脚设置
    Call InitPageFoot
    DataChanged = True
End Sub

Private Sub LoadPage()
    Dim i As Integer
    Dim strPrinter As String
    
    '初始打印机列表
    strPrinter = GetSetting("ZLSOFT", "公共模块\zl9PrintMode\Default", "DeviceName", Printer.DeviceName)
    With cboPrinter
        .Clear
        For i = 0 To Printers.Count - 1
            .AddItem Printers(i).DeviceName
            .ItemData(.ListCount - 1) = i '打印机索引
            
            '读取存储的打印机为当前打印机,并初始化可用页面
            If strPrinter = Printers(i).DeviceName Then .ListIndex = .NewIndex
        Next
        
        '缺省初始化为当前打印机
        If .ListIndex = -1 Then
            For i = 0 To .ListCount - 1
                '读取系统当前的打印机为当前打印机,并初始化可用页面
                If .List(i) = Printer.DeviceName Then .ListIndex = i: Exit For
            Next
        End If
        .Visible = False
        .Enabled = False
    End With
End Sub

Private Sub cboPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboPrinter_Click()
    
    Dim i As Integer, j As Integer
    Dim lngCount As Long, strTmp As String
    Dim strPaperSize As String * 300
    Dim strPrinter As String
    
    Set Printer = Printers(cboPrinter.ItemData(cboPrinter.ListIndex))
    mintPage = Printer.PaperSize
     '如果支持,则保持原有纸张
     If mintPage <> 256 Then
         On Error Resume Next
         Printer.PaperSize = mintPage
         On Error GoTo 0
         mintPage = Printer.PaperSize
         mintOrient = Printer.Orientation
     End If
     
     '特殊处理，对于体温单只支持A4及B5大小的纸张
     cboPage.Clear
     '------------------------------------------------------------------------------------------
     '纸张大小
     lngCount = DeviceCapabilities(Printer.DeviceName, Printer.Port, DC_PAPERS, strPaperSize, 0)
     For i = 1 To lngCount
         j = Asc(Mid(strPaperSize, i * 2, 1)) * 256# + Asc(Mid(strPaperSize, i * 2 - 1, 1))
         
         If mbytMode = 1 Then
             If j = 9 Or j = 13 Then
                 cboPage.AddItem GetPaperName(j)
                 cboPage.ItemData(cboPage.ListCount - 1) = j
                 If j = mintPage Then cboPage.ListIndex = cboPage.NewIndex
             End If
         Else
             If j >= 1 And j <= 41 Then '只列出标准支持的纸张
                 cboPage.AddItem GetPaperName(j)
                 cboPage.ItemData(cboPage.ListCount - 1) = j
                 If j = mintPage Then cboPage.ListIndex = cboPage.NewIndex
             End If
         End If
         
     Next
    
     '------------------------------------------------------------------------------------------
     '自定义纸张处理
     i = 256
     cboPage.AddItem GetPaperName(i)
     cboPage.ItemData(cboPage.ListCount - 1) = i
     If mintPage = 256 Then cboPage.ListIndex = cboPage.NewIndex
     If cboPage.ListIndex = -1 And cboPage.ListCount > 0 Then cboPage.ListIndex = 0
End Sub


Private Function SaveData() As Boolean
    Dim strSQL As String
    Dim arrSQL() As Variant
    Dim rsTemp As New ADODB.Recordset
    Dim rsSaveData As New ADODB.Recordset
    Dim strPaper As String
    Dim blnTrans As Boolean
    Dim i As Long
    Dim lngFixedRows As Long
    
    On Error GoTo errHand
    
    If chkBaby.Value = 1 Then
        If CheckData = False Then Exit Function
    End If
    
    If Not IsNumeric(txtWidth.Text) Then
        MsgBox "请确定打印机的纸张宽度！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    If CInt(txtWidth.Text) > UDWidth.Max Then
        MsgBox "打印机的纸张宽度不能超过" & UDWidth.Max & "毫米！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    
    If Not IsNumeric(txtHeight.Text) Then
        MsgBox "请确定打印机的纸张高度！", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    If CInt(txtHeight.Text) > UDHeight.Max Then
        MsgBox "打印机的纸张高度不能超过" & UDHeight.Max & "毫米！", vbExclamation, App.Title
        txtHeight.SetFocus: Exit Function
    End If
    
    '保存数据
    If Me.optPortrait.Value = True Then
        If Val(Me.txtUP.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "上边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtDown.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "下边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtLeft.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "左边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtRight.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "右边距太大！", vbExclamation, gstrSysName: Exit Function
    Else
        If Val(Me.txtUP.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "上边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtDown.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "下边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtLeft.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "左边距太大！", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtRight.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "右边距太大！", vbExclamation, gstrSysName: Exit Function
    End If
    
    If optTabTiers(0).Value Then
        lngFixedRows = 1
    ElseIf optTabTiers(1).Value Then
        lngFixedRows = 2
    Else
        lngFixedRows = 3
    End If
    
    If Not PageHeadTest Then Exit Function
    
    '自定义纸张始终纵向保存高度和宽度
    If mintPage = 256 Then
        Call SetCustonPager(Me.hWnd, mlngWidth, mlngHeight)
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    Else
        Printer.PaperSize = mintPage
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
    
    '如果超出页面宽度则禁止保存
    If Not OverRun Then Exit Function
    
    If Not GetRecordData(rsSaveData, True) Then Exit Function
    arrSQL = Array()
    With rsSaveData
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            strSQL = "Zl_病历文件结构_Update("
        '  Id_In         In 临时病历内容.Id%Type,
            strSQL = strSQL & !ID & ","
        '  文件id_In     In 临时病历内容.文件id%Type,
            strSQL = strSQL & !文件ID & ","
        '  父id_In       In 临时病历内容.父id%Type,
            strSQL = strSQL & IIf(IsNull(!父ID), "NULL", !父ID) & ","
        '  对象序号_In   In 临时病历内容.对象序号%Type,
            strSQL = strSQL & Val(!对象序号) & ","
        '  对象类型_In   In 临时病历内容.对象类型%Type,
            strSQL = strSQL & NVL(!对象类型, 4) & ","
        '  对象标记_In   In 临时病历内容.对象标记%Type,
            strSQL = strSQL & "NULL" & ","
        '  保留对象_In   In 临时病历内容.保留对象%Type,
            strSQL = strSQL & "NULL" & ",'"
        '  对象属性_In   In 临时病历内容.对象属性%Type,
            strSQL = strSQL & NVL(!对象属性) & "',"
        '  内容行次_In   In 临时病历内容.内容行次%Type,
            strSQL = strSQL & NVL(!内容行次, "Null") & ",'"
        '  内容文本_In   In 临时病历内容.内容文本%Type,
            strSQL = strSQL & NVL(!内容文本) & "',"
        '  是否换行_In   In 临时病历内容.是否换行%Type := 0,
            strSQL = strSQL & IIf(IsNull(!是否换行), "NULL", !是否换行) & ","
        '  预制提纲id_In In 临时病历内容.预制提纲id%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  复用提纲_In   In 临时病历内容.复用提纲%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  使用时机_In   In 临时病历内容.使用时机%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  诊治要素id_In In 临时病历内容.诊治要素id%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  替换域_In     In 临时病历内容.替换域%Type := 0,
            strSQL = strSQL & "NULL" & ",'"
        '  要素名称_In   In 临时病历内容.要素名称%Type := Null,
            strSQL = strSQL & NVL(!要素名称) & "',"
        '  要素类型_In   In 临时病历内容.要素类型%Type := Null,
            strSQL = strSQL & IIf(IsNull(!要素类型), "NULL", !要素类型) & ","
        '  要素长度_In   In 临时病历内容.要素长度%Type := Null,
            strSQL = strSQL & IIf(IsNull(!要素长度), "NULL", !要素长度) & ","
        '  要素小数_In   In 临时病历内容.要素小数%Type := Null,
            strSQL = strSQL & IIf(IsNull(!要素小数), "NULL", !要素小数) & ",'"
        '  要素单位_In   In 临时病历内容.要素单位%Type := Null,
            strSQL = strSQL & NVL(!要素单位) & "',"
        '  要素表示_In   In 临时病历内容.要素表示%Type := 0,
            strSQL = strSQL & IIf(IsNull(!要素表示), "NULL", !要素表示) & ","
        '  输入形态_In   In 临时病历内容.输入形态%Type := 0,
            strSQL = strSQL & IIf(IsNull(!输入形态), "NULL", !输入形态) & ",'"
        '  要素值域_In   In 临时病历内容.要素值域%Type := Null
            strSQL = strSQL & NVL(!要素值域) & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        .MoveNext
        Loop
    End With
    If rsSaveData.RecordCount > 0 Then
        strSQL = "Zl_病历文件结构_Commit(" & mlngFileID & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    blnTrans = False
    gcnOracle.BeginTrans
    blnTrans = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "Zl_病历文件结构_Update")
    Next
        
    strSQL = "Select l.种类,l.编号, l.名称, f.编号 As 页面号, f.名称 As 页面名,f.报表,f.页眉,f.页脚" & _
        " From 病历文件列表 l, 病历页面格式 f" & _
        " Where l.种类 = f.种类(+) And l.页面 = f.编号(+) And l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取文件打印设置", mlngFileID)
    If rsTemp.EOF Then
        MsgBox "该文件在病历文件列表中不存在。请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    picFoot.Tag = NVL(rsTemp!种类) & "-" & NVL(rsTemp!编号)
    gcnOracle.CommitTrans
    blnTrans = False
    '将页眉、页脚和基础属性分开保存
    strPaper = mintPage & ";" & mintOrient & ";" & mlngHeight & ";" & mlngWidth & ";" & CLng(Me.ScaleY(mlngLeft, vbMillimeters, vbTwips)) & ";" & _
        CLng(Me.ScaleY(mlngRight, vbMillimeters, vbTwips)) & ";" & CLng(Me.ScaleY(mlngTop, vbMillimeters, vbTwips)) & ";" & _
        CLng(Me.ScaleY(mlngBottom, vbMillimeters, vbTwips))
    '保存打印数据
    strSQL = "Zl_病历页面格式_Update(3" & ",'"
    '种类_In In 病历页面格式.种类%Type,
    '编号_In In 病历页面格式.编号%Type,
    strSQL = strSQL & NVL(rsTemp!编号) & "','"
    '名称_In In 病历页面格式.名称%Type,
    strSQL = strSQL & NVL(rsTemp!名称) & "','"
    '报表_In In 病历页面格式.报表%Type,
    strSQL = strSQL & NVL(rsTemp!报表) & "','"
    '格式_In In 病历页面格式.格式%Type,
    strSQL = strSQL & strPaper & "','"
    '页眉_In In 病历页面格式.页眉%Type,
    strSQL = strSQL & NVL(rsTemp!页眉) & "','"
    '页脚_In In 病历页面格式.页脚%Type
    strSQL = strSQL & NVL(rsTemp!页脚) & "')"
    
    gcnOracle.BeginTrans
    blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, "Zl_病历页面格式_Update")
    
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
    
    gblnOK = True
    SaveData = True
    cmdSync.Enabled = True
    Exit Function
errHand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetRecordData(rsTemp As ADODB.Recordset, Optional ByVal blnSave As Boolean = False) As Boolean
'-----------------------------------------------------------------------
'功能:将基本页面的属性组织成对应病历文件结构数据集
'-----------------------------------------------------------------------
    Dim strSQL As String
    Dim rsSource As New ADODB.Recordset
    Dim lngParentId As Long, lngId As Long
    Dim lngRow As Long, lngRowNO As Long
    Dim intFields As Integer
    Dim lngItemNO As Long '项目序号
    Dim intTitleNum As Integer
    Dim lngCol As Long
    Dim strData As String
    Dim strSubItem As String, strMidSub As String
    
    On Error GoTo errHand
    strSQL = "SELECT Id, 文件id, 父id, 对象序号, 对象类型, 对象标记, 保留对象, 对象属性, 内容行次, 内容文本, 是否换行, 预制提纲id, 复用提纲, 使用时机, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度," & vbNewLine & _
        "       要素小数, 要素单位, 要素表示, 输入形态, 要素值域" & vbNewLine & _
        " FROM 病历文件结构" & vbNewLine & _
        " WHERE 文件id = 0"
    Call zlDatabase.OpenRecordset(rsSource, strSQL, "病历文件结构")
    '开始复制记录集结构体
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Type = 200 Then       '日期型处理为字符型
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
            End If
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    '1:体温单的基本样式与属性
    lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 100)
    With rsTemp
        '父对象
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = Null
        .Fields("对象序号").Value = 1: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "体温单的基本样式与属性"
        .Fields("内容文本").Value = "格式定义"
        .Update
        '子对象
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 101)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 1: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "标题文本"
        .Fields("内容文本").Value = txtTitleText.Text: .Fields("要素名称").Value = "标题文本"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 102)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 2: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "标题字体"
        .Fields("内容文本").Value = lblTitleFont.Caption: .Fields("要素名称").Value = "标题字体"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 103)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 3: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "天数"
        .Fields("内容文本").Value = Val(txtTabDays.Text): .Fields("要素名称").Value = "天数"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 104)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 4: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "监测次数"
        .Fields("内容文本").Value = Val(txtTabDayTime.Text): .Fields("要素名称").Value = "监测次数"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 105)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 5: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "开始时点"
        .Fields("内容文本").Value = Val(txtTabBeginTime.Text): .Fields("要素名称").Value = "开始时点"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 106)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 6: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "时间间隔"
        .Fields("内容文本").Value = Val(txtTabTimeSplit.Text): .Fields("要素名称").Value = "时间间隔"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 107)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 7: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "一般项目栏表格高度"
        .Fields("内容文本").Value = Val(txtTabRowHeight.Text): .Fields("要素名称").Value = "表格高度"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 108)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 8: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "一般项目栏列头名称"
        .Fields("内容文本").Value = txtTabRowName(0).Text & "@" & txtTabRowName(1).Text & "@" & txtTabRowName(2).Text & "@" & txtTabRowName(3).Text
        .Fields("要素名称").Value = "列头名称"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 109)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 9: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "刻度区域总宽度(缇)"
        .Fields("内容文本").Value = Val(txtScaleColWidth.Text): .Fields("要素名称").Value = "刻度宽度"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 110)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 10: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "绘图区域曲线表格列宽(缇)"
        .Fields("内容文本").Value = Val(txtCurveColWidth.Text): .Fields("要素名称").Value = "曲线列宽"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 111)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 11: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "绘图区域曲线表格列高(缇)"
        .Fields("内容文本").Value = Val(txtCurveRowHeight.Text): .Fields("要素名称").Value = "曲线行高"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 112)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 12: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "曲线表格添加空行数(不针对独立曲线)"
        .Fields("内容文本").Value = Val(txtAddCurveNull.Text) * 2: .Fields("要素名称").Value = "曲线空行"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 113)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 13: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "特殊项目栏表格高度"
        .Fields("内容文本").Value = Val(txtDownTabRowHeight.Text): .Fields("要素名称").Value = "表格高度1"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 114)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 14: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "特殊项目栏表格添加的空行数"
        .Fields("内容文本").Value = Val(txtAddNullTab.Text): .Fields("要素名称").Value = "表格空行"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 115)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
        .Fields("对象序号").Value = 15: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "是否是婴儿体温单"
        .Fields("内容文本").Value = chkBaby.Value: .Fields("要素名称").Value = "婴儿体温单"
        .Update
        
           '婴儿体温单设置
        '婴儿体温单设置
        If chkBaby.Value = 1 Then
            If Me.optTabTiers(0).Value Then
                intTitleNum = 1
            ElseIf Me.optTabTiers(1).Value Then
                intTitleNum = 2
            Else
                intTitleNum = 3
            End If
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 116)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 16: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格表头层数"
            .Fields("内容文本").Value = intTitleNum: .Fields("要素名称").Value = "表头层数"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 117)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 17: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格标题文本"
            .Fields("内容文本").Value = NVL(txtBabyTitleText.Text): .Fields("要素名称").Value = "婴儿标题文本"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 118)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 18: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格标题字体"
            .Fields("内容文本").Value = lblBabyTitleFont.Caption: .Fields("要素名称").Value = "婴儿标题字体"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 119)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 19: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格文本字体"
            .Fields("内容文本").Value = lblBabyFont.Caption: .Fields("要素名称").Value = "婴儿文本字体"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 120)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 20: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格标题文本颜色"
            .Fields("内容文本").Value = lblTabTextColor.ForeColor: .Fields("要素名称").Value = "婴儿文本颜色"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 121)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 21: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格颜色"
            .Fields("内容文本").Value = Me.shpTabGridColor.BorderColor: .Fields("要素名称").Value = "婴儿表格颜色"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 122)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 22: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格栏表格高度"
            .Fields("内容文本").Value = Val(txtBabyTabRowHeight.Text): .Fields("要素名称").Value = "婴儿表格高度"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 123)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 23: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格左边距"
            .Fields("内容文本").Value = Val(txtBabyLeft.Text): .Fields("要素名称").Value = "婴儿表格左边距"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 124)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
            .Fields("对象序号").Value = 24: .Fields("对象类型").Value = 4: .Fields("对象属性").Value = "婴儿体温单表格总列数"
            .Fields("内容文本").Value = vfgThis.Cols - 1: .Fields("要素名称").Value = "总列数"
            .Update
            
        End If
    End With
    '2:体温单曲线项目定义
    lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 200)
    With rsTemp
        '父对象
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = Null
        .Fields("对象序号").Value = 2: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "体温单曲线项目定义"
        .Fields("内容文本").Value = "曲线项目定义"
        .Update
        lngRowNO = 1
        For lngRow = vfgCurve.FixedRows To vfgCurve.Rows - 1
            lngItemNO = Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("项目序号")))
            If Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("选择"))) <> 0 And lngItemNO <> 0 Then
                '子对象
                lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 200 + lngRowNO)
                .AddNew
                .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
                .Fields("对象序号").Value = lngRowNO: .Fields("对象类型").Value = 4
                .Fields("对象属性").Value = vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("项目名称"))
                .Fields("内容文本").Value = vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("项目序号"))
                .Fields("要素名称").Value = vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("项目名称"))
                .Update
                lngRowNO = lngRowNO + 1
                If lngItemNO = 2 And mbln心率应用方式 = True Then
                    lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 200 + lngRowNO)
                    .AddNew
                    .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
                    .Fields("对象序号").Value = lngRowNO: .Fields("对象类型") = 4
                    .Fields("对象属性").Value = "心率"
                    .Fields("内容文本").Value = "-1"
                    .Fields("要素名称").Value = "心率"
                    .Update
                    lngRowNO = lngRowNO + 1
                End If
            End If
        Next lngRow
    End With

    '2:体温单表格项目定义
    lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 300)
    If chkBaby.Value = 0 Then
        With rsTemp
            '父对象
            .AddNew
            .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = Null
            .Fields("对象序号").Value = 3: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "体温单表格项目定义"
            .Fields("内容文本").Value = "表格项目定义"
            .Update
            lngRowNO = 1
            For lngRow = vfgTab.FixedRows To vfgTab.Rows - 1
                lngItemNO = Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号")))
                If lngItemNO <> 0 Then
                    '子对象
                    lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 300 + lngRowNO)
                    .AddNew
                    .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
                    .Fields("对象序号").Value = lngRowNO: .Fields("对象类型") = 4
                    .Fields("对象属性").Value = vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目名称"))
                    .Fields("内容文本").Value = vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号"))
                    .Fields("要素名称").Value = vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目名称"))
                    .Fields("要素表示").Value = Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")))
                    .Update
                    lngRowNO = lngRowNO + 1
                End If
            Next lngRow
        End With
    Else
        
        With rsTemp
            '父对象
            .AddNew
            .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = Null
            .Fields("对象序号").Value = 4: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "婴儿体温单表格表头项目"
            .Fields("内容文本").Value = "婴儿体温单表头项目"
            .Update
            lngRowNO = 1
            For lngRow = 2 To 4
                If vfgThis.RowHidden(lngRow) = False Then
                    For lngCol = vfgThis.FixedCols To vfgThis.Cols - 1
                        '子对象
                        lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 300 + lngRowNO)
                        .AddNew
                        .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
                        .Fields("对象序号").Value = lngCol: .Fields("对象类型") = 4
                        .Fields("内容文本").Value = NVL(vfgThis.TextMatrix(lngRow, lngCol))                       '表头内容
'                        .Fields("要素名称").Value = NVL(vfgThis.TextMatrix(lngRow, lngCol))
'                        .Fields("要素单位").Value = NVL(vfgThis.TextMatrix(lngRow, lngCol))
                        .Fields("内容行次").Value = lngRow - 1
                                                                                                                 '记录频次
                        .Update
                        lngRowNO = lngRowNO + 1
                    
                    Next lngCol
                End If
                
            Next lngRow
            '子对象
           lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 400)
            '父对象
           .AddNew
           .Fields("ID").Value = lngParentId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = Null
           .Fields("对象序号").Value = 3: .Fields("对象类型").Value = 1: .Fields("对象属性").Value = "体温单表格项目定义"
           .Fields("内容文本").Value = "表格项目定义"
           .Update
           lngRowNO = 1
            For lngCol = vfgThis.FixedCols To vfgThis.Cols - 1
               
                lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 400 + lngRowNO)
                .AddNew
                .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
                .Fields("对象序号").Value = lngCol: .Fields("对象类型") = 4
                .Fields("对象属性").Value = vfgThis.ColWidth(lngCol) & "`" & vfgThis.Cell(flexcpAlignment, 5, lngCol, 5, lngCol)
                .Fields("内容文本").Value = ""
                strData = NVL(vfgThis.TextMatrix(5, lngCol))
                
                If InStr(3, strData, "{") Then
                    strData = Mid(strData, 2)
                    strSubItem = Substr(strData, 1, (InStr(1, strData, "}") - 1) * 2)
                    strData = Mid(strData, InStr(1, strData, "}") + 1)
                    strMidSub = (Replace(strData, Mid(strData, InStr(1, strData, "{")), ""))
                    strData = Mid(strData, InStr(1, strData, "{"))
                    .Fields("内容行次").Value = 1
                    .Fields("要素名称").Value = strSubItem
                    .Fields("要素单位").Value = strMidSub
                    .Fields("要素表示").Value = Val(Split(Split(vfgThis.Cell(flexcpData, 5, lngCol, 5), " ")(0), "`")(1))
                    .Update
                    lngRowNO = lngRowNO + 1
                    lngId = IIf(blnSave = True, zlDatabase.GetNextId("病历文件结构"), 400 + lngRowNO)
                    .AddNew
                    .Fields("ID").Value = lngId: .Fields("文件ID").Value = mlngFileID: .Fields("父ID").Value = lngParentId
                    .Fields("对象序号").Value = lngCol: .Fields("对象类型") = 4
                    .Fields("对象属性").Value = vfgThis.ColWidth(lngCol) & "`" & vfgThis.Cell(flexcpAlignment, 5, lngCol, 5, lngCol)
                    .Fields("内容文本").Value = ""
                    strData = Mid(strData, 2)
                    strSubItem = Substr(strData, 1, (InStr(1, strData, "}") - 1) * 2)
                    .Fields("内容行次").Value = 2
                    .Fields("要素名称").Value = strSubItem
                    .Fields("要素单位").Value = ""
                    If InStr(vfgThis.Cell(flexcpData, 5, lngCol, 5), " ") > 0 Then
                        .Fields("要素表示").Value = Val(Split(Split(vfgThis.Cell(flexcpData, 5, lngCol, 5), " ")(1), "`")(1))
                    Else
                        .Fields("要素表示").Value = 1
                    End If
                    .Update
                    lngRowNO = lngRowNO + 1
                    
                Else
                    strData = Mid(Replace(strData, " ", ""), 2)
                    strSubItem = Replace(strData, "}", "")
                    .Fields("内容行次").Value = 1
                    .Fields("要素名称").Value = strSubItem
                    .Fields("要素单位").Value = ""
                    .Fields("要素表示").Value = 0
                    .Update
                    lngRowNO = lngRowNO + 1
                End If
            Next lngCol
        End With
    End If
    rsTemp.Filter = ""
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    rsTemp.Sort = "ID"
    
    GetRecordData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngCount As Long
    Dim lngCol As Long, lngRow As Long
    Dim intIndex As Integer
    Dim rsPreView As New ADODB.Recordset
    Dim Preview As Variant
    
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        If SaveData Then
            DataChanged = False
            Unload Me
        End If
        
    Case conMenu_Edit_Transf_Save
        
        If SaveData Then
           DataChanged = False
        End If
        
    Case conMenu_Edit_Transf_Cancle
        gblnOK = False
        Call zlRefreshData
        DataChanged = False
    Case conMenu_File_Preview
        If GetRecordData(rsPreView) = True Then
            Set mfrmTendWavePrint = New frmTendWavePrint
            Set Preview = rsPreView
            
            Call mfrmTendWavePrint.Preview(Preview, Val(txtWidth.Text), Val(txtHeight.Text), Val(txtLeft.Text))
        End If
    Case conMenu__Curve_AddLeft
        Me.vfgThis.Cols = Me.vfgThis.Cols + 1
        Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
        Me.vfgThis.ColPosition(vfgThis.Cols - 1) = Val(vfgThis.Tag)
        With Me.vfgThis
            For lngCount = .FixedCols To .Cols - 1
                .TextMatrix(0, lngCount) = lngCount
                .ColAlignment(lngCount) = flexAlignCenterCenter
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .CellBorderRange 2, .FixedCols, 2, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 0, 1, 0, 0
            .CellBorderRange 2, .FixedCols, 5, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 1, 0, 0, 0
            .MergeCol(-1) = True
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        Me.udHeadCol.Max = vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Row)
    Case conMenu__Curve_AddRight
        Me.vfgThis.Cols = Me.vfgThis.Cols + 1
        Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
        Me.vfgThis.ColPosition(vfgThis.Cols - 1) = Val(vfgThis.Tag + 1)
        With Me.vfgThis
            For lngCount = .FixedCols To .Cols - 1
                .TextMatrix(0, lngCount) = lngCount
                .ColAlignment(lngCount) = flexAlignCenterCenter
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .CellBorderRange 2, .FixedCols, 2, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 0, 1, 0, 0
            .CellBorderRange 2, .FixedCols, 5, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 1, 0, 0, 0
            .MergeCol(-1) = True
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        Me.udHeadCol.Max = vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Row)
    Case conMenu__Curve_DeleteCol
        If vfgThis.Cols <= 4 Then Exit Sub
        With Me.vfgThis
            For lngCount = vfgThis.Col To .Cols - 2
                .TextMatrix(0, lngCount) = .TextMatrix(0, lngCount + 1)
                .TextMatrix(1, lngCount) = .TextMatrix(1, lngCount + 1)
                .TextMatrix(2, lngCount) = .TextMatrix(2, lngCount + 1)
                .TextMatrix(3, lngCount) = .TextMatrix(3, lngCount + 1)
                .TextMatrix(4, lngCount) = .TextMatrix(4, lngCount + 1)
                .TextMatrix(5, lngCount) = .TextMatrix(5, lngCount + 1)
                .Cell(flexcpData, 5, lngCount, 5, lngCount) = .Cell(flexcpData, 5, lngCount + 1, 5, lngCount + 1)
            Next
        End With
        Me.vfgThis.Cols = Me.vfgThis.Cols - 1
        With Me.vfgThis
            For lngCount = .FixedCols To .Cols - 1
                .TextMatrix(0, lngCount) = lngCount
                .ColAlignment(lngCount) = flexAlignCenterCenter
                .FixedAlignment(lngCount) = flexAlignCenterCenter
            Next
            .CellBorderRange 2, .FixedCols, 2, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 0, 1, 0, 0
            .CellBorderRange 2, .FixedCols, 5, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 1, 0, 0, 0
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        Me.udHeadCol.Max = vfgThis.Cols - 1
        Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Row)
    Case conMenu__Curve_AddUP
        
        With vfgThis
            If optTabTiers(0).Value Then
                intIndex = 0
            ElseIf optTabTiers(1).Value Then
                intIndex = 1
            Else
                intIndex = 2
            End If
            If intIndex < 2 Then optTabTiers(intIndex + 1).Value = True
            For lngRow = 4 To .Row Step -1
                For lngCol = .FixedCols To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow - 1, lngCol)
                Next
            Next
            vfgThis.Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
        End With
        DataChanged = True
    Case conMenu__Curve_AddBottom
        With vfgThis
            If optTabTiers(0).Value Then
                intIndex = 0
            ElseIf optTabTiers(1).Value Then
                intIndex = 1
            Else
                intIndex = 2
            End If
            If intIndex < 2 Then optTabTiers(intIndex + 1).Value = True
            For lngRow = 4 To .Row + 1 Step -1
                For lngCol = .FixedCols To .Cols - 1
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow - 1, lngCol)
                Next
            Next
            vfgThis.Cell(flexcpText, .Row + 1, .FixedCols, .Row + 1, .Cols - 1) = ""
        End With
        DataChanged = True
    Case conMenu__Curve_DeleteRow
        For lngRow = vfgThis.Row To 4
            For lngCol = vfgThis.FixedCols To vfgThis.Cols - 1
               vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow + 1, lngCol)
            Next
        Next
        If optTabTiers(1).Value Then
            optTabTiers(0).Value = True
        ElseIf optTabTiers(2).Value Then
            optTabTiers(1).Value = True
        End If
        vfgThis.Cell(flexcpText, 4, vfgThis.FixedCols, 4, vfgThis.Cols - 1) = ""
    Case conMenu__Curve_BuddySingle
        Call vfgThis_DblClick
    Case conMenu__Curve_BuddyDouble
        picCloumn.Visible = True
        txtSeachColumnItems.Text = ""
        picCloumn.Width = 6825
        picCloumn.Left = 0
        picCloumn.Top = picDraw.Height - picCloumn.Height - 1 * vsb.Value * msinVStep
        picCloumn.Tag = 2
        udColumnNo.Value = vfgThis.Col
        txtColumnNo.Text = vfgThis.Col
    Case conMenu_File_Exit
        If picCloumn.Visible = True Then
            picCloumn.Visible = False
        Else
            Unload Me
        End If
    Case conMenu_Help_Help
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Resize()
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小
    On Error Resume Next
    Call cbsThis.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    tbcStyle.Move 15, 15, picPane(0).Width - 30, picPane(0).Height - 30
    vsbTab.Move picTable.Height - picTable.Height - picTable.Left, picTable.Top, vsbTab.Width, picTable.Height
    tbcStyle.ZOrder 0
    rtbHead.Width = picFoot.Width - 60
    rtbFoot.Width = rtbHead.Width
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Save
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Cancle
                
        Control.Enabled = DataChanged
    Case conMenu__Curve_AddUP, conMenu__Curve_AddBottom
        Control.Enabled = Not (optTabTiers(2).Value = True) And vfgThis.Row <> 5
    Case conMenu__Curve_DeleteRow
        Control.Enabled = Not (optTabTiers(0).Value = True) And vfgThis.Row <> 5
    Case conMenu__Curve_DeleteCol
        Control.Enabled = Not (vfgThis.Cols <= 4)
        
    End Select
End Sub

Private Sub chk_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = chk.Value
    End With
    
    vfgThis.Cell(flexcpData, 5, udColumnNo.Value, 5, udColumnNo.Value) = Replace(vfgThis.Cell(flexcpData, 5, udColumnNo.Value, 5, udColumnNo.Value), IIf(chk.Value = 0, "`1", "`0"), "`" & lstColumnUsed.ItemData(lstColumnUsed.ListIndex))
    mblnChanged = True
End Sub

Private Sub chkBaby_Click()
    Dim lngIndex As Long
    For lngIndex = 1 To lvwTabItem.ListItems.Count
        If lvwTabItem.ListItems(lngIndex).Checked = True Then
            lvwTabItem.ListItems(lngIndex).Checked = False
            Call lvwTabItem_ItemCheck(lvwTabItem.ListItems(lngIndex))
        End If
    Next
    mblnRedraw = True
    Call udBabyCol_Change
    Picbaby.Visible = chkBaby.Value
    picBabyTable.Visible = chkBaby.Value
    picSpecial.Visible = (chkBaby.Value = 0)
    Call CalcScrollBarSize
End Sub

Private Sub cmdBybyTitleFont_Click()
    Dim strCurFont As String
    Dim objFont As StdFont
    
    strCurFont = Me.lblTitleFont.Caption
    Call zlFontSet("标题字体", strCurFont)
    If strCurFont = Me.lblBabyTitleFont.Caption Then Exit Sub
    Me.lblBabyTitleFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
        If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
    End With
    With Me.vfgThis
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
         Dim lngCol As Long
        txtBabyTitle.Height = vfgThis.ROWHEIGHT(1) - 20
        Set txtBabyTitle.Font = objFont
    End With
    DataChanged = True
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim strTemp As String
    Dim strTmp As String, strTmp1 As String
    Dim blnSplit As Boolean                         '多个项目时检查,如果前面的项目无后缀且后一个项目无前缀,则blnSplit=False,不允许继续
    Dim intType As Integer, intFace As Integer, intLen As Integer      '项目类型
    Dim strFaces As String                          '绑定两项目,只能是录入项目0与单选项目4;绑定两个以上项目,只能是录入项目0
    Dim strName As String                           '项目名称
    Dim intCount As Integer, arrTmp() As String
    Dim lngCount As Long
    Dim arrCol(), arrColValue()
    
    With Me.lstColumnUsed
        Select Case Index
        Case 0
            If Me.lstColumnItems.ListIndex = -1 Then Exit Sub
            .AddItem "{" & Me.lstColumnItems.List(Me.lstColumnItems.ListIndex) & "}"
            .ListIndex = .NewIndex
            Me.cmdColumn(1).Enabled = True
            Me.txtColumnPrefix.Enabled = True
            Me.txtColumnPostfix.Enabled = True
            chk.Enabled = True
        Case 1
            If .ListIndex = -1 Then Exit Sub
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                .ListIndex = -1
                Me.cmdColumn(1).Enabled = False
                Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
                Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
            End If
            
            chk.Enabled = True
        Case 2
            '当一列绑定2个项目时，项目之间必须存在前缀/或后缀符号加以区分
            '当一列绑定多个项目时，项目类型必须是录入型项目
            '单选与多选项目不能与其它项目一起绑定于某列
            '系统固定的项目，如签名人，日期，时间等，一列只能绑定一个
            strTemp = ""
            strTmp = ""
            strTmp1 = ""
            
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
                If lngCount <> udColumnNo.Value And Trim(vfgThis.Cell(flexcpData, 5, lngCount, 5, lngCount)) <> "" Then
                    arrTmp = Split(Trim(vfgThis.Cell(flexcpData, 5, lngCount, 5, lngCount)), " ")
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
                
                If strName = "收缩压" Then
                    If lngCount <> 0 Then
                        MsgBox "收缩压和舒张压绑定格式应为:收缩压/舒张压", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                End If
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
                    mrsItems.Filter = "项目名称='" & strName & "'"
                    If mrsItems.RecordCount <> 0 Then
                        If Not (intType = mrsItems!项目类型 And intFace = mrsItems!项目表示) Then
                            MsgBox "当一列绑定多个项目时，项目的类型必须一致！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        If mrsItems!项目类型 = 1 And mrsItems!项目表示 = 0 And NVL(mrsItems!项目长度, 1) > 3 Then
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
                    mrsItems.Filter = "项目名称='" & strName & "'"
                    If mrsItems.RecordCount <> 0 Then
                        intType = mrsItems!项目类型
                        intFace = mrsItems!项目表示
                        intLen = NVL(mrsItems!项目长度, 1)
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
            Next
            strTemp = Trim(strTemp)
            strTmp = Replace(Trim(strTmp), " ", "")
            mrsItems.Filter = 0
            
            With vfgThis
                .TextMatrix(5, Me.udColumnNo.Value) = strTmp
                '根据对齐方式设置其内容
                .Cell(flexcpData, 5, udColumnNo.Value, 5, udColumnNo.Value) = strTemp

            End With
            DataChanged = True
            picCloumn.Visible = False
        End Select
    End With
End Sub

Private Sub cmdOpen_Click()
    Dim picTemp As StdPicture
    
    With Me.dlgThis
        .DialogTitle = "标志图选择"
        .Filename = ""
        .Filter = "图像(*.jpg;*.bmp;*.ico;*.gif)|*.jpg;*.bmp;*.ico;*.gif"
        .CancelError = False
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
    DataChanged = True
End Sub

Private Sub cboFont_Click()
    Call GetrtbObject
    If rtbThis.SelFontName <> cboFont.List(cboFont.ListIndex) Then
        rtbThis.SelFontName = cboFont.List(cboFont.ListIndex)
        DataChanged = True
    End If
End Sub

Private Sub cboFSize_Click()
    Dim sngNum As Single
    Call GetrtbObject
    sngNum = GetFontSizeNumber(cboFSize.List(cboFSize.ListIndex))
    If rtbThis.SelFontSize <> sngNum Then
        rtbThis.SelFontSize = sngNum
        DataChanged = True
    End If
End Sub

Private Sub chkB_Click()
    Call GetrtbObject
    If chkB.Value = vbChecked Then
        rtbThis.SelBold = True
    Else
        rtbThis.SelBold = False
    End If
    DataChanged = True
End Sub

Private Sub chkI_Click()
    Call GetrtbObject
    If chkI.Value = vbChecked Then
        rtbThis.SelItalic = True
    Else
        rtbThis.SelItalic = False
    End If
    DataChanged = True
End Sub

Private Sub chkU_Click()
    Call GetrtbObject
    If chkU.Value = vbChecked Then
        rtbThis.SelUnderline = True
    Else
        rtbThis.SelUnderline = False
    End If
    DataChanged = True
End Sub

Private Sub cmdTabFont_Click()
    Dim strCurFont As String
    Dim objFont As StdFont
    
    strCurFont = Me.lblBabyFont.Caption
    Call zlFontSet("文本字体", strCurFont)
    If strCurFont = Me.lblBabyFont.Caption Then Exit Sub
    Me.lblBabyFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
         .Bold = False: .Italic = False
        If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
        If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
    End With
    Set Me.vfgThis.Font = objFont
    DataChanged = True
End Sub

Private Sub cmdTabGridColor_Click()
    Dim lngCurColor As Long
    lngCurColor = Me.shpTabGridColor.BorderColor
    Call zlColorSet("表格颜色", lngCurColor)
    If lngCurColor = Me.shpTabGridColor.BorderColor Then Exit Sub
    Me.shpTabGridColor.BorderColor = lngCurColor
    With Me.vfgThis
        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 5, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
    End With
    DataChanged = True
End Sub

Private Sub cmdTabTextColor_Click()
    Dim lngCurColor As Long
    lngCurColor = Me.lblTabTextColor.ForeColor
    Call zlColorSet("文本颜色", lngCurColor)
    If lngCurColor = Me.lblTabTextColor.ForeColor Then Exit Sub
    Me.lblTabTextColor.ForeColor = lngCurColor
    Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
    DataChanged = True
End Sub

Private Sub cmdTitleFont_Click()
    Dim strCurFont As String
    strCurFont = Me.lblTitleFont.Caption
    Call zlFontSet("标题字体", strCurFont)
    If strCurFont = Me.lblTitleFont.Caption Then Exit Sub
    Me.lblTitleFont.Caption = strCurFont
    mblnRedraw = True
    DataChanged = mblnRedraw
End Sub

Private Sub cmdSync_Click()
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
    
    '获取当前设置的页眉页脚
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

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Dim lngCount As Long
    Dim strCurFont As String
    Dim objFont As StdFont
    
    On Error GoTo errHand
    
    TimDraw.Enabled = False
    Me.picPrint.BackColor = Me.BackColor
    Me.picFoot.BackColor = Me.BackColor
    Me.picTable.BackColor = Me.BackColor
    Me.picOutput.BackColor = Me.BackColor
    Me.Picbaby.BackColor = Me.BackColor
    Me.picCloumn.BackColor = Me.BackColor
    Me.picSpecial.BackColor = Me.BackColor
    Me.picBabyTable.BackColor = Me.BackColor
'    Dim Frmsub As frmTendWaveStyleSub
    If Not ExistsPrinter Then
        MsgBox "系统中没有安装任何打印机,请先安装打印机！", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Call RestoreWinState(Me, App.ProductName)

    With Me.tbcStyle
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .InsertItem 0, "基本属性", Me.picTable.hWnd, 0
        .InsertItem 1, "打印设置", Me.picOutput.hWnd, 0
        .Item(0).Selected = True
    End With
    Call InitMenuBar  '加载菜单
    
    Dim objPane As Pane
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False '实时拖动
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis

    Set objPane = dkpMan.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "设计": objPane.Options = PaneNoCaption
    Set objPane = dkpMan.CreatePane(2, 100, 200, DockRightOf, objPane): objPane.Title = "样式": objPane.Options = PaneNoCaption
    
    With Me.vfgThis
        .MergeCellsFixed = flexMergeFree
        .Rows = 6
        .ROWHEIGHT(0) = 300
        .ROWHEIGHT(1) = 300
        .TextMatrix(1, 0) = "标题文本"
        .TextMatrix(2, 0) = "表头单元"
        .TextMatrix(3, 0) = "表头单元"
        .TextMatrix(4, 0) = "表头单元"
        .TextMatrix(5, 0) = "表列内容"
        .ROWHEIGHT(5) = 800
        
        .ColWidth(0) = 1200
        .MergeCol(0) = True
        .RowHidden(3) = True
        .RowHidden(4) = True
        
        .MergeRow(1) = True
        .MergeRow(2) = True: .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(3) = True: .Cell(flexcpAlignment, 3, 1, 3, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(4) = True: .Cell(flexcpAlignment, 4, 1, 4, .Cols - 1) = flexAlignGeneralCenter
        
        .Cell(flexcpAlignment, 2, 1, 4, .Cols - 1) = flexAlignGeneralCenter
        
        Call txtTabRowHeight_Change
        strCurFont = Me.lblBabyTitleFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
            If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
        End With
        
        .ForeColor = Me.lblTabTextColor.ForeColor
        .GridColor = Me.shpTabGridColor.BorderColor
        
        .CellBorderRange 1, .FixedCols, 1, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 1
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 1, 0, 1, 1
'        .CellBorderRange 3, .FixedCols, 5, .Cols - 1, .GridColor, 0, 1, 1, 1, 1, 1
        .CellBorderRange 2, 0, 5, 0, .GridColor, 0, 0, 1, 0, 0, 0
            
        Call txtTitleText_Change
        strCurFont = Me.lblTitleFont.Caption
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
        .RowHidden(3) = True
        .RowHidden(4) = True
        Set txtBabyTitle.Font = objFont
        
    End With
    vfgThis.Editable = flexEDKbdMouse
    
    
    If Not zlRefreshData Then Unload Me
    TimDraw.Enabled = True
    DataChanged = False
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function zlRefreshData()
    Dim i As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rsItems As New ADODB.Recordset
    Dim objItem As ListItem
    Dim strSQL As String, strPaper As String
    Dim blnHead As Boolean, blnFoot As Boolean
    Dim strTitle As String, arrTitle() As String
    Dim strHead As String
    Dim lngRow As Long, lngIndex As Long
    Dim lngCount As Long
    
    On Error Resume Next
    mintOrient = 1
    mintPage = 0
    mlngLeft = 20: mlngRight = 20: mlngTop = 20: mlngBottom = 20
    Printer.Orientation = 1
    mlngWidth = CLng(txtWidth.Text * conRatemmToTwip)
    mlngHeight = CLng(txtHeight.Text * conRatemmToTwip)
    Err = 0: On Error GoTo errHand
    '刷新数据信息
    gblnOK = False
    mbln心率应用方式 = False
    '------------------------------------------------------------------
    '初始化基础数据页面
    '------------------------------------------------------------------
    With vfgCurve
        strHead = "序号,500,4,1;选择,500,4,1;项目序号,0,4,1;项目名称,1200,1,1;项目单位,800,1,1"
        Call SetVsFlexGridChangeHead(strHead, vfgCurve, 1)
        .ColDataType(.ColIndex("选择")) = flexDTBoolean
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
        .FrozenCols = .ColIndex("选择")
        .SheetBorder = &H40C0&
    End With
                    
    With Me.lvwTabItem.ColumnHeaders
        .Clear
        .Add , "_名称", "项目名称", 2050
        .Add , "_序号", "项目序号", 0
        .Add , "_单位", "项目单位", 0
        .Add , "_表示", "项目表示", 0
        Me.lvwTabItem.ListItems.Clear
    End With
    
    With vfgThis
        .Cols = 3
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vfgTab
        strHead = "序号,500,4,1;项目序号,0,4,1;项目名称,1200,1,1;项目单位,800,1,1;项目表示,0,1,1;记录频次,800,4,1"
        Call SetVsFlexGridChangeHead(strHead, vfgTab, 1)
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
    End With
    
    gstrSQL = "Select l.编号, l.名称, l.说明 From 病历文件列表 l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    Me.Caption = "体温单样式 - " & rsTemp!名称
    strTitle = rsTemp!名称
    Me.txtTitleText.Text = strTitle
    
    gstrSQL = "Select 1 From 护理记录项目 where 项目序号=[1] And 应用方式=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, -1)
    If rsTemp.RecordCount > 0 Then mbln心率应用方式 = True
    
    gstrSQL = " SELECT DECODE(A.项目序号,-1,2,A.项目序号) 排列序号,A.项目序号,A.项目名称 as 项目名, B.记录名 项目名称, A.项目单位, B.记录法,DECODE(NVL(C.项目序号,''),'',A.项目表示,4) 项目表示" & vbNewLine & _
        " FROM 护理记录项目 A, 体温记录项目 B,护理波动项目 C" & vbNewLine & _
        " WHERE A.项目序号 = B.项目序号 And A.项目序号=C.项目序号(+) AND NOT (NVL(A.应用方式,0)=2 And A.项目序号=-1) And A.项目性质=1 " & vbNewLine & _
        " ORDER BY 项目序号"
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    rsItems.Filter = "记录法=1 Or 记录法=3"
    rsItems.Sort = "排列序号,项目序号"
    With rsItems
        Do While Not .EOF
            If .AbsolutePosition > vfgCurve.Rows - 1 Then vfgCurve.Rows = .AbsolutePosition + 1
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("序号")) = .AbsolutePosition
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("项目序号")) = Val(!项目序号)
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("项目名称")) = NVL(!项目名称)
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("项目单位")) = NVL(!项目单位)
        .MoveNext
        Loop
    End With
    
    rsItems.Filter = "记录法=2"
    rsItems.Sort = "项目序号"
    With rsItems
        Do While Not .EOF
            If !项目序号 = 4 Then
                Set objItem = Me.lvwTabItem.ListItems.Add(, "_" & !项目序号, "血压")
                objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_序号").Index - 1) = "4,5"
            Else
                Set objItem = Me.lvwTabItem.ListItems.Add(, "_" & !项目序号, NVL(!项目名称))
                objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_序号").Index - 1) = !项目序号
            End If
            objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_单位").Index - 1) = NVL(!项目单位)
            objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_表示").Index - 1) = NVL(!项目表示)
            
        .MoveNext
        Loop
    End With
    


    gstrSQL = "Select b.项目序号, b.项目名称, b.项目类型, b.项目表示, b.项目长度" & vbNewLine & _
              "      From 体温记录项目 A, 护理记录项目 B " & vbNewLine & _
              "      Where a.项目序号 = b.项目序号 And 记录法 = 2 " & vbNewLine & _
              "      Order By 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    rsItems.Filter = "记录法=2"
    rsItems.Sort = "项目序号"
    With rsItems
        Me.lstColumnItems.Clear
        Me.lstColumnItems.AddItem "日期"
        Me.lstColumnItems.AddItem "时间"
        Me.lstColumnItems.AddItem "出生日期"
        Do While Not .EOF
            Me.lstColumnItems.AddItem "" & !项目名
            .MoveNext
        Loop
        Me.lstColumnItems.AddItem "护士"
        Me.lstColumnItems.ListIndex = 0
        .MoveFirst
    End With
    
    gstrSQL = "Select 项目序号,项目名称,项目类型,项目表示,项目长度 From 护理记录项目 Order By 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    '---------------------------------------------------
    '定义样式获取
    '---------------------------------------------------
    
    Me.optTabTiers(0).Value = True
    Call optTabTiers_Click(0)
    txtTabBeginTime.Tag = 4: txtTabTimeSplit.Tag = 4
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '格式定义'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "标题文本"
                Me.txtTitleText.Text = "" & !内容文本
            Case "标题字体"
                Me.lblTitleFont.Caption = "" & !内容文本
            Case "表格高度"
                Me.txtTabRowHeight.Text = Val("" & !内容文本)
                If Me.txtTabRowHeight.Text < 225 Or Me.txtTabRowHeight.Text > 600 Then
                    Me.txtTabRowHeight.Text = 225
                End If
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
            Case "婴儿标题文本"
                Me.txtBabyTitleText.Text = "" & !内容文本
            Case "婴儿文本字体"
                Me.lblBabyTitleFont.Caption = "" & !内容文本
            Case "婴儿表格高度"
                Me.txtBabyTabRowHeight = Val("" & !内容文本)
                If Me.txtBabyTabRowHeight.Text < 225 Or Me.txtBabyTabRowHeight.Text > 600 Then
                    Me.txtBabyTabRowHeight.Text = 225
                End If
            Case "文本颜色"
                Me.lblTabTextColor.ForeColor = Val("" & !内容文本)
                Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
            Case "表格颜色"
                Me.shpTabGridColor.BorderColor = Val("" & !内容文本)
                With Me.vfgThis
                    .GridColor = Me.shpTabGridColor.BorderColor
                    .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
                    .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
                End With
            Case "天数"
                Me.txtTabDays.Text = Val("" & !内容文本)
                If Me.txtTabDays.Text = "" Then Me.txtTabDays.Text = 7
            Case "总列数":  Me.udBabyCol.Value = Val("" & !内容文本)
            Case "监测次数"
                If InStr(1, ",2,4,6,8,12,24,", "," & Val("" & !内容文本) & ",") = 0 Then
                    Me.txtTabDayTime.Text = 6
                Else
                    Me.txtTabDayTime.Text = Val("" & !内容文本)
                End If
            Case "开始时点"
                Me.txtTabBeginTime.Text = Val("" & !内容文本)
                txtTabBeginTime.Tag = txtTabBeginTime.Text
            Case "时间间隔"
                Me.txtTabTimeSplit.Text = Val("" & !内容文本)
                txtTabTimeSplit.Tag = txtTabTimeSplit.Text
            Case "列头名称"
                strTitle = "" & !内容文本
                If strTitle <> "" Then
                    arrTitle = Split(strTitle, "@")
                    txtTabRowName(0).Text = arrTitle(0)
                    If UBound(arrTitle) > 0 Then
                        txtTabRowName(1).Text = arrTitle(1)
                    End If
                    If UBound(arrTitle) > 1 Then
                        txtTabRowName(2).Text = arrTitle(2)
                    End If
                    If UBound(arrTitle) > 2 Then
                        txtTabRowName(3).Text = arrTitle(3)
                    End If
                End If
                If txtTabRowName(0).Text = "" Then txtTabRowName(0).Text = "日    期"
                If txtTabRowName(1).Text = "" Then txtTabRowName(1).Text = "住院天数"
                If txtTabRowName(2).Text = "" Then txtTabRowName(2).Text = "手术后天数"
                If txtTabRowName(3).Text = "" Then txtTabRowName(3).Text = "时    间"
            Case "刻度宽度"
                txtScaleColWidth.Text = Val("" & !内容文本)
            Case "曲线列宽"
                txtCurveColWidth.Text = Val("" & !内容文本)
            Case "曲线行高"
                txtCurveRowHeight.Text = Val("" & !内容文本)
            Case "曲线空行"
                txtAddCurveNull.Text = Val("" & !内容文本) \ 2
                If Val(txtAddCurveNull.Text) < 0 Then txtAddCurveNull.Text = 0
            Case "表格高度1"
                Me.txtDownTabRowHeight.Text = Val("" & !内容文本)
                If Val(Me.txtDownTabRowHeight.Text) < 225 Or Val(Me.txtDownTabRowHeight.Text) > 600 Then
                    Me.txtDownTabRowHeight.Text = 225
                End If
            Case "表格空行"
                txtAddNullTab.Text = Val("" & !内容文本)
                If Val(txtAddNullTab.Text) < 0 Then txtAddNullTab.Text = 0
                Case "婴儿体温单"
                chkBaby.Value = Val("" & !内容文本)
            Case "表头层数"
                If Val("" & !内容文本) - 1 >= 0 Then
                    Me.optTabTiers(Val("" & !内容文本) - 1).Value = 1
                End If
            Case "婴儿标题文本"
                txtBabyTitleText.Text = "" & !内容文本
            Case "婴儿标题字体"
                lblBabyTitleFont.Caption = "" & !内容文本
            Case "婴儿文本颜色"
                lblTabTextColor.ForeColor = Val("" & !内容文本)
            Case "婴儿文本字体"
                lblBabyFont.Caption = "" & !内容文本
            Case "婴儿表格颜色"
                txtBabyTabRowHeight.Text = Val("" & !内容文本)
            Case "婴儿表格高度"
                txtBabyTabRowHeight.Text = Val("" & !内容文本)
            Case "婴儿表格左边距"
                txtBabyLeft.Text = Val("" & !内容文本)
            End Select
            .MoveNext
        Loop
    End With
    If chkBaby.Value = 0 Then Picbaby.Visible = False
    
    If Val(txtTabDays.Text) >= udTabDays.Min And Val(txtTabDays.Text) <= udTabDays.Max Then
        udTabDays.Value = Val(txtTabDays.Text)
    Else
        udTabDays.Value = udTabDays.Min
        txtTabDays.Text = udTabDays.Value
    End If
    If Val(txtTabDayTime.Text) >= udTabDayTime.Min And Val(txtTabDayTime.Text) <= udTabDayTime.Max Then
        udTabDayTime.Value = Val(txtTabDayTime.Text)
    Else
        udTabDayTime.Value = udTabDayTime.Min
        txtTabDayTime.Text = udTabDayTime.Value
    End If
    Call txtTabDayTime_Change
    
    If Val(txtTabBeginTime.Tag) >= udTabBeginTime.Min And Val(txtTabBeginTime.Tag) <= udTabBeginTime.Max Then
        udTabBeginTime.Value = Val(txtTabBeginTime.Tag)
        txtTabBeginTime.Text = udTabBeginTime.Value
    End If
    If Val(txtTabTimeSplit.Tag) >= udTabTimeSplit.Min And Val(txtTabTimeSplit.Tag) <= udTabTimeSplit.Max Then
        udTabTimeSplit.Value = Val(txtTabTimeSplit.Tag)
        txtTabTimeSplit.Text = udTabTimeSplit.Value
    End If
    '--曲线项目定义
    gstrSQL = "Select d.对象序号, d.内容文本, d.要素名称,d.要素表示 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '曲线项目定义'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            For lngRow = vfgCurve.FixedRows To vfgCurve.Rows - 1
                If Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("项目序号"))) = Val(NVL(!内容文本)) Or _
                    Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("项目序号"))) = 1 Then
                    vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("选择")) = 1
                End If
            Next
        .MoveNext
        Loop
    End With
    
    
    '--表格项目定义
    If chkBaby.Value = 0 Then
        gstrSQL = "Select d.对象序号, d.内容文本,d.要素名称,d.要素表示 " & _
            " From 病历文件结构 d, 病历文件结构 p" & _
            " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格项目定义'" & _
            " Order By d.对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            lngRow = 1
            Do While Not .EOF
                For lngIndex = 1 To lvwTabItem.ListItems.Count
                    If lvwTabItem.ListItems(lngIndex).SubItems(1) = NVL(!内容文本) Then
                        lvwTabItem.ListItems(lngIndex).Checked = True
                        If lngRow > vfgTab.Rows - 1 Then vfgTab.Rows = vfgTab.Rows + 1
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("序号")) = lngRow
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号")) = NVL(!内容文本)
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目名称")) = lvwTabItem.ListItems(lngIndex).Text
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目单位")) = lvwTabItem.ListItems(lngIndex).SubItems(2)
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目表示")) = lvwTabItem.ListItems(lngIndex).SubItems(3)
                        If InStr(1, GetTabFrequency(Val(txtTabDayTime.Text), Val(lvwTabItem.ListItems(lngIndex).SubItems(3))), Val(NVL(!要素表示))) = 0 Then
                            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
                        Else
                            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = NVL(!要素表示)
                        End If
                        If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号"))) = 3 Then
                            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = Val(txtTabDayTime.Text)
                        End If
                        lngRow = lngRow + 1
                        Exit For
                    End If
                Next
            .MoveNext
            Loop
        End With
    Else
        gstrSQL = "Select d.对象序号, d.内容行次, d.内容文本" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '婴儿体温单表头项目'" & _
        " Order By d.对象序号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            Do While Not .EOF
                If Val(!对象序号) > vfgThis.Cols - 1 Then vfgThis.Cols = vfgThis.Cols + 1
                Me.vfgThis.TextMatrix(!内容行次 + 1, !对象序号) = "" & !内容文本
                .MoveNext
            Loop
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select d.对象序号,d.对象标记, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示,d.要素值域 " & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格项目定义'" & _
        " Order By d.对象序号, d.内容行次"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        
        vfgThis.Cell(flexcpData, 5, vfgThis.FixedCols, 5, vfgThis.Cols - 1) = ""
        With rsTemp
            Me.lstColumnUsed.Clear
            Do While Not .EOF
                Me.vfgThis.ColWidth(!对象序号) = Val(Split("" & !对象属性, "`")(0))
                If InStr(1, "" & !对象属性, "`") <> 0 Then
                    vfgThis.Cell(flexcpAlignment, 5, !对象序号, 5, !对象序号) = Val(Split("" & !对象属性, "`")(1))
                Else
                    vfgThis.Cell(flexcpAlignment, 5, !对象序号, 5, !对象序号) = flexAlignLeftCenter
                End If
                If Me.udColumnNo.Value <> !对象序号 Then Me.udColumnNo.Value = !对象序号
                Me.lstColumnUsed.AddItem !内容文本 & "{" & !要素名称 & "}" & !要素单位
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = zlCommFun.NVL(!要素表示, 0)
                Call cmdColumn_Click(2)
                .MoveNext
                Loop
        End With
        vfgThis.Cell(flexcpAlignment, vfgThis.FixedRows, vfgThis.FixedCols, 4, vfgThis.Cols - 1) = flexAlignCenterCenter
    End If
        
    mblnChange = True
    Call LoadPage
    Call PrepareFont
    mblnChange = False
    
    '从病历页面格式中提取打印设置数据
    strSQL = "Select l.种类,l.编号,f.编号 As 页面号, f.格式" & _
        " From 病历文件列表 l, 病历页面格式 f" & _
        " Where l.种类 = f.种类(+) And l.页面 = f.编号(+) And l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取文件打印设置", mlngFileID)
    If Not rsTemp.EOF Then
        picFoot.Tag = NVL(rsTemp!种类) & "-" & NVL(rsTemp!页面号, rsTemp!编号)
        strPaper = "" & rsTemp!格式
        blnHead = ReadPageHead(rtbHead, picFoot.Tag)
        blnFoot = ReadPageFoot(rtbFoot, picFoot.Tag)
        cmdSync.Enabled = blnHead Or blnFoot
    End If
    
    If UBound(Split(strPaper, ";")) >= 4 Then mlngLeft = Round(Me.ScaleY(Val(Split(strPaper, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 5 Then mlngRight = Round(Me.ScaleY(Val(Split(strPaper, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 6 Then mlngTop = Round(Me.ScaleX(Val(Split(strPaper, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 7 Then mlngBottom = Round(Me.ScaleX(Val(Split(strPaper, ";")(7)), vbTwips, vbMillimeters), 2)

    If UBound(Split(strPaper, ";")) >= 0 Then
        For i = 0 To Me.cboPage.ListCount - 1
            If Me.cboPage.ItemData(i) = Val(Split(strPaper, ";")(0)) Then Me.cboPage.ListIndex = i: Exit For
        Next
        mblnChange = False
        mintPage = cboPage.ItemData(i)
        If Me.cboPage.ListIndex = i Then
            If UBound(Split(strPaper, ";")) >= 2 Then mlngHeight = Val(Split(strPaper, ";")(2))
            If UBound(Split(strPaper, ";")) >= 3 Then mlngWidth = Val(Split(strPaper, ";")(3))
            Me.txtHeight.Text = CLng(mlngHeight / conRatemmToTwip)
            Me.txtWidth.Text = CLng(mlngWidth / conRatemmToTwip)
        End If
    End If
    
    If UBound(Split(strPaper, ";")) >= 1 Then
        mintOrient = Val(Split(strPaper, ";")(1))
        If Val(Split(strPaper, ";")(1)) = 2 Then
            Me.optCross.Value = True
        Else
            Me.optPortrait.Value = True
        End If
    End If
        
    txtLeft.Text = mlngLeft
    txtRight.Text = mlngRight
    txtUP.Text = mlngTop
    txtDown.Text = mlngBottom
    On Error Resume Next
    If mintOrient = Printer.Orientation And mintPage = 256 Then
        If mintOrient = 1 Then
            Printer.Orientation = 2
        Else
            Printer.Orientation = 1
        End If
    End If
    Err.Clear: On Error GoTo errHand
    Call cboPage_Click: mblnChange = True
    DataChanged = False: mblnRedraw = True
    
    For lngCount = vfgThis.FixedCols To vfgThis.Cols - 1
        vfgThis.TextMatrix(0, lngCount) = lngCount
        vfgThis.ColAlignment(lngCount) = flexAlignCenterCenter
        vfgThis.FixedAlignment(lngCount) = flexAlignCenterCenter
    Next
    Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Row)
    vfgThis.MergeCol(-1) = True
    
    zlRefreshData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function PrepareFont()
    Dim sFont As String, i As Integer
    
    For i = 0 To Screen.FontCount - 1
       sFont = Screen.Fonts(i)
       cboFont.AddItem sFont
       If sFont = "宋体" Then cboFont.ListIndex = i
    Next i
    With cboFSize
        .AddItem "初号"
        .AddItem "小初"
        .AddItem "一号"
        .AddItem "小一"
        .AddItem "二号"
        .AddItem "小二"
        .AddItem "三号"
        .AddItem "小三"
        .AddItem "四号"
        .AddItem "小四"
        .AddItem "五号"
        .AddItem "小五"
        .AddItem "六号"
        .AddItem "小六"
        .AddItem "七号"
        .AddItem "八号"
        .AddItem 5
        .AddItem 5.5
        .AddItem 6.5
        .AddItem 7.5
        .AddItem 8
        .AddItem 9
        .AddItem 10
        .AddItem 10.5
        .AddItem 11
        .AddItem 12
        .AddItem 14
        .AddItem 16
        .AddItem 18
        .AddItem 20
        .AddItem 22
        .AddItem 24
        .AddItem 26
        .AddItem 28
        .AddItem 36
        .AddItem 48
        .AddItem 72
        .ListIndex = 10
    End With
End Function

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
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
    
    
     '快键绑定
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, "保存并退出"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "保存"): cbrControl.ToolTipText = "保存已更改的数据(Ctrl+S,F2)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "恢复"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "恢复到上次保存时的数据状态"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "进行当前体温单预览"
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "帮助(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出"): cbrControl.ToolTipText = "退出当前的设计窗体(Esc)"

    End With
        
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
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

Private Sub Form_Resize()
    On Error Resume Next
    Call SetPaneRange(dkpMan, 1, 200, Me.ScaleHeight, 500, Me.ScaleHeight)
    dkpMan.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If DataChanged Then
        Cancel = (MsgBox("更改后的设计必须保存后才生效，是否放弃保存？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    
    DataChanged = False
    
    Call SaveWinState(Me, App.ProductName)
    
    Set rtbThis = Nothing
End Sub


Private Sub lstColumnItems_DblClick()
    If picCloumn.Tag = 1 Then
        Call cmdColumn_Click(1)
        Call lstColumnItems_KeyDown(vbKeyReturn, 0)
'        picCloumn.Visible = False
    Else
        Call cmdColumn_Click(0)
    End If
End Sub

Private Sub lstColumnItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call cmdColumn_Click(0)
            Call cmdColumn_Click(2)
        Case vbKeyDelete
            Call cmdColumn_Click(1)
            Call cmdColumn_Click(2)
    End Select
End Sub

Private Sub lstColumnUsed_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        Me.txtColumnPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
        Me.txtColumnPostfix.Text = Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "}") + 1)
        chk.Value = .ItemData(.ListIndex)
    End With
End Sub

Private Sub lvwTabItem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lngIndex As Long, lngRow As Long, i As Integer, blnTrue As Boolean
    Dim arrTab() As TabItemCol
    '将之前的表格项目信息保存起来
    With vfgTab
        For lngRow = .FixedRows To .Rows - 1
            ReDim Preserve arrTab(0 To i)
            arrTab(UBound(arrTab)).ItemNO = .TextMatrix(lngRow, vfgTab.ColIndex("项目序号"))
            arrTab(UBound(arrTab)).ItemName = .TextMatrix(lngRow, vfgTab.ColIndex("项目名称"))
            arrTab(UBound(arrTab)).ItemUnit = .TextMatrix(lngRow, vfgTab.ColIndex("项目单位"))
            arrTab(UBound(arrTab)).ItemShow = Val(.TextMatrix(lngRow, vfgTab.ColIndex("项目表示")))
            arrTab(UBound(arrTab)).ItemFrequency = .TextMatrix(lngRow, vfgTab.ColIndex("记录频次"))
            i = i + 1
        Next
        .Rows = 2
        .Cell(flexcpText, 1, 0, 1, .Cols - 1) = ""
        .RowData(1) = ""
    End With
    lngRow = 1
    For lngIndex = 1 To lvwTabItem.ListItems.Count
        If lvwTabItem.ListItems(lngIndex).Checked = True Then
            blnTrue = False
            For i = 0 To UBound(arrTab)
                If arrTab(i).ItemNO = lvwTabItem.ListItems(lngIndex).SubItems(1) Then
                    blnTrue = True
                    Exit For
                End If
            Next
            If lngRow > vfgTab.Rows - 1 Then vfgTab.Rows = vfgTab.Rows + 1
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("序号")) = lngRow
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号")) = lvwTabItem.ListItems(lngIndex).SubItems(1)
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目名称")) = lvwTabItem.ListItems(lngIndex).Text
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目单位")) = lvwTabItem.ListItems(lngIndex).SubItems(2)
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目表示")) = lvwTabItem.ListItems(lngIndex).SubItems(3)
            If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号"))) = 3 Then
                vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = Val(txtTabDayTime.Text)
            Else
                If blnTrue = True Then
                    If InStr(1, GetTabFrequency(Val(txtTabDayTime.Text), Val(lvwTabItem.ListItems(lngIndex).SubItems(3))), Val(arrTab(i).ItemFrequency)) = 0 Then
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
                    Else
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = arrTab(i).ItemFrequency
                    End If
                Else
                    vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
                End If
            End If
            lngRow = lngRow + 1
        End If
    Next
    DataChanged = True
    mblnRedraw = True
End Sub

Private Sub lvwTabItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optAlign_Click(Index As Integer)
    Dim intAlign As Integer
    
    Select Case Index
    Case 0
        intAlign = flexAlignLeftCenter
    Case 1
        intAlign = flexAlignCenterCenter
    Case 2
        intAlign = flexAlignRightCenter
    End Select
    vfgThis.Cell(flexcpAlignment, 5, vfgThis.Col, 5, vfgThis.Col) = intAlign
    DataChanged = True
    
    On Error Resume Next
    chk.SetFocus
End Sub

Private Sub optCross_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If optCross.Value Then
        imgPortrait.Visible = False
        imgCross.Visible = True
        
        If mintOrient = 1 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
            
            mlngLeft = lngB
            mlngRight = lngT
            mlngTop = lngL
            mlngBottom = lngR
            If mintPage = 256 Then
                Call SetCustonPager(Me.hWnd, mlngWidth, mlngHeight)
                mlngWidth = Printer.Width
                mlngHeight = Printer.Height
            End If
        End If
        
        mintOrient = 2
        
        If mblnChange Then Call cboPage_Click
        
        DataChanged = True
    End If
End Sub

Private Sub optPortrait_Click()
    Dim lngL As Long, lngR As Long
    Dim lngT As Long, lngB As Long
    
    If optPortrait.Value Then
        imgPortrait.Visible = True
        imgCross.Visible = False
        
        If mintOrient = 2 Then
            lngL = mlngLeft
            lngR = mlngRight
            lngT = mlngTop
            lngB = mlngBottom
              
            mlngLeft = lngT
            mlngRight = lngB
            mlngTop = lngR
            mlngBottom = lngL
            
            If mintPage = 256 Then
                Call SetCustonPager(Me.hWnd, mlngWidth, mlngHeight)
                mlngWidth = Printer.Width
                mlngHeight = Printer.Height
            End If
        End If
        
        mintOrient = 1
        
        If mblnChange Then Call cboPage_Click
        
        DataChanged = True
    End If
End Sub

Private Sub optTabTiers_Click(Index As Integer)
    With vfgThis
        If optTabTiers(0).Value Then
            If .Row = 4 Or .Row = 3 Then .Row = 2
            .RowHidden(2) = False
            .RowHidden(3) = True
            .RowHidden(4) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 1
        ElseIf optTabTiers(1).Value Then
            If .Row = 4 Then .Row = 3
            .RowHidden(2) = False
            .RowHidden(3) = False
            .RowHidden(4) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 2
        Else
            .RowHidden(2) = False
            .RowHidden(3) = False
            .RowHidden(4) = False
            udHeadRow.Min = 1
            udHeadRow.Max = 3
        End If
    End With
    DataChanged = True
End Sub

Private Sub optTabTiers_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Picbaby_Resize()
    Dim lngCol As Long
    Dim longWidth As Long
    With vfgThis
        .Left = 0
        .Top = 0
        .Height = Picbaby.Height
        .Width = Picbaby.Width
    End With
    longWidth = 0
    For lngCol = 1 To vfgThis.Cols - 1
        longWidth = vfgThis.ColWidth(lngCol) + longWidth
    Next
    With txtBabyTitle
        .Left = vfgThis.ColWidth(0)
        .Top = vfgThis.ROWHEIGHT(0) + vfgThis.Top
        .Height = vfgThis.ROWHEIGHT(1) - 20
        .Width = longWidth
    End With
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        With vsb
            .Left = picPane(Index).Width - .Width
            .Top = 0
            .Height = picPane(Index).Height
        End With
        
        With hsb
            .Left = 0
            .Top = picPane(Index).Height - .Height
            .Width = picPane(Index).Width - vsb.Width
        End With
        
        With Picbaby
            .Left = 0
            .Top = picDraw.Top + picDraw.Height
            .Height = 4500
            .Width = picPane(Index).Width
        End With
        
        Call CalcScrollBarSize
    End If
End Sub

Private Sub picResize_Click(Index As Integer)
    If chkBaby.Value = 1 Then
        If Index = 2 Then Exit Sub
    Else
        If Index = 3 Then Exit Sub
    End If
End Sub

Private Sub rtbFoot_Change()
    DataChanged = True
End Sub

Private Sub rtbFoot_GotFocus()
    mblnRTBFoot = True
End Sub

Private Sub rtbHead_Change()
    DataChanged = True
End Sub

Private Sub rtbHead_GotFocus()
    mblnRTBFoot = False
End Sub

Private Sub rtbHead_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub tbcStyle_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Caption
    Case "基本属性"
        If txtTitleText.Enabled And txtTitleText.Visible Then txtTitleText.SetFocus
    Case "打印设置"
        If cboPage.Enabled And cboPage.Visible Then cboPage.SetFocus
    End Select
End Sub

Private Sub InitPageFoot()
    '页眉页脚设置
    Dim intPage As Integer
    On Error Resume Next
    intPage = cboPage.ItemData(cboPage.ListIndex)
    Printer.PaperSize = intPage
    Printer.Orientation = IIf(optPortrait.Value, 1, 2)
    If intPage = 256 Then
        If Printer.Orientation = 1 Then
            mlngWidth = CLng(Val(txtWidth.Text) * conRatemmToTwip)
            mlngHeight = CLng(Val(txtHeight.Text) * conRatemmToTwip)
        Else
            mlngHeight = CLng(Val(txtWidth.Text) * conRatemmToTwip)
            mlngWidth = CLng(Val(txtHeight.Text) * conRatemmToTwip)
        End If
        Call SetCustonPager(Me.hWnd, mlngWidth, mlngHeight)
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    Else
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
    Call SendMessage(rtbHead.hWnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
    SendMessageLong rtbHead.hWnd, EM_HIDESELECTION, 0, 0
    Call SendMessage(rtbFoot.hWnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
    SendMessageLong rtbFoot.hWnd, EM_HIDESELECTION, 0, 0

    rtbHead.Width = picFoot.Width - 140
    rtbFoot.Width = rtbHead.Width
End Sub

Private Sub TimDraw_Timer()
    Dim rsStyle As ADODB.Recordset
    If mblnRedraw = False Then Exit Sub
    If Not GetRecordData(rsStyle) Then mblnRedraw = False: Exit Sub
    picDraw.AutoRedraw = True
    Call DrawWaveStyle(picDraw, rsStyle)
    
    mblnRedraw = False
    Call CalcScrollBarSize
End Sub

Private Sub txtAddCurveNull_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtAddCurveNull_GotFocus()
    Me.txtAddCurveNull.SelStart = 0: Me.txtAddCurveNull.SelLength = 10
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtAddCurveNull_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtAddCurveNull_Validate(Cancel As Boolean)
    If Val(txtAddCurveNull.Text) < 0 Then txtAddCurveNull = 0
    If Val(txtAddCurveNull.Text) > 20 Then txtAddCurveNull = 20
End Sub

Private Sub txtAddNullTab_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtAddNullTab_GotFocus()
    Me.txtAddNullTab.SelStart = 0: Me.txtAddNullTab.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtAddNullTab_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        tbcStyle.Item(1).Selected = True: Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtAddNullTab_Validate(Cancel As Boolean)
    If Val(txtAddNullTab.Text) < 0 Then
        txtAddNullTab.Text = 0
    ElseIf Val(txtAddNullTab.Text) > 40 Then
        txtAddNullTab.Text = 40
    End If
End Sub

Private Sub txtBabyCol_Change()
    DataChanged = True
End Sub

Private Sub txtBabyLeft_Change()
    mblnChanged = True
End Sub

Private Sub txtBabyTabRowHeight_Change()
    mblnChanged = True
End Sub

Private Sub txtBabyTabRowHeight_GotFocus()
    Me.txtTabRowHeight.SelStart = 0: Me.txtTabRowHeight.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtBabyTabRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab):  Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtBabyTitle_Change()
    txtBabyTitleText.Text = txtBabyTitle.Text
End Sub


Private Sub txtBabyTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim i As Integer, lngWidth As Long
    
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC, AddressOf WndMessage)
        Set cbrPopupBar = cbsThis.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddLeft, "在左侧新增列"): cbrPopupItem.IconId = conMenu__Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddRight, "在右侧新增列"): cbrPopupItem.IconId = conMenu__Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteCol, "删除列"): cbrPopupItem.IconId = conMenu__Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddUP, "在上方新增行"): cbrPopupItem.IconId = conMenu__Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddBottom, "在下方新增行"): cbrPopupItem.IconId = conMenu__Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteRow, "删除行"): cbrPopupItem.IconId = conMenu__Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddySingle, "绑定单列"): cbrPopupItem.IconId = conMenu__Curve_BuddySingle
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddyDouble, "绑定双列"): cbrPopupItem.IconId = conMenu__Curve_BuddyDouble
        cbrPopupBar.ShowPopup
        Call SetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub txtBabyTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And glngTXTProc <> 0 Then
        glngTXTProc = GetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub txtBabyTitleText_Change()
    DataChanged = True
    txtBabyTitle.Text = txtBabyTitleText.Text
End Sub

Private Sub txtBabyTitleText_GotFocus()
    Me.txtBabyTitleText.SelStart = 0: Me.txtTabRowHeight.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtBabyTitleText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtBabyTitleText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim i As Integer, lngWidth As Long
    
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC, AddressOf WndMessage)
        Set cbrPopupBar = cbsThis.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddLeft, "在左侧新增列"): cbrPopupItem.IconId = conMenu__Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddRight, "在右侧新增列"): cbrPopupItem.IconId = conMenu__Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteCol, "删除列"): cbrPopupItem.IconId = conMenu__Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddUP, "在上方新增行"): cbrPopupItem.IconId = conMenu__Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddBottom, "在下方新增行"): cbrPopupItem.IconId = conMenu__Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteRow, "删除行"): cbrPopupItem.IconId = conMenu__Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddySingle, "绑定单列"): cbrPopupItem.IconId = conMenu__Curve_BuddySingle
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddyDouble, "绑定双列"): cbrPopupItem.IconId = conMenu__Curve_BuddyDouble
        cbrPopupBar.ShowPopup
        Call SetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub txtBabyTitleText_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And glngTXTProc <> 0 Then
        glngTXTProc = GetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtBabyTitle.hWnd, GWL_WNDPROC, glngTXTProc)
        glngTXTProc = 0
    End If
End Sub

Private Sub txtColumnNo_Change()
    vfgThis.Col = txtColumnNo.Text
End Sub

Private Sub txtColumnPostfix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "}")) & Me.txtColumnPostfix.Text
    End With
End Sub

Private Sub txtColumnPostfix_GotFocus()
    Me.txtColumnPostfix.SelStart = 0: Me.txtColumnPostfix.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPostfix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
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
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtCurveColWidth_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtCurveColWidth_GotFocus()
    Me.txtCurveColWidth.SelStart = 0: Me.txtCurveColWidth.SelLength = 10
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtCurveColWidth_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtCurveRowHeight_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtCurveRowHeight_GotFocus()
    Me.txtCurveRowHeight.SelStart = 0: Me.txtCurveRowHeight.SelLength = 10
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtCurveRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtDownTabRowHeight_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtDownTabRowHeight_GotFocus()
    Me.txtDownTabRowHeight.SelStart = 0: Me.txtDownTabRowHeight.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtDownTabRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtDownTabRowHeight_Validate(Cancel As Boolean)
    If Val(txtDownTabRowHeight.Text) < 225 Then
        txtDownTabRowHeight.Text = 225
    ElseIf Val(txtDownTabRowHeight.Text) > 600 Then
        txtDownTabRowHeight.Text = 600
    End If
End Sub



Private Sub txtHeadText_Change()
    Dim strInput As String
    Dim lngStart As Long, lngRow As Long, lngCol As Long
    Dim blnExist As Boolean
    
    strInput = Trim(Me.txtHeadText.Text)
    '检查,如果有相邻的四个单元格的值相同,则不允许设置(有可能需要和左上,左下,右上,右下进行检查)
    lngRow = udHeadRow.Value + 1
    lngCol = Me.udHeadCol.Value
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 1, lngCol) = strInput
    
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
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 1, lngCol) = strInput
'    vfgThis.AutoSize 0, vfgThis.Cols - 1
'    Call cmdLabel_Click(2)
    Call vfgThis_AfterUserResize(Me.udHeadRow.Value + 1, lngCol)
    DataChanged = True
End Sub

Private Sub txtHeight_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtHeight.Text) Then
        txtHeight.Tag = CLng(txtHeight.Text * conRatemmToTwip)
        mlngHeight = CLng(txtHeight.Text * conRatemmToTwip)
        
        If mintPage = 256 Then cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
    '页眉页脚设置
    Call InitPageFoot
    DataChanged = True
End Sub

Private Sub txtScaleColWidth_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtScaleColWidth_GotFocus()
    Me.txtScaleColWidth.SelStart = 0: Me.txtScaleColWidth.SelLength = 10
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtScaleColWidth_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtSeachColumnItems_Change()
    txtSeachColumnItems.Tag = ""
End Sub

Private Sub txtSeachColumnItems_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intCount As Integer
    Dim intIndex As Integer
    Dim arrIndex()
    
    If KeyAscii = vbKeyReturn Then
        intCount = 0
        arrIndex = Array()
        intIndex = 0
        For i = 0 To lstColumnItems.ListCount
            If InStr(1, lstColumnItems.List(i), txtSeachColumnItems.Text) > 0 Then
                ReDim Preserve arrIndex(UBound(arrIndex) + 1)
                arrIndex(intCount) = i
                intCount = intCount + 1
                intIndex = i
            End If
        Next
        If intCount > 1 Then
            If Val(txtSeachColumnItems.Tag) < intCount Then
                txtSeachColumnItems.Tag = Val(txtSeachColumnItems.Tag) + 1
                lstColumnItems.ListIndex = arrIndex(Val(txtSeachColumnItems.Tag) - 1)
            End If
        Else
            lstColumnItems.ListIndex = intIndex
        End If
        
        
    End If
End Sub

Private Sub txtTabBeginTime_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtTabDays_Change()
    DataChanged = True
    mblnRedraw = True
End Sub

Private Sub txtTabDayTime_Change()
    '根据一天的监测次数决定开始时点和时间间隔的最大值
    Dim intHour As Integer
    Dim lngRow As Long, lngFrequency As Long
    Dim blnSetTab As Boolean
    
    '68649:刘鹏飞,2013-12-19,开始时点和时间间隔最大值的设置
    intHour = 24 \ Val(txtTabDayTime.Text)
    udTabBeginTime.Min = 0
    udTabBeginTime.Max = 24 - Val(txtTabDayTime.Text) + 1 'intHour
    If (intHour \ 2) + 1 > udTabBeginTime.Max Then
        intHour = udTabBeginTime.Max
    Else
        intHour = (intHour \ 2) + 1
    End If
    blnSetTab = (udTabBeginTime.Value = intHour)
    udTabBeginTime.Value = intHour
    txtTabBeginTime.Text = udTabBeginTime.Value
    
    If blnSetTab = True Then Call udTabBeginTime_Change
    
    '表格项目频次检查
    For lngRow = vfgTab.FixedRows To vfgTab.Rows - 1
        If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号"))) <> 0 Then
            lngFrequency = Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")))
            If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目序号"))) = 3 Then
                vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = Val(txtTabDayTime.Text)
            Else
                If InStr(1, GetTabFrequency(Val(txtTabDayTime.Text), Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("项目表示")))), lngFrequency) = 0 Then
                    vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("记录频次")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
                End If
            End If
        End If
    Next
    Call vfgTab_EnterCell
    
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_GotFocus()
    Me.txtTabRowHeight.SelStart = 0: Me.txtTabRowHeight.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtTabRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtTabRowHeight_Validate(Cancel As Boolean)
    If Val(txtTabRowHeight.Text) < 225 Then
        txtTabRowHeight.Text = 225
    ElseIf Val(txtTabRowHeight.Text) > 600 Then
        txtTabRowHeight.Text = 600
    End If
End Sub

Private Sub txtTabRowName_Change(Index As Integer)
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtTabRowName_GotFocus(Index As Integer)
    Me.txtTabRowName(Index).SelStart = 0: Me.txtTabRowName(Index).SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtTabRowName_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    'If InStr("~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr("@'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtTabTimeSplit_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtTitleText_Change()
    mblnRedraw = True
    DataChanged = True
End Sub

Private Sub txtTitleText_GotFocus()
    Me.txtTitleText.SelStart = 0: Me.txtTitleText.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtTitleText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtWidth.Text) Then
        txtWidth.Tag = CLng(txtWidth.Text * conRatemmToTwip)
        mlngWidth = CLng(txtWidth.Text * conRatemmToTwip)
        
        If mintPage = 256 Then cboPage.ListIndex = cboPage.ListCount - 1
    End If
    Call ShowPaper
    
    '页眉页脚设置
    Call InitPageFoot
    DataChanged = True
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
    End If
End Sub

Private Sub txtHeight_GotFocus()
    txtHeight.SelStart = 0: txtHeight.SelLength = Len(txtHeight.Text)
End Sub

Private Sub txtWidth_GotFocus()
    txtWidth.SelStart = 0: txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: VBA.Beep
    End If
End Sub

Private Sub txtUP_Change()
    DataChanged = True
End Sub

Private Sub txtUP_GotFocus()
    zlControl.TxtSelAll txtUP
End Sub

Private Sub txtUP_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtUP_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtUP.Text) Then
        If txtUP.Text >= UDUp.Min And txtUP.Text <= UDUp.Max Then
            UDUp.Value = txtUP.Text
        Else
            UDUp.Value = UDUp.Min
        End If
    End If
End Sub

Private Sub txtDown_Change()
    DataChanged = True
End Sub

Private Sub txtDown_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtDown.Text) Then
        If txtDown.Text >= UDDown.Min And txtDown.Text <= UDDown.Max Then
            UDDown.Value = txtDown.Text
        Else
            UDDown.Value = UDDown.Min
        End If
    End If
End Sub

Private Sub txtDown_GotFocus()
    zlControl.TxtSelAll txtDown
End Sub

Private Sub txtDown_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtRight_Change()
    DataChanged = True
End Sub

Private Sub txtRight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtRight_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtRight.Text) Then
        If txtRight.Text >= UDRight.Min And txtRight.Text <= UDRight.Max Then
            UDRight.Value = Val(txtRight.Text)
        Else
            UDRight.Value = UDRight.Min
        End If
    End If
End Sub

Private Sub txtRight_GotFocus()
    zlControl.TxtSelAll txtRight
End Sub

Private Sub txtLeft_Change()
    DataChanged = True
End Sub

Private Sub txtLeft_Validate(Cancel As Boolean)
    If Not mblnChange Then Exit Sub
    If IsNumeric(txtLeft.Text) Then
        If txtLeft.Text >= UDLeft.Min And txtLeft.Text <= UDLeft.Max Then
            UDLeft.Value = Val(txtLeft.Text)
        Else
            UDLeft.Value = UDLeft.Min
        End If
    End If
End Sub

Private Sub txtLeft_GotFocus()
    zlControl.TxtSelAll txtLeft
End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub udBabyCol_Change()
    Dim lngCount As Long
    
    Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True

    Me.vfgThis.Cols = Me.udBabyCol.Value + 1
    Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
    Me.udHeadCol.Max = Me.udBabyCol.Value
    If Val(Me.txtHeadCol.Text) > Me.udHeadCol.Max Then Me.txtHeadCol.Text = Me.udHeadCol.Max
    Me.udColumnNo.Max = Me.udBabyCol.Value
    If Val(Me.txtColumnNo.Text) > Me.udColumnNo.Max Then Me.txtColumnNo.Text = Me.udColumnNo.Max
    
'    Me.vfgThis.ColPosition(vfgThis.Cols - 1) = Val(vfgThis.Tag + 1)
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 2, .FixedCols, 5, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 1, 0, 0, 0
        .MergeCol(-1) = True
    End With
    Me.udColumnNo.Max = vfgThis.Cols - 1
    Call vfgThis_AfterUserResize(vfgThis.Row, vfgThis.Row)
    
End Sub

Private Sub udColumnNo_Change()
    Dim strTemp As String
    Dim lngCount As Long
    
    Me.lstColumnUsed.Clear
    strTemp = vfgThis.Cell(flexcpData, 5, udColumnNo.Value, 5, udColumnNo.Value)
    
    If strTemp = "" Then
        Me.lstColumnUsed.ListIndex = -1
        Me.cmdColumn(1).Enabled = False
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
        Me.cmdColumn(1).Enabled = True
        Me.txtColumnPrefix.Enabled = True
        Me.txtColumnPostfix.Enabled = True
        
        chk.Enabled = True
        
    End If
End Sub

Private Sub udHeadCol_Change()
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    txtHeadText.Text = vfgThis.TextMatrix(udHeadRow.Value + 1, udHeadCol.Value)

    DataChanged = blnSvrChanged
End Sub

Private Sub udHeadRow_Change()
    Call udHeadCol_Change
End Sub

Private Sub udTabBeginTime_Change()
    '68649:刘鹏飞,2013-12-19,开始时点和时间间隔最大值的设置
    If Val(txtTabDayTime.Text) <= 1 Then Exit Sub
    
    udTabTimeSplit.Min = 1
    udTabTimeSplit.Max = (24 - udTabBeginTime.Value) \ (Val(txtTabDayTime.Text) - 1)
    udTabTimeSplit.Value = udTabTimeSplit.Max
    txtTabTimeSplit.Text = udTabTimeSplit.Value
End Sub

Private Sub udTabDayTime_DownClick()
    If udTabDayTime.Value = 0 Then
        udTabDayTime.Value = udTabDayTime.Value + udTabDayTime.Increment
    ElseIf udTabDayTime.Value = 10 Then
        udTabDayTime.Value = udTabDayTime.Value - udTabDayTime.Increment
    ElseIf udTabDayTime.Value > 12 Then
        udTabDayTime.Value = 12
    End If
End Sub

Private Sub udTabDayTime_UpClick()
    If udTabDayTime.Value = 0 Or udTabDayTime.Value = 10 Then
        udTabDayTime.Value = udTabDayTime.Value + udTabDayTime.Increment
    ElseIf udTabDayTime.Value > 12 Then
        udTabDayTime.Value = 24
    End If
End Sub

Private Sub UDUp_Change()
    mlngTop = UDUp.Value
    Call ShowPaper
End Sub

Private Sub UDDown_Change()
    mlngBottom = UDDown.Value
    Call ShowPaper
End Sub

Private Sub UDRight_Change()
    mlngRight = UDRight.Value
    Call ShowPaper
End Sub

Private Sub UDLeft_Change()
    mlngLeft = UDLeft.Value
    Call ShowPaper
End Sub

Private Sub ShowPaper()
'功能：显示设置的纸张的预览
    On Error Resume Next
    
    picPaper.Cls
    
    picPaper.Width = mlngWidth / conRatemmToTwip
    picPaper.Height = mlngHeight / conRatemmToTwip
    picPaper.Left = (picBack.ScaleWidth - picPaper.Width) / 2
    picPaper.Top = (picBack.ScaleHeight - picPaper.Height) / 2
    
    picShadow.Width = picPaper.Width
    picShadow.Height = picPaper.Height
    
    picShadow.Left = picPaper.Left + 5
    picShadow.Top = picPaper.Top + 5
    
    picPaper.ScaleWidth = mlngWidth
    picPaper.ScaleHeight = mlngHeight
    
    picPaper.Line (0, mlngTop * conRatemmToTwip)-(picPaper.ScaleWidth, mlngTop * conRatemmToTwip), &H808080
    picPaper.Line (0, picPaper.ScaleHeight - (mlngBottom + 2) * conRatemmToTwip)-(picPaper.ScaleWidth, picPaper.ScaleHeight - (mlngBottom + 2) * conRatemmToTwip), &H808080
    
    picPaper.Line (mlngLeft * conRatemmToTwip, 0)-(mlngLeft * conRatemmToTwip, picPaper.ScaleHeight), &H808080
    picPaper.Line (picPaper.ScaleWidth - (mlngRight + 2) * conRatemmToTwip, 0)-(picPaper.ScaleWidth - (mlngRight + 2) * conRatemmToTwip, picPaper.ScaleHeight), &H808080
    
    Me.Refresh
End Sub

Private Sub GetrtbObject()
    If mblnRTBFoot Then
        Set rtbThis = rtbFoot
    Else
        Set rtbThis = rtbHead
    End If
End Sub


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

Private Function SavePageHead(ByVal StrKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:产生文件并压缩;True:已产生压缩文件
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Head_S.rtf"
        If gobjFSO.FileExists(strFile) = True Then gobjFSO.DeleteFile strFile, True
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
        If gobjFSO.FileExists(strFile) = True Then gobjFSO.DeleteFile strFile, True
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
        .Bottom = Printer.ScaleX(txtUP.Text, vbMillimeters, vbTwips)
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
        MsgBox "设计的页眉内容超过上边距！", vbInformation, gstrSysName
        Exit Function
    End If
    PageHeadTest = True
End Function

Private Function OverRun() As Boolean
    Dim intPageMargin As Integer    '边距
    Dim lngPageWidth As Long, lngPageHeight As Long      '纸张宽度、高度
    Dim lngTrimWidth  As Long, lngTrimHeight As Long       '体温实际占用的宽度和高度(含边距)
    Dim lngTabWidth As Long, lngTabHeight As Long          '婴儿体温单表格
    Dim i As Integer
    
    '检查体温单的宽度是否超出页面有效打印范围
    If mintPage = 256 And optCross.Value = True Then
        lngPageWidth = mlngHeight
        lngPageHeight = mlngWidth
    Else
        lngPageWidth = mlngWidth
        lngPageHeight = mlngHeight
    End If
    
    lngTrimWidth = WaveWidth + Int(Me.ScaleX(Val(Me.txtLeft.Text), vbMillimeters, vbTwips)) + Int(Me.ScaleX(Val(Me.txtRight.Text), vbMillimeters, vbTwips))
    
    lngTabWidth = 0
    If chkBaby.Value = 1 Then
        For i = vfgThis.FixedCols To vfgThis.Cols - 1
            lngTabWidth = lngTabWidth + vfgThis.ColWidth(i)
        Next
        lngTabWidth = Me.ScaleX(Val(Me.txtBabyLeft.Text), vbMillimeters, vbTwips) + lngTabWidth
        For i = vfgThis.FixedRows To vfgThis.Rows - 1
            lngTabHeight = lngTabHeight + vfgThis.ROWHEIGHT(i)
        Next
        lngTabHeight = lngTabHeight + vfgThis.ROWHEIGHT(5) * 6
        
    End If
    
    lngTrimWidth = IIf(WaveWidth < lngTabWidth, lngTabWidth, WaveWidth) + Int(Me.ScaleX(Val(Me.txtLeft.Text), vbMillimeters, vbTwips)) + Int(Me.ScaleX(Val(Me.txtRight.Text), vbMillimeters, vbTwips))
    lngTrimHeight = WaveHeight + Int(Me.ScaleX(Val(Me.txtUP.Text), vbMillimeters, vbTwips)) + Int(Me.ScaleX(Val(Me.txtDown.Text), vbMillimeters, vbTwips))
    
    If lngTabWidth > lngPageWidth - 100 Then
        MsgBox "婴儿体温单表格输出的实际宽度大于了纸张的宽度,请调整!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngTrimWidth > lngPageWidth - 100 Then
        MsgBox "体温单输出的实际宽度大于了纸张的宽度,请调整!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngTrimHeight > lngPageHeight - 100 Then
        MsgBox "体温单输出的实际高度大于了纸张的高度,请调整!", vbInformation, gstrSysName
        Exit Function
    End If
    
    OverRun = True
End Function

Private Sub zlFontSet(strTitle As String, strFont As String)
    With Me.dlgThis
        .flags = &H3 Or &H400 Or &H200 Or &H10000
        .DialogTitle = strTitle
        .FontName = Split(strFont, ",")(0)
        .FontSize = Val(Split(strFont, ",")(1))
        If InStr(1, strFont, "粗") > 0 Then
            .FontBold = True
        Else
            .FontBold = False
        End If
        If InStr(1, strFont, "斜") > 0 Then
            .FontItalic = True
        Else
            .FontItalic = False
        End If
        Err = 0: On Error Resume Next
        .ShowFont
        .flags = 0
        If Err.Number <> 0 Then Exit Sub
        strFont = .FontName & "," & .FontSize
        If .FontBold Or .FontItalic Then
            strFont = strFont & "," & IIf(.FontBold, "粗", "") & IIf(.FontItalic, "斜", "")
        End If
    End With
End Sub

Private Sub SetVsFlexGridChangeHead(ByVal strHead As String, ByRef vsgrid As VSFlexGrid, lngNO As Long)
    '功能：初始vsFlexGrid
    '           有一固定行，初始化后，只有一行记录，无固定列。
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'vsGrid:    要初始化的控件

    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgrid
        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1
        If lngNO = 0 Then
            .FixedCols = 0
            .Cols = .FixedCols + UBound(arrHead) + 1
            .Rows = .FixedRows + 1
        Else
            .FixedCols = 1
            .Cols = .FixedCols + UBound(arrHead)
            .Rows = .FixedRows + 1
        End If

        For i = 0 To UBound(arrHead)
            If .FixedCols > 0 Then
                .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            Else
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            End If
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               '为了支持zl9PrintMode
                If .FixedCols > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                    .Cell(flexcpAlignment, .FixedRows, i, .Rows - 1, i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(.FixedCols + i) = False
                    .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
'                    .ColData
                    '为了支持zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  '为了支持zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
                End If
            End If
            .ColData(i) = Split(arrHead(i), ",")(3) '将标提作为列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        Next
        
        '固定行文字居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .ROWHEIGHT(0) = 300
        
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
    End With
End Sub

Private Sub vfgCurve_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vfgCurve
        If Row >= .FixedRows And Col >= .FixedCols Then
            If .ColIndex("选择") = Col And Val(.TextMatrix(Row, .ColIndex("项目序号"))) <> 0 Then
                mblnRedraw = True
                DataChanged = True
            End If
        End If
    End With
End Sub

Private Sub vfgCurve_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vfgCurve.ColIndex("序号") Or Col = vfgCurve.ColIndex("选择") Then
        Cancel = True
    End If
End Sub

Private Sub vfgCurve_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long, strValue As String
    With vfgCurve
        If KeyCode = vbKeyReturn Then
            If .Row >= .Rows - 1 Then
                zlCommFun.PressKey vbKeyTab
                Exit Sub
            Else
                .Row = .Row + 1
                .Col = .ColIndex("选择")
                If .CellLeft + .CellWidth > .Width Then .LeftCol = .Col
                If .CellTop + .CellHeight > .Height Then .TopRow = .Row
                If .Enabled And .Visible Then .SetFocus
            End If
        End If
    End With
End Sub

Private Sub vfgCurve_KeyPress(KeyAscii As Integer)
    Dim strValue As String
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    With vfgCurve
        If .Row < .FixedRows And .Col < .FixedCols Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("项目序号"))) = 0 Then Exit Sub
        If KeyAscii = vbKeySpace Then
            If Val(.TextMatrix(.Row, .ColIndex("项目序号"))) = 1 Then
                .TextMatrix(.Row, .ColIndex("选择")) = 1
                Exit Sub
            Else
                strValue = .TextMatrix(.Row, .ColIndex("选择"))
                .TextMatrix(.Row, .ColIndex("选择")) = IIf(Val(strValue) = 1, 0, 1)
            End If
            DataChanged = True
            mblnRedraw = True
            On Error Resume Next
            If .Enabled And .Visible Then .SetFocus
        End If
    End With
End Sub

Private Sub vfgCurve_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   If Col <> vfgCurve.ColIndex("选择") Then Cancel = True: Exit Sub
   If Trim(vfgCurve.TextMatrix(Row, vfgCurve.ColIndex("项目序号"))) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vfgCurve_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgCurve
        If Row >= .FixedRows And Col >= .FixedCols Then
            If .ColIndex("选择") = Col And Val(.TextMatrix(Row, .ColIndex("项目序号"))) = 1 Then
                .TextMatrix(Row, .ColIndex("选择")) = 1
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vfgTab_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vfgTab
        If Not (Row >= .FixedRows And Col = .ColIndex("记录频次")) Then Exit Sub
        If .ComboIndex < 0 Then Exit Sub
        .TextMatrix(Row, Col) = .EditText
    End With
End Sub

Private Sub vfgTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim StrKey As String
    With vfgTab
        If Not (NewRow >= .FixedRows And NewCol = .ColIndex("记录频次")) Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Editable = flexEDKbdMouse Then
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusNone
        End If
    End With
End Sub

Private Sub vfgTab_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vfgTab.ColIndex("序号") Then
        Cancel = True
    End If
End Sub

Private Sub vfgTab_EnterCell()
    With vfgTab
        .ColEditMask(.Col) = ""
        .ColComboList(.Col) = ""
        .CellBorderRange .FixedRows, .FixedCols, .Rows - 1, .Cols - 1, .GridColor, 1, 1, 1, 1, 0, 0
        If (.Row >= .FixedRows And .Col = .ColIndex("记录频次")) Then
            If Val(.TextMatrix(.Row, .ColIndex("项目序号"))) = 3 Then Exit Sub
            .CellBorderRange .Row, .Col, .Row, .Col, &HFF0000, 1, 1, 1, 1, 0, 0
            '根据监测次数决定记录频次
            .ColComboList(.Col) = GetTabFrequency(Val(txtTabDayTime.Text), Val(.TextMatrix(.Row, .ColIndex("项目表示"))))
        End If
    End With
End Sub

Private Sub vfgTab_KeyDown(KeyCode As Integer, Shift As Integer)
    With vfgTab
        If KeyCode = vbKeyReturn Then
            If Not .Col = .ColIndex("记录频次") Then
                .Col = .ColIndex("记录频次")
            Else
                If .Row >= .Rows - 1 Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    .Row = .Row + 1
                    .Col = .ColIndex("记录频次")
                    If .CellLeft + .CellWidth > .Width Then .LeftCol = .Col
                    If .CellTop + .CellHeight > .Height Then .TopRow = .Row
                End If
            End If
            If .Enabled And .Visible Then .SetFocus
        End If
    End With
End Sub

Private Sub vfgTab_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vfgTab_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgTab
        '呼吸不能选择
        If .TextMatrix(Row, .ColIndex("项目序号")) = 3 Then Cancel = True: Exit Sub
        If Col <> .ColIndex("记录频次") Then Cancel = True: Exit Sub
        If Trim(.TextMatrix(Row, .ColIndex("项目序号"))) = "" Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vfgTab_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgTab
        If Col <> .ColIndex("记录频次") Then Exit Sub
        If .TextMatrix(Row, Col) = .EditText Then Exit Sub
    End With
    mblnRedraw = True
    DataChanged = True
End Sub


Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：
    '参数：
    '返回： 调用成功返回TRUE；否则FALSE
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    vsb.Value = 0: hsb.Value = 0
    picDraw.Top = 0: picDraw.Left = 0
    
    With Picbaby
        .Left = 0
        .Top = picDraw.Top + picDraw.Height
        .Height = 4500
        .Width = picPane(1).Width
    End With
    
    hsb.Max = picDraw.Width - picPane(1).Width
    vsb.Max = picDraw.Height + IIf(Picbaby.Visible = True, Picbaby.Height, 0) - picPane(1).Height
    hsb.Enabled = (hsb.Max > 0)
    hsb.Visible = hsb.Enabled
    If hsb.Visible = True Then hsb.ZOrder 0
    vsb.Enabled = (vsb.Max > 0)
    vsb.Visible = vsb.Enabled
    If vsb.Visible = True Then vsb.ZOrder 0
    
    With vsb
        .Height = picPane(1).Height
    End With
    
    With hsb
        .Width = picPane(1).Width - IIf(vsb.Visible = True, vsb.Width, 0)
    End With
    
    '只根据没显示出来的那部分来计算步长
    msinHStep = (picDraw.Width - picPane(1).Width + IIf(vsb.Visible = True, vsb.Width, 0)) / 10
    msinVStep = (picDraw.Height + IIf(Picbaby.Visible = True, Picbaby.Height, 0) - picPane(1).Height + IIf(hsb.Visible = True, hsb.Height, 0)) / 10
    
    '恒定为100,只是步长发生变化
    If hsb.Enabled Then
        hsb.Max = 10
        hsb.LargeChange = 10 / Int((Round((picDraw.Width - picPane(1).Width + IIf(vsb.Visible = True, vsb.Width, 0)) / picPane(1).Width, 2) + 1))
        hsb.SmallChange = hsb.LargeChange / 2
    End If
    
    If vsb.Enabled Then
        vsb.Max = 10
        vsb.LargeChange = 10 / Int((Round((picDraw.Height + Picbaby.Height - picPane(1).Height + IIf(hsb.Visible = True, hsb.Height, 0)) / picPane(1).Height, 2) + 1))
        vsb.SmallChange = vsb.LargeChange / 2
    End If
    
    CalcScrollBarSize = True
End Function

Private Sub vfgThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    Me.udHeadCol.Value = NewCol
    Me.udColumnNo.Value = NewCol
    If NewRow >= 2 And NewRow <= 4 Then udHeadRow.Value = NewRow - 1
    
    DataChanged = blnSvrChanged
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    Dim longWidth As Long
    longWidth = 0
    For lngCol = 1 To vfgThis.Cols - 1
        longWidth = vfgThis.ColWidth(lngCol) + longWidth
    Next
    With txtBabyTitle
        .Left = vfgThis.ColWidth(0)
        .Top = vfgThis.ROWHEIGHT(0) + vfgThis.Top
        .Height = vfgThis.ROWHEIGHT(1) - 20
        .Width = longWidth
    End With
    DataChanged = True
End Sub

Private Sub vfgThis_Click()
    udColumnNo.Value = vfgThis.Col
End Sub

Private Sub vfgThis_DblClick()
    If vfgThis.Row = 5 Then
        picCloumn.Visible = True
        txtSeachColumnItems.Text = ""
        If InStr(2, vfgThis.TextMatrix(vfgThis.Row, vfgThis.Col), "{") > 0 Then
            picCloumn.Width = lstColumnUsed.Left + lstColumnUsed.Width + 100
            picCloumn.Tag = 2
        Else
            picCloumn.Width = cmdColumn(0).Left
            picCloumn.Tag = 1
        End If
        picCloumn.Top = picDraw.Height - picCloumn.Height - 1 * vsb.Value * msinVStep
        picCloumn.Left = vfgThis.CellLeft
        udColumnNo.Value = vfgThis.Col
        txtColumnNo.Text = vfgThis.Col
    End If
End Sub

Private Sub vfgThis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call cmdColumn_Click(1)
        Call cmdColumn_Click(2)
    End If
End Sub

Private Sub vfgThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim i As Integer, lngWidth As Long
    
    If Button = 2 Then
        vfgThis.Tag = vfgThis.Col
        Set cbrPopupBar = cbsThis.Add("右键菜单", xtpBarPopup)
        cbrPopupBar.Title = "右键菜单"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddLeft, "在左侧新增列"): cbrPopupItem.IconId = conMenu__Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddRight, "在右侧新增列"): cbrPopupItem.IconId = conMenu__Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteCol, "删除列"): cbrPopupItem.IconId = conMenu__Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddUP, "在上方新增行"): cbrPopupItem.IconId = conMenu__Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddBottom, "在下方新增行"): cbrPopupItem.IconId = conMenu__Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteRow, "删除行"): cbrPopupItem.IconId = conMenu__Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddySingle, "绑定单列"): cbrPopupItem.IconId = conMenu__Curve_BuddySingle
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddyDouble, "绑定双列"): cbrPopupItem.IconId = conMenu__Curve_BuddyDouble
        cbrPopupBar.ShowPopup
        
    End If
End Sub

Private Sub vfgThis_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row = vfgThis.Rows - 1 Then
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

Private Sub vsb_Change()
    picDraw.Top = -1 * vsb.Value * msinVStep
    picCloumn.Top = picDraw.Height - picCloumn.Height - 1 * vsb.Value * msinVStep
    If Picbaby.Visible = True Then Picbaby.Top = picDraw.Height - 1 * vsb.Value * msinVStep
End Sub

Private Sub hsb_Change()
    picDraw.Left = -1 * hsb.Value * msinHStep
    If Picbaby.Visible = True Then Picbaby.Left = -1 * hsb.Value * msinHStep
End Sub


Private Sub zlColorSet(strTitle As String, lngColor As Long)
    With Me.dlgThis
        .DialogTitle = strTitle
        .COLOR = lngColor
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        lngColor = .COLOR
    End With
End Sub

Private Function CheckData() As Boolean
    Dim arrData
    Dim intType As Integer, intFace As Integer, intLen As Integer                  '两个项目的表示方式一样才允许
    Dim bln护士 As Boolean, bln日期 As Boolean, bln时间 As Boolean
    Dim StrText As String, strItem As String
    Dim lngCol As Long, lngCount As Long
    Dim intDo As Integer, intHead As Integer
    Dim intRow As Integer
    
    '每种护理记录单都必须要有一列绑定护士才行
    lngCount = vfgThis.Cols - 1
   
    If Replace(vfgThis.TextMatrix(5, 1), " ", "") <> "{日期}" Then
        MsgBox "第一列必须绑定日期项目！", vbInformation, gstrSysName
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
        If Replace(vfgThis.TextMatrix(5, lngCol), " ", "") = "{时间}" And lngCol <> 2 Then
            MsgBox "要绑定时间必须绑定在第二列！", vbInformation, gstrSysName
            Exit Function
        End If
        If vfgThis.Cell(flexcpData, 5, lngCol) <> "" Then
            StrText = Val(Split(vfgThis.Cell(flexcpData, 5, lngCol), "`")(1))
            If InStr(1, vfgThis.Cell(flexcpData, 5, lngCol), "收缩压") > 0 Then
                If Not InStr(1, vfgThis.Cell(flexcpData, 5, lngCol), "舒张压") > 0 Then
                    MsgBox "要绑定收缩压必须绑定舒张压！格式为:收缩压/舒张压", vbInformation, gstrSysName
                End If
            End If
            If StrText = 1 Then
                '格式：{A}{B}，按}分解，2列分解出来的数组就=2
                StrText = vfgThis.TextMatrix(5, lngCol)
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
                    mrsItems.Filter = "项目名称='" & strItem & "'"
                    '录入的时候检查过了,系统项如日期,时间等不允许与固定项绑定,所以,此时不存在找不到的情况
                    If mrsItems.RecordCount = 0 Then
                        MsgBox "项目:" & strItem & "已改名或已经被删除！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If intDo > 0 Then
                        If Not (intFace = mrsItems!项目表示 And intType = mrsItems!项目类型) Then
                            MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 绑定的两个项目的编辑方式必须一致！", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If mrsItems!项目类型 = 1 And mrsItems!项目表示 = 0 And NVL(mrsItems!项目长度, 1) > 3 Then
                            MsgBox "第" & lngCol & "列 " & vfgThis.TextMatrix(intHead, lngCol) & " 绑定的两个文本项目长度不能大于3！", vbInformation, gstrSysName
                            Exit Function
                        End If
                    Else
                        intFace = mrsItems!项目表示
                        intType = mrsItems!项目类型
                        intLen = NVL(mrsItems!项目长度, 1)
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
    
    mrsItems.Filter = 0
    CheckData = True
End Function

Private Function GetItemName(ByVal strData As String, ByVal intOrder As Integer) As String
    Dim intDo As Integer, intPos As Integer
    '获取指定格式串中指定序号的项目名称，格式如：{收缩压}/ {舒张压}mmHg
    
    intPos = InStr(1, strData, "{")
    If intOrder > 0 Then intPos = InStr(intPos + 1, strData, "{")
    strData = Mid(strData, intPos + 1)
    strData = Mid(strData, 1, InStr(1, strData, "}") - 1)
    GetItemName = strData
End Function

