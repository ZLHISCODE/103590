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
   Caption         =   "���µ�����"
   ClientHeight    =   10350
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   16200
   Icon            =   "frmTendWaveStyle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   16200
   StartUpPosition =   1  '����������
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
            Caption         =   "�ı�����(&F)"
            Height          =   300
            Left            =   390
            TabIndex        =   97
            Top             =   3450
            Width           =   1215
         End
         Begin VB.CommandButton cmdTabTextColor 
            Caption         =   "�ı���ɫ(&R)"
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
            Text            =   "Ӥ��ÿ�ռ�¼"
            Top             =   2655
            Width           =   2805
         End
         Begin VB.CommandButton cmdBybyTitleFont 
            Caption         =   "��������(&T)"
            Height          =   300
            Left            =   3270
            TabIndex        =   92
            Top             =   2655
            Width           =   1260
         End
         Begin VB.OptionButton optTabTiers 
            Caption         =   "��(&1)"
            Height          =   180
            Index           =   0
            Left            =   1185
            TabIndex        =   91
            Top             =   1710
            Value           =   -1  'True
            Width           =   780
         End
         Begin VB.OptionButton optTabTiers 
            Caption         =   "˫(&2)"
            Height          =   180
            Index           =   1
            Left            =   1980
            TabIndex        =   90
            Top             =   1710
            Width           =   780
         End
         Begin VB.OptionButton optTabTiers 
            Caption         =   "��(&3)"
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
            Caption         =   "�����ɫ(&G)"
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
            Caption         =   "��ͷ��Ԫ"
            Height          =   180
            Left            =   210
            TabIndex        =   173
            Top             =   255
            Width           =   720
         End
         Begin VB.Label lblHeadRow 
            AutoSize        =   -1  'True
            Caption         =   "���"
            Height          =   180
            Left            =   3390
            TabIndex        =   172
            Top             =   690
            Width           =   360
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "�к�"
            Height          =   180
            Left            =   3390
            TabIndex        =   171
            Top             =   1095
            Width           =   360
         End
         Begin VB.Label lblHeadText 
            AutoSize        =   -1  'True
            Caption         =   "�ı�"
            Height          =   180
            Left            =   4395
            TabIndex        =   170
            Top             =   825
            Width           =   360
         End
         Begin VB.Label lblleft 
            AutoSize        =   -1  'True
            Caption         =   "�����������߲��ֵ���Ծ���          mm"
            Height          =   180
            Left            =   420
            TabIndex        =   112
            Top             =   2070
            Width           =   3600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Left            =   210
            TabIndex        =   111
            Top             =   3090
            Width           =   720
         End
         Begin VB.Label lblBabyStyle 
            AutoSize        =   -1  'True
            Caption         =   "Ӥ�����µ�����(�±��)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "����,9"
            Height          =   180
            Left            =   1650
            TabIndex        =   109
            Top             =   3555
            Width           =   1875
         End
         Begin VB.Label lblTabTextColor 
            Caption         =   "�ı���ɫ"
            Height          =   180
            Left            =   1650
            TabIndex        =   108
            Top             =   3915
            Width           =   1635
         End
         Begin VB.Label lblBabyBasic 
            AutoSize        =   -1  'True
            Caption         =   "������̬"
            Height          =   180
            Left            =   210
            TabIndex        =   107
            Top             =   1455
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "�����ı�"
            Height          =   180
            Left            =   210
            TabIndex        =   106
            Top             =   2355
            Width           =   720
         End
         Begin VB.Label lblBabyTitleFont 
            Caption         =   "����,20"
            Height          =   180
            Left            =   4590
            TabIndex        =   105
            Top             =   2730
            Width           =   1695
         End
         Begin VB.Label lblTabTiers 
            AutoSize        =   -1  'True
            Caption         =   "��ͷ����"
            Height          =   180
            Left            =   420
            TabIndex        =   104
            Top             =   1710
            Width           =   720
         End
         Begin VB.Label lblBabyTabRowHeight 
            AutoSize        =   -1  'True
            Caption         =   "��С�и�"
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
            Caption         =   "�к�"
            Height          =   180
            Left            =   420
            TabIndex        =   102
            Top             =   690
            Width           =   360
         End
      End
      Begin VB.CheckBox chkBaby 
         Caption         =   "Ӥ�����µ�"
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
            Caption         =   "������Ŀ��(�±��)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "��ѡ������Ŀ:"
            Height          =   180
            Left            =   375
            TabIndex        =   141
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label lblSelectColumnItems 
            AutoSize        =   -1  'True
            Caption         =   "��ѡ������Ŀ:"
            Height          =   180
            Left            =   2715
            TabIndex        =   140
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label lblDownTabRowHeight 
            AutoSize        =   -1  'True
            Caption         =   "��С�߶�"
            Height          =   180
            Left            =   2700
            TabIndex        =   139
            Top             =   3255
            Width           =   720
         End
         Begin VB.Label lblAddNullTab 
            AutoSize        =   -1  'True
            Caption         =   "������"
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
         Text            =   "ʱ       ��"
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
         Text            =   "����������"
         Top             =   1365
         Width           =   1290
      End
      Begin VB.TextBox txtTabRowName 
         Height          =   300
         Index           =   1
         Left            =   2490
         MaxLength       =   10
         TabIndex        =   124
         Text            =   "סԺ����"
         Top             =   1365
         Width           =   1290
      End
      Begin VB.TextBox txtTabRowName 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   10
         TabIndex        =   123
         Text            =   "��       ��"
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
         Text            =   "����ר�����µ�"
         Top             =   420
         Width           =   2895
      End
      Begin VB.CommandButton cmdTitleFont 
         Caption         =   "��������(&T)"
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
         Caption         =   "���߱���������"
         Height          =   180
         Left            =   420
         TabIndex        =   163
         Top             =   3435
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "��Ŀѡ��:"
         Height          =   180
         Left            =   2760
         TabIndex        =   162
         Top             =   2070
         Width           =   810
      End
      Begin VB.Label lblCurveRowHeight 
         AutoSize        =   -1  'True
         Caption         =   "���߱����С�߶�"
         Height          =   180
         Left            =   420
         TabIndex        =   161
         Top             =   3090
         Width           =   1440
      End
      Begin VB.Label lblCurveColWidth 
         AutoSize        =   -1  'True
         Caption         =   "���߱����С���"
         Height          =   180
         Left            =   420
         TabIndex        =   160
         Top             =   2745
         Width           =   1440
      End
      Begin VB.Label lblScaleColWidth 
         AutoSize        =   -1  'True
         Caption         =   "�̶�����С�ܿ��"
         Height          =   180
         Left            =   420
         TabIndex        =   159
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label lblTabRowName 
         AutoSize        =   -1  'True
         Caption         =   "��ͷ����"
         Height          =   180
         Left            =   420
         TabIndex        =   158
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label lblTitleText 
         AutoSize        =   -1  'True
         Caption         =   "�����ı�"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "һ����Ŀ��(�ϱ��)"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ʱ����"
         Height          =   180
         Left            =   4950
         TabIndex        =   155
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTabBeginTime 
         AutoSize        =   -1  'True
         Caption         =   "��ʼʱ��"
         Height          =   180
         Left            =   3450
         TabIndex        =   154
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTabDays 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   420
         TabIndex        =   153
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTabRowHeight 
         AutoSize        =   -1  'True
         Caption         =   "��С�߶�"
         Height          =   180
         Left            =   420
         TabIndex        =   152
         Top             =   1755
         Width           =   720
      End
      Begin VB.Label lblTabDayTime 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   1920
         TabIndex        =   151
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTitleFont 
         Caption         =   "����,20"
         Height          =   180
         Left            =   4650
         TabIndex        =   150
         Top             =   525
         Width           =   1605
      End
      Begin VB.Label lblRecordStyle 
         AutoSize        =   -1  'True
         Caption         =   "����������(����)"
         BeginProperty Font 
            Name            =   "����"
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
            Caption         =   "ѡ��(&S)"
            Height          =   300
            Index           =   0
            Left            =   2760
            TabIndex        =   73
            Top             =   1125
            Width           =   1100
         End
         Begin VB.CommandButton cmdColumn 
            Caption         =   "ɾ��(&E)"
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
            Caption         =   "Ӧ��(&Y)"
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
            Caption         =   "�Խ���"
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
            Caption         =   "����"
            Height          =   180
            Left            =   270
            TabIndex        =   82
            Top             =   780
            Width           =   360
         End
         Begin VB.Label lblBaby 
            AutoSize        =   -1  'True
            Caption         =   "��ѡ�����¼��Ŀ:"
            Height          =   180
            Left            =   240
            TabIndex        =   80
            Top             =   420
            Width           =   1530
         End
         Begin VB.Label lblColumnNo 
            AutoSize        =   -1  'True
            Caption         =   "��        ��������Ŀ:"
            Height          =   180
            Left            =   4005
            TabIndex        =   79
            Top             =   420
            Width           =   1890
         End
         Begin VB.Label lblColumnPrefix 
            AutoSize        =   -1  'True
            Caption         =   "ǰ׺"
            Height          =   180
            Left            =   3990
            TabIndex        =   78
            Top             =   2475
            Width           =   360
         End
         Begin VB.Label lblColumnPostfix 
            AutoSize        =   -1  'True
            Caption         =   "��׺"
            Height          =   180
            Left            =   3990
            TabIndex        =   77
            Top             =   2850
            Width           =   360
         End
         Begin VB.Label lbl�ж��� 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   3990
            TabIndex        =   76
            Top             =   3240
            Width           =   360
         End
         Begin VB.Label lblColText 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
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
            Text            =   "Ӥ��ÿ�ռ�¼"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   24686
            Text            =   "���Ը���ҽԺʵ����������õ������µ���ʽ����ӡ��ҳüҳ�š�"
            TextSave        =   "���Ը���ҽԺʵ����������õ������µ���ʽ����ӡ��ҳüҳ�š�"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
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
            Caption         =   "�߾�(mm)"
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
               Caption         =   "��"
               Height          =   180
               Left            =   1695
               TabIndex        =   27
               Top             =   660
               Width           =   180
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               Height          =   180
               Left            =   1695
               TabIndex        =   21
               Top             =   330
               Width           =   180
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               Height          =   180
               Left            =   510
               TabIndex        =   18
               Top             =   330
               Width           =   180
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               Height          =   180
               Left            =   510
               TabIndex        =   24
               Top             =   675
               Width           =   180
            End
         End
         Begin VB.Frame fraOrient 
            Caption         =   "ֽ��"
            Height          =   1065
            Left            =   2925
            TabIndex        =   30
            Top             =   1755
            Width           =   1425
            Begin VB.OptionButton optPortrait 
               Caption         =   "����"
               Height          =   285
               Left            =   675
               TabIndex        =   31
               Top             =   315
               Value           =   -1  'True
               Width           =   660
            End
            Begin VB.OptionButton optCross 
               Caption         =   "����"
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
            Caption         =   "ֽ��"
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
               Caption         =   "��С"
               Height          =   180
               Left            =   285
               TabIndex        =   7
               Top             =   300
               Width           =   360
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���"
               Height          =   180
               Left            =   300
               TabIndex        =   9
               Top             =   675
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�߶�"
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
            Caption         =   "ע��:  ���ʵ�ʴ�ӡ���͵�ǰ��ӡ�����������ܵ���ֽ������ʧЧ��"
            Height          =   180
            Left            =   135
            TabIndex        =   37
            Top             =   2985
            Width           =   5490
         End
         Begin VB.Label lblNote 
            AutoSize        =   -1  'True
            Caption         =   "���µ�����ӡ������"
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
            Caption         =   "ͬ��(&G)"
            Height          =   350
            Left            =   5730
            TabIndex        =   48
            ToolTipText     =   "���л����ļ���ҳüҳ���뵱ǰ�ļ���ҳüҳ�Ÿ�ʽһ��"
            Top             =   1530
            Width           =   1100
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "��ͼ(&I)"
            Height          =   350
            Left            =   135
            TabIndex        =   42
            Top             =   1530
            Width           =   1710
         End
         Begin VB.CheckBox chkI 
            BeginProperty Font 
               Name            =   "����"
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
            ToolTipText     =   "б��(Alt+I)"
            Top             =   1530
            Width           =   345
         End
         Begin VB.CheckBox chkU 
            BeginProperty Font 
               Name            =   "����"
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
            ToolTipText     =   "�»���(Alt+U)"
            Top             =   1530
            Width           =   345
         End
         Begin VB.CheckBox chkB 
            BeginProperty Font 
               Name            =   "����"
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
            ToolTipText     =   "����(Alt+B)"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "����"
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
         Caption         =   "ҳüҳ��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��ӡ����"
         BeginProperty Font 
            Name            =   "����"
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

'ҳüҳ�����
'######################################################################################################
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'����
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'�������ڸ�ʽ��ָ���豸�������Ϣ
Private Type FORMATRANGE
    hDC As Long             '��Ⱦ�豸
    hdcTarget As Long       'Ŀ���豸
    rc As RECT              '��Ⱦ���򣬵�λ��羡�
    rcPage As RECT          '��Ⱦ�豸���������򣬵�λ��羡�
    chrg As CHARRANGE       '���ڸ�ʽ�����ı���Χ��
End Type

Private Type PageInfo
    PageNumber As Long      'ҳ��
    Start As Long           '�ַ���ʼλ��
    End As Long             '�ַ���ֹλ��
    ActualHeight As Long    '��ҳʵ�ʴ�ӡ�߶�
End Type
Private AllPages() As PageInfo   'ҳ��Ϣ
Private Const WM_PASTE = &H302&              'ճ��
Private Const WM_USER = &H400                'ͨ���� WM_USER + X ���Զ�����Ϣ
Private Const EM_FORMATRANGE = (WM_USER + 57)    'Ϊĳһ�豸��ʽ��ָ����Χ���ı���
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '�����������������õ�Ŀ���豸���п�
Private Const EM_HIDESELECTION = (WM_USER + 63)  '��ʾ/�����ı���
Private Const PHYSICALOFFSETX = 112  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ�����Ե���ɴ�ӡ��������Ե�ľ��룬�����豸��λ��
Private Const PHYSICALOFFSETY = 113  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ���ϱ�Ե���ɴ�ӡ������ϱ�Ե�ľ��룬�����豸��λ��
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '��ȡ��Ӣ�Ļ���ַ�������
'######################################################################################################

Private mdblW As Double  '��߲��ɴ�ӡ����
Private mdblH As Double  '�ϱ߲��ɴ�ӡ����
Private msinVStep As Single      '�������Ĳ���
Private msinHStep As Single      '�������Ĳ���
'��ӡ��������
Private mintPage As Integer 'ֽ��
Private mlngWidth As Long '�Զ���ֽ�ſ��,Twip
Private mlngHeight As Long '�Զ���ֽ�Ÿ߶�'Twip
Private mintOrient As Integer   'ֽ��
Private mlngLeft As Long '��߾�'mm
Private mlngRight As Long '�ұ߾�'mm
Private mlngTop As Long '�ϱ߾�'mm
Private mlngBottom As Long '�±߾�'mm
Private mblnRTBFoot As Boolean
'�¼�����
Private mblnChange As Boolean  '���ƴ�ӡ����
Private mblnChanged As Boolean '��¼�����Ƿ����仯
Private mblnRedraw As Boolean '��¼�Ƿ���Ҫ���»�ͼ
Private rtbThis As Object
Public mbytMode As Byte
Public mlngFileID As Long  '�����ļ��б��ID

Private Type TabItemCol
    ItemNO As String '��Ŀ���
    ItemName As String '��Ŀ����
    ItemUnit As String '��Ŀ��λ
    ItemShow As Integer  '��Ŀ��ʾ
    ItemFrequency As String '��¼Ƶ��
End Type

Private strCurFont As String
Private objFont As StdFont
Private mbln����Ӧ�÷�ʽ As Boolean
Private mrsItems As New ADODB.Recordset
'--�޸�˵����50182,������,2012-08-24,�������µ�����ҳüҳ�Ź���
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
    
    'ֽ��
    Select Case cboPage.ItemData(cboPage.ListIndex)
    Case 256
        'ǿ�������Զ���ֽ�ſ���,�����
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
    '���ʵ������ֽ�Ŵ�С(ֽ��Ӱ��֮��)
    Select Case mintPage
    Case 256
        '�Զ���ֽ����Ϊȫ�����Դ�ӡ
        mdblW = 0
        mdblH = 0
        
'        If cboPage.Text = "B5, 182 x 257 ����" Then
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
        'ȡ�ô�ӡ��֧�ָ÷������ʵ�ߴ�
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
        
        '���ɴ�ӡ�������
        mdblW = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX) / GetDeviceCaps(Printer.hDC, PHYSICALWIDTH)
        mdblH = GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY) / GetDeviceCaps(Printer.hDC, PHYSICALHEIGHT)
        
        txtWidth.Enabled = False
        txtHeight.Enabled = False
        UDWidth.Enabled = False
        UDHeight.Enabled = False
    
    End Select
        
    '��ʾֽ�ųߴ�
    mblnChange = False
    txtWidth.Tag = mlngWidth
    txtWidth.Text = CLng(mlngWidth / conRatemmToTwip)
    txtHeight.Tag = mlngHeight
    txtHeight.Text = CLng(mlngHeight / conRatemmToTwip)
    mblnChange = True
    
    '��ʾ���ñ߾�
    '��С�ڿɴ�ӡ����֮��
    '��󲻳�����ߵ�1/4
'    If cboPage.Text = "B5, 182 x 257 ����" Then
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
    
    '��ʾֽ��
    mblnChange = False
    If mintOrient = 1 Then
        optPortrait.Value = True: optPortrait_Click
    Else
        optCross.Value = True: optCross_Click
    End If
    mblnChange = True
    
    '��ʾԤ��ֽ��
    Call ShowPaper
    'ҳüҳ������
    Call InitPageFoot
    DataChanged = True
End Sub

Private Sub LoadPage()
    Dim i As Integer
    Dim strPrinter As String
    
    '��ʼ��ӡ���б�
    strPrinter = GetSetting("ZLSOFT", "����ģ��\zl9PrintMode\Default", "DeviceName", Printer.DeviceName)
    With cboPrinter
        .Clear
        For i = 0 To Printers.Count - 1
            .AddItem Printers(i).DeviceName
            .ItemData(.ListCount - 1) = i '��ӡ������
            
            '��ȡ�洢�Ĵ�ӡ��Ϊ��ǰ��ӡ��,����ʼ������ҳ��
            If strPrinter = Printers(i).DeviceName Then .ListIndex = .NewIndex
        Next
        
        'ȱʡ��ʼ��Ϊ��ǰ��ӡ��
        If .ListIndex = -1 Then
            For i = 0 To .ListCount - 1
                '��ȡϵͳ��ǰ�Ĵ�ӡ��Ϊ��ǰ��ӡ��,����ʼ������ҳ��
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
     '���֧��,�򱣳�ԭ��ֽ��
     If mintPage <> 256 Then
         On Error Resume Next
         Printer.PaperSize = mintPage
         On Error GoTo 0
         mintPage = Printer.PaperSize
         mintOrient = Printer.Orientation
     End If
     
     '���⴦���������µ�ֻ֧��A4��B5��С��ֽ��
     cboPage.Clear
     '------------------------------------------------------------------------------------------
     'ֽ�Ŵ�С
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
             If j >= 1 And j <= 41 Then 'ֻ�г���׼֧�ֵ�ֽ��
                 cboPage.AddItem GetPaperName(j)
                 cboPage.ItemData(cboPage.ListCount - 1) = j
                 If j = mintPage Then cboPage.ListIndex = cboPage.NewIndex
             End If
         End If
         
     Next
    
     '------------------------------------------------------------------------------------------
     '�Զ���ֽ�Ŵ���
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
        MsgBox "��ȷ����ӡ����ֽ�ſ�ȣ�", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    If CInt(txtWidth.Text) > UDWidth.Max Then
        MsgBox "��ӡ����ֽ�ſ�Ȳ��ܳ���" & UDWidth.Max & "���ף�", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    
    If Not IsNumeric(txtHeight.Text) Then
        MsgBox "��ȷ����ӡ����ֽ�Ÿ߶ȣ�", vbExclamation, App.Title
        txtWidth.SetFocus: Exit Function
    End If
    If CInt(txtHeight.Text) > UDHeight.Max Then
        MsgBox "��ӡ����ֽ�Ÿ߶Ȳ��ܳ���" & UDHeight.Max & "���ף�", vbExclamation, App.Title
        txtHeight.SetFocus: Exit Function
    End If
    
    '��������
    If Me.optPortrait.Value = True Then
        If Val(Me.txtUP.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�ϱ߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtDown.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�±߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtLeft.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "��߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtRight.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�ұ߾�̫��", vbExclamation, gstrSysName: Exit Function
    Else
        If Val(Me.txtUP.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�ϱ߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtDown.Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�±߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtLeft.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "��߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtRight.Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�ұ߾�̫��", vbExclamation, gstrSysName: Exit Function
    End If
    
    If optTabTiers(0).Value Then
        lngFixedRows = 1
    ElseIf optTabTiers(1).Value Then
        lngFixedRows = 2
    Else
        lngFixedRows = 3
    End If
    
    If Not PageHeadTest Then Exit Function
    
    '�Զ���ֽ��ʼ�����򱣴�߶ȺͿ��
    If mintPage = 256 Then
        Call SetCustonPager(Me.hWnd, mlngWidth, mlngHeight)
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    Else
        Printer.PaperSize = mintPage
        mlngWidth = Printer.Width
        mlngHeight = Printer.Height
    End If
    
    '�������ҳ�������ֹ����
    If Not OverRun Then Exit Function
    
    If Not GetRecordData(rsSaveData, True) Then Exit Function
    arrSQL = Array()
    With rsSaveData
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            strSQL = "Zl_�����ļ��ṹ_Update("
        '  Id_In         In ��ʱ��������.Id%Type,
            strSQL = strSQL & !ID & ","
        '  �ļ�id_In     In ��ʱ��������.�ļ�id%Type,
            strSQL = strSQL & !�ļ�ID & ","
        '  ��id_In       In ��ʱ��������.��id%Type,
            strSQL = strSQL & IIf(IsNull(!��ID), "NULL", !��ID) & ","
        '  �������_In   In ��ʱ��������.�������%Type,
            strSQL = strSQL & Val(!�������) & ","
        '  ��������_In   In ��ʱ��������.��������%Type,
            strSQL = strSQL & NVL(!��������, 4) & ","
        '  ������_In   In ��ʱ��������.������%Type,
            strSQL = strSQL & "NULL" & ","
        '  ��������_In   In ��ʱ��������.��������%Type,
            strSQL = strSQL & "NULL" & ",'"
        '  ��������_In   In ��ʱ��������.��������%Type,
            strSQL = strSQL & NVL(!��������) & "',"
        '  �����д�_In   In ��ʱ��������.�����д�%Type,
            strSQL = strSQL & NVL(!�����д�, "Null") & ",'"
        '  �����ı�_In   In ��ʱ��������.�����ı�%Type,
            strSQL = strSQL & NVL(!�����ı�) & "',"
        '  �Ƿ���_In   In ��ʱ��������.�Ƿ���%Type := 0,
            strSQL = strSQL & IIf(IsNull(!�Ƿ���), "NULL", !�Ƿ���) & ","
        '  Ԥ�����id_In In ��ʱ��������.Ԥ�����id%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  �������_In   In ��ʱ��������.�������%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  ʹ��ʱ��_In   In ��ʱ��������.ʹ��ʱ��%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  ����Ҫ��id_In In ��ʱ��������.����Ҫ��id%Type := Null,
            strSQL = strSQL & "NULL" & ","
        '  �滻��_In     In ��ʱ��������.�滻��%Type := 0,
            strSQL = strSQL & "NULL" & ",'"
        '  Ҫ������_In   In ��ʱ��������.Ҫ������%Type := Null,
            strSQL = strSQL & NVL(!Ҫ������) & "',"
        '  Ҫ������_In   In ��ʱ��������.Ҫ������%Type := Null,
            strSQL = strSQL & IIf(IsNull(!Ҫ������), "NULL", !Ҫ������) & ","
        '  Ҫ�س���_In   In ��ʱ��������.Ҫ�س���%Type := Null,
            strSQL = strSQL & IIf(IsNull(!Ҫ�س���), "NULL", !Ҫ�س���) & ","
        '  Ҫ��С��_In   In ��ʱ��������.Ҫ��С��%Type := Null,
            strSQL = strSQL & IIf(IsNull(!Ҫ��С��), "NULL", !Ҫ��С��) & ",'"
        '  Ҫ�ص�λ_In   In ��ʱ��������.Ҫ�ص�λ%Type := Null,
            strSQL = strSQL & NVL(!Ҫ�ص�λ) & "',"
        '  Ҫ�ر�ʾ_In   In ��ʱ��������.Ҫ�ر�ʾ%Type := 0,
            strSQL = strSQL & IIf(IsNull(!Ҫ�ر�ʾ), "NULL", !Ҫ�ر�ʾ) & ","
        '  ������̬_In   In ��ʱ��������.������̬%Type := 0,
            strSQL = strSQL & IIf(IsNull(!������̬), "NULL", !������̬) & ",'"
        '  Ҫ��ֵ��_In   In ��ʱ��������.Ҫ��ֵ��%Type := Null
            strSQL = strSQL & NVL(!Ҫ��ֵ��) & "')"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = strSQL
        .MoveNext
        Loop
    End With
    If rsSaveData.RecordCount > 0 Then
        strSQL = "Zl_�����ļ��ṹ_Commit(" & mlngFileID & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    End If
    blnTrans = False
    gcnOracle.BeginTrans
    blnTrans = True
    For i = 0 To UBound(arrSQL)
        If CStr(arrSQL(i)) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "Zl_�����ļ��ṹ_Update")
    Next
        
    strSQL = "Select l.����,l.���, l.����, f.��� As ҳ���, f.���� As ҳ����,f.����,f.ҳü,f.ҳ��" & _
        " From �����ļ��б� l, ����ҳ���ʽ f" & _
        " Where l.���� = f.����(+) And l.ҳ�� = f.���(+) And l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ļ���ӡ����", mlngFileID)
    If rsTemp.EOF Then
        MsgBox "���ļ��ڲ����ļ��б��в����ڡ����飡", vbInformation, gstrSysName
        Exit Function
    End If
    picFoot.Tag = NVL(rsTemp!����) & "-" & NVL(rsTemp!���)
    gcnOracle.CommitTrans
    blnTrans = False
    '��ҳü��ҳ�źͻ������Էֿ�����
    strPaper = mintPage & ";" & mintOrient & ";" & mlngHeight & ";" & mlngWidth & ";" & CLng(Me.ScaleY(mlngLeft, vbMillimeters, vbTwips)) & ";" & _
        CLng(Me.ScaleY(mlngRight, vbMillimeters, vbTwips)) & ";" & CLng(Me.ScaleY(mlngTop, vbMillimeters, vbTwips)) & ";" & _
        CLng(Me.ScaleY(mlngBottom, vbMillimeters, vbTwips))
    '�����ӡ����
    strSQL = "Zl_����ҳ���ʽ_Update(3" & ",'"
    '����_In In ����ҳ���ʽ.����%Type,
    '���_In In ����ҳ���ʽ.���%Type,
    strSQL = strSQL & NVL(rsTemp!���) & "','"
    '����_In In ����ҳ���ʽ.����%Type,
    strSQL = strSQL & NVL(rsTemp!����) & "','"
    '����_In In ����ҳ���ʽ.����%Type,
    strSQL = strSQL & NVL(rsTemp!����) & "','"
    '��ʽ_In In ����ҳ���ʽ.��ʽ%Type,
    strSQL = strSQL & strPaper & "','"
    'ҳü_In In ����ҳ���ʽ.ҳü%Type,
    strSQL = strSQL & NVL(rsTemp!ҳü) & "','"
    'ҳ��_In In ����ҳ���ʽ.ҳ��%Type
    strSQL = strSQL & NVL(rsTemp!ҳ��) & "')"
    
    gcnOracle.BeginTrans
    blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, "Zl_����ҳ���ʽ_Update")
    
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
'����:������ҳ���������֯�ɶ�Ӧ�����ļ��ṹ���ݼ�
'-----------------------------------------------------------------------
    Dim strSQL As String
    Dim rsSource As New ADODB.Recordset
    Dim lngParentId As Long, lngId As Long
    Dim lngRow As Long, lngRowNO As Long
    Dim intFields As Integer
    Dim lngItemNO As Long '��Ŀ���
    Dim intTitleNum As Integer
    Dim lngCol As Long
    Dim strData As String
    Dim strSubItem As String, strMidSub As String
    
    On Error GoTo errHand
    strSQL = "SELECT Id, �ļ�id, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���," & vbNewLine & _
        "       Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��" & vbNewLine & _
        " FROM �����ļ��ṹ" & vbNewLine & _
        " WHERE �ļ�id = 0"
    Call zlDatabase.OpenRecordset(rsSource, strSQL, "�����ļ��ṹ")
    '��ʼ���Ƽ�¼���ṹ��
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
            End If
        Next
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    '1:���µ��Ļ�����ʽ������
    lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 100)
    With rsTemp
        '������
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = Null
        .Fields("�������").Value = 1: .Fields("��������").Value = 1: .Fields("��������").Value = "���µ��Ļ�����ʽ������"
        .Fields("�����ı�").Value = "��ʽ����"
        .Update
        '�Ӷ���
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 101)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 1: .Fields("��������").Value = 4: .Fields("��������").Value = "�����ı�"
        .Fields("�����ı�").Value = txtTitleText.Text: .Fields("Ҫ������").Value = "�����ı�"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 102)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 2: .Fields("��������").Value = 4: .Fields("��������").Value = "��������"
        .Fields("�����ı�").Value = lblTitleFont.Caption: .Fields("Ҫ������").Value = "��������"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 103)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 3: .Fields("��������").Value = 4: .Fields("��������").Value = "����"
        .Fields("�����ı�").Value = Val(txtTabDays.Text): .Fields("Ҫ������").Value = "����"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 104)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 4: .Fields("��������").Value = 4: .Fields("��������").Value = "������"
        .Fields("�����ı�").Value = Val(txtTabDayTime.Text): .Fields("Ҫ������").Value = "������"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 105)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 5: .Fields("��������").Value = 4: .Fields("��������").Value = "��ʼʱ��"
        .Fields("�����ı�").Value = Val(txtTabBeginTime.Text): .Fields("Ҫ������").Value = "��ʼʱ��"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 106)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 6: .Fields("��������").Value = 4: .Fields("��������").Value = "ʱ����"
        .Fields("�����ı�").Value = Val(txtTabTimeSplit.Text): .Fields("Ҫ������").Value = "ʱ����"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 107)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 7: .Fields("��������").Value = 4: .Fields("��������").Value = "һ����Ŀ�����߶�"
        .Fields("�����ı�").Value = Val(txtTabRowHeight.Text): .Fields("Ҫ������").Value = "���߶�"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 108)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 8: .Fields("��������").Value = 4: .Fields("��������").Value = "һ����Ŀ����ͷ����"
        .Fields("�����ı�").Value = txtTabRowName(0).Text & "@" & txtTabRowName(1).Text & "@" & txtTabRowName(2).Text & "@" & txtTabRowName(3).Text
        .Fields("Ҫ������").Value = "��ͷ����"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 109)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 9: .Fields("��������").Value = 4: .Fields("��������").Value = "�̶������ܿ��(�)"
        .Fields("�����ı�").Value = Val(txtScaleColWidth.Text): .Fields("Ҫ������").Value = "�̶ȿ��"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 110)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 10: .Fields("��������").Value = 4: .Fields("��������").Value = "��ͼ�������߱���п�(�)"
        .Fields("�����ı�").Value = Val(txtCurveColWidth.Text): .Fields("Ҫ������").Value = "�����п�"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 111)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 11: .Fields("��������").Value = 4: .Fields("��������").Value = "��ͼ�������߱���и�(�)"
        .Fields("�����ı�").Value = Val(txtCurveRowHeight.Text): .Fields("Ҫ������").Value = "�����и�"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 112)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 12: .Fields("��������").Value = 4: .Fields("��������").Value = "���߱����ӿ�����(����Զ�������)"
        .Fields("�����ı�").Value = Val(txtAddCurveNull.Text) * 2: .Fields("Ҫ������").Value = "���߿���"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 113)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 13: .Fields("��������").Value = 4: .Fields("��������").Value = "������Ŀ�����߶�"
        .Fields("�����ı�").Value = Val(txtDownTabRowHeight.Text): .Fields("Ҫ������").Value = "���߶�1"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 114)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 14: .Fields("��������").Value = 4: .Fields("��������").Value = "������Ŀ�������ӵĿ�����"
        .Fields("�����ı�").Value = Val(txtAddNullTab.Text): .Fields("Ҫ������").Value = "������"
        .Update
        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 115)
        .AddNew
        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
        .Fields("�������").Value = 15: .Fields("��������").Value = 4: .Fields("��������").Value = "�Ƿ���Ӥ�����µ�"
        .Fields("�����ı�").Value = chkBaby.Value: .Fields("Ҫ������").Value = "Ӥ�����µ�"
        .Update
        
           'Ӥ�����µ�����
        'Ӥ�����µ�����
        If chkBaby.Value = 1 Then
            If Me.optTabTiers(0).Value Then
                intTitleNum = 1
            ElseIf Me.optTabTiers(1).Value Then
                intTitleNum = 2
            Else
                intTitleNum = 3
            End If
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 116)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 16: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ�����ͷ����"
            .Fields("�����ı�").Value = intTitleNum: .Fields("Ҫ������").Value = "��ͷ����"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 117)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 17: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ��������ı�"
            .Fields("�����ı�").Value = NVL(txtBabyTitleText.Text): .Fields("Ҫ������").Value = "Ӥ�������ı�"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 118)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 18: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ�����������"
            .Fields("�����ı�").Value = lblBabyTitleFont.Caption: .Fields("Ҫ������").Value = "Ӥ����������"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 119)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 19: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ�����ı�����"
            .Fields("�����ı�").Value = lblBabyFont.Caption: .Fields("Ҫ������").Value = "Ӥ���ı�����"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 120)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 20: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ��������ı���ɫ"
            .Fields("�����ı�").Value = lblTabTextColor.ForeColor: .Fields("Ҫ������").Value = "Ӥ���ı���ɫ"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 121)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 21: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ������ɫ"
            .Fields("�����ı�").Value = Me.shpTabGridColor.BorderColor: .Fields("Ҫ������").Value = "Ӥ�������ɫ"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 122)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 22: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ���������߶�"
            .Fields("�����ı�").Value = Val(txtBabyTabRowHeight.Text): .Fields("Ҫ������").Value = "Ӥ�����߶�"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 123)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 23: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ������߾�"
            .Fields("�����ı�").Value = Val(txtBabyLeft.Text): .Fields("Ҫ������").Value = "Ӥ�������߾�"
            .Update
            lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 124)
            .AddNew
            .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
            .Fields("�������").Value = 24: .Fields("��������").Value = 4: .Fields("��������").Value = "Ӥ�����µ����������"
            .Fields("�����ı�").Value = vfgThis.Cols - 1: .Fields("Ҫ������").Value = "������"
            .Update
            
        End If
    End With
    '2:���µ�������Ŀ����
    lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 200)
    With rsTemp
        '������
        .AddNew
        .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = Null
        .Fields("�������").Value = 2: .Fields("��������").Value = 1: .Fields("��������").Value = "���µ�������Ŀ����"
        .Fields("�����ı�").Value = "������Ŀ����"
        .Update
        lngRowNO = 1
        For lngRow = vfgCurve.FixedRows To vfgCurve.Rows - 1
            lngItemNO = Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("��Ŀ���")))
            If Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("ѡ��"))) <> 0 And lngItemNO <> 0 Then
                '�Ӷ���
                lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 200 + lngRowNO)
                .AddNew
                .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
                .Fields("�������").Value = lngRowNO: .Fields("��������").Value = 4
                .Fields("��������").Value = vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("��Ŀ����"))
                .Fields("�����ı�").Value = vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("��Ŀ���"))
                .Fields("Ҫ������").Value = vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("��Ŀ����"))
                .Update
                lngRowNO = lngRowNO + 1
                If lngItemNO = 2 And mbln����Ӧ�÷�ʽ = True Then
                    lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 200 + lngRowNO)
                    .AddNew
                    .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
                    .Fields("�������").Value = lngRowNO: .Fields("��������") = 4
                    .Fields("��������").Value = "����"
                    .Fields("�����ı�").Value = "-1"
                    .Fields("Ҫ������").Value = "����"
                    .Update
                    lngRowNO = lngRowNO + 1
                End If
            End If
        Next lngRow
    End With

    '2:���µ������Ŀ����
    lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 300)
    If chkBaby.Value = 0 Then
        With rsTemp
            '������
            .AddNew
            .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = Null
            .Fields("�������").Value = 3: .Fields("��������").Value = 1: .Fields("��������").Value = "���µ������Ŀ����"
            .Fields("�����ı�").Value = "�����Ŀ����"
            .Update
            lngRowNO = 1
            For lngRow = vfgTab.FixedRows To vfgTab.Rows - 1
                lngItemNO = Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���")))
                If lngItemNO <> 0 Then
                    '�Ӷ���
                    lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 300 + lngRowNO)
                    .AddNew
                    .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
                    .Fields("�������").Value = lngRowNO: .Fields("��������") = 4
                    .Fields("��������").Value = vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ����"))
                    .Fields("�����ı�").Value = vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���"))
                    .Fields("Ҫ������").Value = vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ����"))
                    .Fields("Ҫ�ر�ʾ").Value = Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")))
                    .Update
                    lngRowNO = lngRowNO + 1
                End If
            Next lngRow
        End With
    Else
        
        With rsTemp
            '������
            .AddNew
            .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = Null
            .Fields("�������").Value = 4: .Fields("��������").Value = 1: .Fields("��������").Value = "Ӥ�����µ�����ͷ��Ŀ"
            .Fields("�����ı�").Value = "Ӥ�����µ���ͷ��Ŀ"
            .Update
            lngRowNO = 1
            For lngRow = 2 To 4
                If vfgThis.RowHidden(lngRow) = False Then
                    For lngCol = vfgThis.FixedCols To vfgThis.Cols - 1
                        '�Ӷ���
                        lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 300 + lngRowNO)
                        .AddNew
                        .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
                        .Fields("�������").Value = lngCol: .Fields("��������") = 4
                        .Fields("�����ı�").Value = NVL(vfgThis.TextMatrix(lngRow, lngCol))                       '��ͷ����
'                        .Fields("Ҫ������").Value = NVL(vfgThis.TextMatrix(lngRow, lngCol))
'                        .Fields("Ҫ�ص�λ").Value = NVL(vfgThis.TextMatrix(lngRow, lngCol))
                        .Fields("�����д�").Value = lngRow - 1
                                                                                                                 '��¼Ƶ��
                        .Update
                        lngRowNO = lngRowNO + 1
                    
                    Next lngCol
                End If
                
            Next lngRow
            '�Ӷ���
           lngParentId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 400)
            '������
           .AddNew
           .Fields("ID").Value = lngParentId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = Null
           .Fields("�������").Value = 3: .Fields("��������").Value = 1: .Fields("��������").Value = "���µ������Ŀ����"
           .Fields("�����ı�").Value = "�����Ŀ����"
           .Update
           lngRowNO = 1
            For lngCol = vfgThis.FixedCols To vfgThis.Cols - 1
               
                lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 400 + lngRowNO)
                .AddNew
                .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
                .Fields("�������").Value = lngCol: .Fields("��������") = 4
                .Fields("��������").Value = vfgThis.ColWidth(lngCol) & "`" & vfgThis.Cell(flexcpAlignment, 5, lngCol, 5, lngCol)
                .Fields("�����ı�").Value = ""
                strData = NVL(vfgThis.TextMatrix(5, lngCol))
                
                If InStr(3, strData, "{") Then
                    strData = Mid(strData, 2)
                    strSubItem = Substr(strData, 1, (InStr(1, strData, "}") - 1) * 2)
                    strData = Mid(strData, InStr(1, strData, "}") + 1)
                    strMidSub = (Replace(strData, Mid(strData, InStr(1, strData, "{")), ""))
                    strData = Mid(strData, InStr(1, strData, "{"))
                    .Fields("�����д�").Value = 1
                    .Fields("Ҫ������").Value = strSubItem
                    .Fields("Ҫ�ص�λ").Value = strMidSub
                    .Fields("Ҫ�ر�ʾ").Value = Val(Split(Split(vfgThis.Cell(flexcpData, 5, lngCol, 5), " ")(0), "`")(1))
                    .Update
                    lngRowNO = lngRowNO + 1
                    lngId = IIf(blnSave = True, zlDatabase.GetNextId("�����ļ��ṹ"), 400 + lngRowNO)
                    .AddNew
                    .Fields("ID").Value = lngId: .Fields("�ļ�ID").Value = mlngFileID: .Fields("��ID").Value = lngParentId
                    .Fields("�������").Value = lngCol: .Fields("��������") = 4
                    .Fields("��������").Value = vfgThis.ColWidth(lngCol) & "`" & vfgThis.Cell(flexcpAlignment, 5, lngCol, 5, lngCol)
                    .Fields("�����ı�").Value = ""
                    strData = Mid(strData, 2)
                    strSubItem = Substr(strData, 1, (InStr(1, strData, "}") - 1) * 2)
                    .Fields("�����д�").Value = 2
                    .Fields("Ҫ������").Value = strSubItem
                    .Fields("Ҫ�ص�λ").Value = ""
                    If InStr(vfgThis.Cell(flexcpData, 5, lngCol, 5), " ") > 0 Then
                        .Fields("Ҫ�ر�ʾ").Value = Val(Split(Split(vfgThis.Cell(flexcpData, 5, lngCol, 5), " ")(1), "`")(1))
                    Else
                        .Fields("Ҫ�ر�ʾ").Value = 1
                    End If
                    .Update
                    lngRowNO = lngRowNO + 1
                    
                Else
                    strData = Mid(Replace(strData, " ", ""), 2)
                    strSubItem = Replace(strData, "}", "")
                    .Fields("�����д�").Value = 1
                    .Fields("Ҫ������").Value = strSubItem
                    .Fields("Ҫ�ص�λ").Value = ""
                    .Fields("Ҫ�ر�ʾ").Value = 0
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
    Dim lngBottom As Long  '�ͻ�����Ĵ�С
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
    Call zlFontSet("��������", strCurFont)
    If strCurFont = Me.lblBabyTitleFont.Caption Then Exit Sub
    Me.lblBabyTitleFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
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
    Dim blnSplit As Boolean                         '�����Ŀʱ���,���ǰ�����Ŀ�޺�׺�Һ�һ����Ŀ��ǰ׺,��blnSplit=False,���������
    Dim intType As Integer, intFace As Integer, intLen As Integer      '��Ŀ����
    Dim strFaces As String                          '������Ŀ,ֻ����¼����Ŀ0�뵥ѡ��Ŀ4;������������Ŀ,ֻ����¼����Ŀ0
    Dim strName As String                           '��Ŀ����
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
            '��һ�а�2����Ŀʱ����Ŀ֮��������ǰ׺/���׺���ż�������
            '��һ�а󶨶����Ŀʱ����Ŀ���ͱ�����¼������Ŀ
            '��ѡ���ѡ��Ŀ������������Ŀһ�����ĳ��
            'ϵͳ�̶�����Ŀ����ǩ���ˣ����ڣ�ʱ��ȣ�һ��ֻ�ܰ�һ��
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
            
            '��ǰ��
            For lngCount = 0 To .ListCount - 1
                strName = Mid(.List(lngCount), InStr(1, .List(lngCount), "{"))
                strName = Mid(strName, 1, InStr(1, strName, "}"))
                strTmp1 = strTmp1 & "'" & strName
            Next lngCount
            
            '������
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
                
                If strName = "����ѹ" Then
                    If lngCount <> 0 Then
                        MsgBox "����ѹ������ѹ�󶨸�ʽӦΪ:����ѹ/����ѹ", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    
                End If
                'ÿһ�а󶨵���Ŀ�����ظ�
                If .ListCount > 1 Then
                    If UBound(Split(strTmp1, "'{" & strName & "}")) > 1 Then
                        MsgBox "��һ�а󶨶����Ŀʱ����Ŀ���Ʋ����ظ���", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                
                
                '�����е���Ŀ�Ƿ�����������Ŀ�ظ�
                For intCount = 0 To UBound(arrCol)
                    If CStr(arrColValue(intCount)) = "{" & strName & "}" And Trim(strName) <> "" Then
                        MsgBox "{" & strName & "}�Ѿ��ڵ�" & CInt(arrCol(intCount)) & "����,���飡", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next intCount
                
                
                If lngCount > 0 Then
                    '�����Ŀ���Ƿ���ڷָ���
                    If Not blnSplit Then
                        If Trim(Split(.List(lngCount), "{")(0) = "") Then
                            MsgBox "��һ�а󶨶����Ŀʱ����Ŀ֮�����Ҫ����ǰ׺���׺���ż������֣�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    '�����Ŀ�����Ƿ�һ��
                    mrsItems.Filter = "��Ŀ����='" & strName & "'"
                    If mrsItems.RecordCount <> 0 Then
                        If Not (intType = mrsItems!��Ŀ���� And intFace = mrsItems!��Ŀ��ʾ) Then
                            MsgBox "��һ�а󶨶����Ŀʱ����Ŀ�����ͱ���һ�£�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        If mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ��ʾ = 0 And NVL(mrsItems!��Ŀ����, 1) > 3 Then
                            MsgBox "һ�����ֻ�ܰ���������С�ڻ����3���ı���Ŀ��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                    Else
                        If InStr(1, "����,ʱ��", strName) = 0 Then
                            MsgBox "�̶���Ŀ������������Ŀ����һ��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    'ֻ��Ҫȡ��һ����Ŀ�����Լ���
                    mrsItems.Filter = "��Ŀ����='" & strName & "'"
                    If mrsItems.RecordCount <> 0 Then
                        intType = mrsItems!��Ŀ����
                        intFace = mrsItems!��Ŀ��ʾ
                        intLen = NVL(mrsItems!��Ŀ����, 1)
                        If .ListCount > 1 Then
                            If intType = 1 And intFace = 0 And intLen > 3 Then
                                MsgBox "һ�����ֻ�ܰ���������С�ڻ����3���ı���Ŀ��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If intFace = 3 Then
                                MsgBox "��ѡ��ֻ�ܵ����󶨣�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        'һ��������������������<=3���ı���Ŀ
                        If .ListCount > 2 Then
                            If intType = 1 And intFace = 0 Then
                                MsgBox "һ�����ֻ�ܰ���������С�ڻ����3���ı���Ŀ��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    Else
                        If .ListCount > 1 Then
                            If InStr(1, "����,ʱ��", strName) = 0 Then
                                MsgBox "�̶���Ŀ������������Ŀ����һ��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                '�̶�ȡ��һ����Ŀ�ĺ�׺
                blnSplit = Trim(Split(.List(lngCount), "}")(1) <> "")
            Next
            strTemp = Trim(strTemp)
            strTmp = Replace(Trim(strTmp), " ", "")
            mrsItems.Filter = 0
            
            With vfgThis
                .TextMatrix(5, Me.udColumnNo.Value) = strTmp
                '���ݶ��뷽ʽ����������
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
        .DialogTitle = "��־ͼѡ��"
        .Filename = ""
        .Filter = "ͼ��(*.jpg;*.bmp;*.ico;*.gif)|*.jpg;*.bmp;*.ico;*.gif"
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
    If picTemp Is Nothing Then MsgBox "������Ч��ͼƬ�ļ���", vbExclamation, Me.Caption: Exit Sub
    
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
    Call zlFontSet("�ı�����", strCurFont)
    If strCurFont = Me.lblBabyFont.Caption Then Exit Sub
    Me.lblBabyFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
         .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
    End With
    Set Me.vfgThis.Font = objFont
    DataChanged = True
End Sub

Private Sub cmdTabGridColor_Click()
    Dim lngCurColor As Long
    lngCurColor = Me.shpTabGridColor.BorderColor
    Call zlColorSet("�����ɫ", lngCurColor)
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
    Call zlColorSet("�ı���ɫ", lngCurColor)
    If lngCurColor = Me.lblTabTextColor.ForeColor Then Exit Sub
    Me.lblTabTextColor.ForeColor = lngCurColor
    Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
    DataChanged = True
End Sub

Private Sub cmdTitleFont_Click()
    Dim strCurFont As String
    strCurFont = Me.lblTitleFont.Caption
    Call zlFontSet("��������", strCurFont)
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
    '����ǰ��ʽӦ�õ����л����ļ�
    
    gstrSQL = " Select ����||'-'||��� AS KEY From �����ļ��б� Where ����=3 and ID<>[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�", mlngFileID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "��ǰֻ��һ�ݻ����ļ�������Ҫִ��ͬ�����ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("���ٴ�ȷ�ϣ�" & vbCrLf & "        ִ�иù��ܺ����л����ļ���ҳüҳ�Ÿ�ʽ��ͳһ�뵱ǰ�ļ����ñ���һ�£�", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '��ȡ��ǰ���õ�ҳüҳ��
    strZIPHead = ReadPageHeadFile(picFoot.Tag)
    strZIPFoot = ReadPageFootFile(picFoot.Tag)
    
    gcnOracle.BeginTrans
    blnTrans = True
    'ѭ��д�����ݿ�
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
    
    'ɾ����ʱ�ļ�
    gobjFSO.DeleteFile strZIPHead, True
    gobjFSO.DeleteFile strZIPFoot, True
    
    MsgBox "ͬ���ɹ���", vbInformation, gstrSysName
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
        MsgBox "ϵͳ��û�а�װ�κδ�ӡ��,���Ȱ�װ��ӡ����", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    Call RestoreWinState(Me, App.ProductName)

    With Me.tbcStyle
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .InsertItem 0, "��������", Me.picTable.hWnd, 0
        .InsertItem 1, "��ӡ����", Me.picOutput.hWnd, 0
        .Item(0).Selected = True
    End With
    Call InitMenuBar  '���ز˵�
    
    Dim objPane As Pane
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis

    Set objPane = dkpMan.CreatePane(1, 100, 100, DockLeftOf, Nothing): objPane.Title = "���": objPane.Options = PaneNoCaption
    Set objPane = dkpMan.CreatePane(2, 100, 200, DockRightOf, objPane): objPane.Title = "��ʽ": objPane.Options = PaneNoCaption
    
    With Me.vfgThis
        .MergeCellsFixed = flexMergeFree
        .Rows = 6
        .ROWHEIGHT(0) = 300
        .ROWHEIGHT(1) = 300
        .TextMatrix(1, 0) = "�����ı�"
        .TextMatrix(2, 0) = "��ͷ��Ԫ"
        .TextMatrix(3, 0) = "��ͷ��Ԫ"
        .TextMatrix(4, 0) = "��ͷ��Ԫ"
        .TextMatrix(5, 0) = "��������"
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
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
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
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
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
    'ˢ��������Ϣ
    gblnOK = False
    mbln����Ӧ�÷�ʽ = False
    '------------------------------------------------------------------
    '��ʼ����������ҳ��
    '------------------------------------------------------------------
    With vfgCurve
        strHead = "���,500,4,1;ѡ��,500,4,1;��Ŀ���,0,4,1;��Ŀ����,1200,1,1;��Ŀ��λ,800,1,1"
        Call SetVsFlexGridChangeHead(strHead, vfgCurve, 1)
        .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
        .FrozenCols = .ColIndex("ѡ��")
        .SheetBorder = &H40C0&
    End With
                    
    With Me.lvwTabItem.ColumnHeaders
        .Clear
        .Add , "_����", "��Ŀ����", 2050
        .Add , "_���", "��Ŀ���", 0
        .Add , "_��λ", "��Ŀ��λ", 0
        .Add , "_��ʾ", "��Ŀ��ʾ", 0
        Me.lvwTabItem.ListItems.Clear
    End With
    
    With vfgThis
        .Cols = 3
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vfgTab
        strHead = "���,500,4,1;��Ŀ���,0,4,1;��Ŀ����,1200,1,1;��Ŀ��λ,800,1,1;��Ŀ��ʾ,0,1,1;��¼Ƶ��,800,4,1"
        Call SetVsFlexGridChangeHead(strHead, vfgTab, 1)
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDKbdMouse
    End With
    
    gstrSQL = "Select l.���, l.����, l.˵�� From �����ļ��б� l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    Me.Caption = "���µ���ʽ - " & rsTemp!����
    strTitle = rsTemp!����
    Me.txtTitleText.Text = strTitle
    
    gstrSQL = "Select 1 From �����¼��Ŀ where ��Ŀ���=[1] And Ӧ�÷�ʽ=2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, -1)
    If rsTemp.RecordCount > 0 Then mbln����Ӧ�÷�ʽ = True
    
    gstrSQL = " SELECT DECODE(A.��Ŀ���,-1,2,A.��Ŀ���) �������,A.��Ŀ���,A.��Ŀ���� as ��Ŀ��, B.��¼�� ��Ŀ����, A.��Ŀ��λ, B.��¼��,DECODE(NVL(C.��Ŀ���,''),'',A.��Ŀ��ʾ,4) ��Ŀ��ʾ" & vbNewLine & _
        " FROM �����¼��Ŀ A, ���¼�¼��Ŀ B,��������Ŀ C" & vbNewLine & _
        " WHERE A.��Ŀ��� = B.��Ŀ��� And A.��Ŀ���=C.��Ŀ���(+) AND NOT (NVL(A.Ӧ�÷�ʽ,0)=2 And A.��Ŀ���=-1) And A.��Ŀ����=1 " & vbNewLine & _
        " ORDER BY ��Ŀ���"
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    rsItems.Filter = "��¼��=1 Or ��¼��=3"
    rsItems.Sort = "�������,��Ŀ���"
    With rsItems
        Do While Not .EOF
            If .AbsolutePosition > vfgCurve.Rows - 1 Then vfgCurve.Rows = .AbsolutePosition + 1
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("���")) = .AbsolutePosition
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("��Ŀ���")) = Val(!��Ŀ���)
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("��Ŀ����")) = NVL(!��Ŀ����)
            vfgCurve.TextMatrix(.AbsolutePosition, vfgCurve.ColIndex("��Ŀ��λ")) = NVL(!��Ŀ��λ)
        .MoveNext
        Loop
    End With
    
    rsItems.Filter = "��¼��=2"
    rsItems.Sort = "��Ŀ���"
    With rsItems
        Do While Not .EOF
            If !��Ŀ��� = 4 Then
                Set objItem = Me.lvwTabItem.ListItems.Add(, "_" & !��Ŀ���, "Ѫѹ")
                objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_���").Index - 1) = "4,5"
            Else
                Set objItem = Me.lvwTabItem.ListItems.Add(, "_" & !��Ŀ���, NVL(!��Ŀ����))
                objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_���").Index - 1) = !��Ŀ���
            End If
            objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_��λ").Index - 1) = NVL(!��Ŀ��λ)
            objItem.SubItems(Me.lvwTabItem.ColumnHeaders("_��ʾ").Index - 1) = NVL(!��Ŀ��ʾ)
            
        .MoveNext
        Loop
    End With
    


    gstrSQL = "Select b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��Ŀ��ʾ, b.��Ŀ����" & vbNewLine & _
              "      From ���¼�¼��Ŀ A, �����¼��Ŀ B " & vbNewLine & _
              "      Where a.��Ŀ��� = b.��Ŀ��� And ��¼�� = 2 " & vbNewLine & _
              "      Order By ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    rsItems.Filter = "��¼��=2"
    rsItems.Sort = "��Ŀ���"
    With rsItems
        Me.lstColumnItems.Clear
        Me.lstColumnItems.AddItem "����"
        Me.lstColumnItems.AddItem "ʱ��"
        Me.lstColumnItems.AddItem "��������"
        Do While Not .EOF
            Me.lstColumnItems.AddItem "" & !��Ŀ��
            .MoveNext
        Loop
        Me.lstColumnItems.AddItem "��ʿ"
        Me.lstColumnItems.ListIndex = 0
        .MoveFirst
    End With
    
    gstrSQL = "Select ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ��ʾ,��Ŀ���� From �����¼��Ŀ Order By ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    '---------------------------------------------------
    '������ʽ��ȡ
    '---------------------------------------------------
    
    Me.optTabTiers(0).Value = True
    Call optTabTiers_Click(0)
    txtTabBeginTime.Tag = 4: txtTabTimeSplit.Tag = 4
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ʽ����'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "�����ı�"
                Me.txtTitleText.Text = "" & !�����ı�
            Case "��������"
                Me.lblTitleFont.Caption = "" & !�����ı�
            Case "���߶�"
                Me.txtTabRowHeight.Text = Val("" & !�����ı�)
                If Me.txtTabRowHeight.Text < 225 Or Me.txtTabRowHeight.Text > 600 Then
                    Me.txtTabRowHeight.Text = 225
                End If
            Case "��ͷ����"
                If Val("" & !�����ı�) = 1 Then
                    Me.optTabTiers(0).Value = True
                    Call optTabTiers_Click(0)
                ElseIf Val("" & !�����ı�) = 2 Then
                    Me.optTabTiers(1).Value = True
                    Call optTabTiers_Click(1)
                Else
                    Me.optTabTiers(2).Value = True
                    Call optTabTiers_Click(2)
                End If
            Case "Ӥ�������ı�"
                Me.txtBabyTitleText.Text = "" & !�����ı�
            Case "Ӥ���ı�����"
                Me.lblBabyTitleFont.Caption = "" & !�����ı�
            Case "Ӥ�����߶�"
                Me.txtBabyTabRowHeight = Val("" & !�����ı�)
                If Me.txtBabyTabRowHeight.Text < 225 Or Me.txtBabyTabRowHeight.Text > 600 Then
                    Me.txtBabyTabRowHeight.Text = 225
                End If
            Case "�ı���ɫ"
                Me.lblTabTextColor.ForeColor = Val("" & !�����ı�)
                Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
            Case "�����ɫ"
                Me.shpTabGridColor.BorderColor = Val("" & !�����ı�)
                With Me.vfgThis
                    .GridColor = Me.shpTabGridColor.BorderColor
                    .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
                    .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
                End With
            Case "����"
                Me.txtTabDays.Text = Val("" & !�����ı�)
                If Me.txtTabDays.Text = "" Then Me.txtTabDays.Text = 7
            Case "������":  Me.udBabyCol.Value = Val("" & !�����ı�)
            Case "������"
                If InStr(1, ",2,4,6,8,12,24,", "," & Val("" & !�����ı�) & ",") = 0 Then
                    Me.txtTabDayTime.Text = 6
                Else
                    Me.txtTabDayTime.Text = Val("" & !�����ı�)
                End If
            Case "��ʼʱ��"
                Me.txtTabBeginTime.Text = Val("" & !�����ı�)
                txtTabBeginTime.Tag = txtTabBeginTime.Text
            Case "ʱ����"
                Me.txtTabTimeSplit.Text = Val("" & !�����ı�)
                txtTabTimeSplit.Tag = txtTabTimeSplit.Text
            Case "��ͷ����"
                strTitle = "" & !�����ı�
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
                If txtTabRowName(0).Text = "" Then txtTabRowName(0).Text = "��    ��"
                If txtTabRowName(1).Text = "" Then txtTabRowName(1).Text = "סԺ����"
                If txtTabRowName(2).Text = "" Then txtTabRowName(2).Text = "����������"
                If txtTabRowName(3).Text = "" Then txtTabRowName(3).Text = "ʱ    ��"
            Case "�̶ȿ��"
                txtScaleColWidth.Text = Val("" & !�����ı�)
            Case "�����п�"
                txtCurveColWidth.Text = Val("" & !�����ı�)
            Case "�����и�"
                txtCurveRowHeight.Text = Val("" & !�����ı�)
            Case "���߿���"
                txtAddCurveNull.Text = Val("" & !�����ı�) \ 2
                If Val(txtAddCurveNull.Text) < 0 Then txtAddCurveNull.Text = 0
            Case "���߶�1"
                Me.txtDownTabRowHeight.Text = Val("" & !�����ı�)
                If Val(Me.txtDownTabRowHeight.Text) < 225 Or Val(Me.txtDownTabRowHeight.Text) > 600 Then
                    Me.txtDownTabRowHeight.Text = 225
                End If
            Case "������"
                txtAddNullTab.Text = Val("" & !�����ı�)
                If Val(txtAddNullTab.Text) < 0 Then txtAddNullTab.Text = 0
                Case "Ӥ�����µ�"
                chkBaby.Value = Val("" & !�����ı�)
            Case "��ͷ����"
                If Val("" & !�����ı�) - 1 >= 0 Then
                    Me.optTabTiers(Val("" & !�����ı�) - 1).Value = 1
                End If
            Case "Ӥ�������ı�"
                txtBabyTitleText.Text = "" & !�����ı�
            Case "Ӥ����������"
                lblBabyTitleFont.Caption = "" & !�����ı�
            Case "Ӥ���ı���ɫ"
                lblTabTextColor.ForeColor = Val("" & !�����ı�)
            Case "Ӥ���ı�����"
                lblBabyFont.Caption = "" & !�����ı�
            Case "Ӥ�������ɫ"
                txtBabyTabRowHeight.Text = Val("" & !�����ı�)
            Case "Ӥ�����߶�"
                txtBabyTabRowHeight.Text = Val("" & !�����ı�)
            Case "Ӥ�������߾�"
                txtBabyLeft.Text = Val("" & !�����ı�)
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
    '--������Ŀ����
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '������Ŀ����'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            For lngRow = vfgCurve.FixedRows To vfgCurve.Rows - 1
                If Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("��Ŀ���"))) = Val(NVL(!�����ı�)) Or _
                    Val(vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("��Ŀ���"))) = 1 Then
                    vfgCurve.TextMatrix(lngRow, vfgCurve.ColIndex("ѡ��")) = 1
                End If
            Next
        .MoveNext
        Loop
    End With
    
    
    '--�����Ŀ����
    If chkBaby.Value = 0 Then
        gstrSQL = "Select d.�������, d.�����ı�,d.Ҫ������,d.Ҫ�ر�ʾ " & _
            " From �����ļ��ṹ d, �����ļ��ṹ p" & _
            " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����Ŀ����'" & _
            " Order By d.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            lngRow = 1
            Do While Not .EOF
                For lngIndex = 1 To lvwTabItem.ListItems.Count
                    If lvwTabItem.ListItems(lngIndex).SubItems(1) = NVL(!�����ı�) Then
                        lvwTabItem.ListItems(lngIndex).Checked = True
                        If lngRow > vfgTab.Rows - 1 Then vfgTab.Rows = vfgTab.Rows + 1
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("���")) = lngRow
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���")) = NVL(!�����ı�)
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ����")) = lvwTabItem.ListItems(lngIndex).Text
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ��λ")) = lvwTabItem.ListItems(lngIndex).SubItems(2)
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ��ʾ")) = lvwTabItem.ListItems(lngIndex).SubItems(3)
                        If InStr(1, GetTabFrequency(Val(txtTabDayTime.Text), Val(lvwTabItem.ListItems(lngIndex).SubItems(3))), Val(NVL(!Ҫ�ر�ʾ))) = 0 Then
                            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
                        Else
                            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = NVL(!Ҫ�ر�ʾ)
                        End If
                        If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���"))) = 3 Then
                            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = Val(txtTabDayTime.Text)
                        End If
                        lngRow = lngRow + 1
                        Exit For
                    End If
                Next
            .MoveNext
            Loop
        End With
    Else
        gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = 'Ӥ�����µ���ͷ��Ŀ'" & _
        " Order By d.�������"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        With rsTemp
            Do While Not .EOF
                If Val(!�������) > vfgThis.Cols - 1 Then vfgThis.Cols = vfgThis.Cols + 1
                Me.vfgThis.TextMatrix(!�����д� + 1, !�������) = "" & !�����ı�
                .MoveNext
            Loop
        End With
        Me.udColumnNo.Max = vfgThis.Cols - 1
        '------------------------------------------------------------------------------------------------------------------
        gstrSQL = "Select d.�������,d.������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ,d.Ҫ��ֵ�� " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����Ŀ����'" & _
        " Order By d.�������, d.�����д�"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
        
        vfgThis.Cell(flexcpData, 5, vfgThis.FixedCols, 5, vfgThis.Cols - 1) = ""
        With rsTemp
            Me.lstColumnUsed.Clear
            Do While Not .EOF
                Me.vfgThis.ColWidth(!�������) = Val(Split("" & !��������, "`")(0))
                If InStr(1, "" & !��������, "`") <> 0 Then
                    vfgThis.Cell(flexcpAlignment, 5, !�������, 5, !�������) = Val(Split("" & !��������, "`")(1))
                Else
                    vfgThis.Cell(flexcpAlignment, 5, !�������, 5, !�������) = flexAlignLeftCenter
                End If
                If Me.udColumnNo.Value <> !������� Then Me.udColumnNo.Value = !�������
                Me.lstColumnUsed.AddItem !�����ı� & "{" & !Ҫ������ & "}" & !Ҫ�ص�λ
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = zlCommFun.NVL(!Ҫ�ر�ʾ, 0)
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
    
    '�Ӳ���ҳ���ʽ����ȡ��ӡ��������
    strSQL = "Select l.����,l.���,f.��� As ҳ���, f.��ʽ" & _
        " From �����ļ��б� l, ����ҳ���ʽ f" & _
        " Where l.���� = f.����(+) And l.ҳ�� = f.���(+) And l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ļ���ӡ����", mlngFileID)
    If Not rsTemp.EOF Then
        picFoot.Tag = NVL(rsTemp!����) & "-" & NVL(rsTemp!ҳ���, rsTemp!���)
        strPaper = "" & rsTemp!��ʽ
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
       If sFont = "����" Then cboFont.ListIndex = i
    Next i
    With cboFSize
        .AddItem "����"
        .AddItem "С��"
        .AddItem "һ��"
        .AddItem "Сһ"
        .AddItem "����"
        .AddItem "С��"
        .AddItem "����"
        .AddItem "С��"
        .AddItem "�ĺ�"
        .AddItem "С��"
        .AddItem "���"
        .AddItem "С��"
        .AddItem "����"
        .AddItem "С��"
        .AddItem "�ߺ�"
        .AddItem "�˺�"
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
    '���ܣ�
    '������
    '���أ�
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
    cbsThis.ActiveMenuBar.Title = "�˵���"
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
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '�����
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, "���沢�˳�"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����"): cbrControl.ToolTipText = "�����Ѹ��ĵ�����(Ctrl+S,F2)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "�ָ�"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�ָ����ϴα���ʱ������״̬"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "���е�ǰ���µ�Ԥ��"
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): cbrControl.ToolTipText = "�˳���ǰ����ƴ���(Esc)"

    End With
        
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
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
        Cancel = (MsgBox("���ĺ����Ʊ��뱣������Ч���Ƿ�������棿", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
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
    '��֮ǰ�ı����Ŀ��Ϣ��������
    With vfgTab
        For lngRow = .FixedRows To .Rows - 1
            ReDim Preserve arrTab(0 To i)
            arrTab(UBound(arrTab)).ItemNO = .TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���"))
            arrTab(UBound(arrTab)).ItemName = .TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ����"))
            arrTab(UBound(arrTab)).ItemUnit = .TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ��λ"))
            arrTab(UBound(arrTab)).ItemShow = Val(.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ��ʾ")))
            arrTab(UBound(arrTab)).ItemFrequency = .TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��"))
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
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("���")) = lngRow
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���")) = lvwTabItem.ListItems(lngIndex).SubItems(1)
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ����")) = lvwTabItem.ListItems(lngIndex).Text
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ��λ")) = lvwTabItem.ListItems(lngIndex).SubItems(2)
            vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ��ʾ")) = lvwTabItem.ListItems(lngIndex).SubItems(3)
            If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���"))) = 3 Then
                vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = Val(txtTabDayTime.Text)
            Else
                If blnTrue = True Then
                    If InStr(1, GetTabFrequency(Val(txtTabDayTime.Text), Val(lvwTabItem.ListItems(lngIndex).SubItems(3))), Val(arrTab(i).ItemFrequency)) = 0 Then
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
                    Else
                        vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = arrTab(i).ItemFrequency
                    End If
                Else
                    vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
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
    Case "��������"
        If txtTitleText.Enabled And txtTitleText.Visible Then txtTitleText.SetFocus
    Case "��ӡ����"
        If cboPage.Enabled And cboPage.Visible Then cboPage.SetFocus
    End Select
End Sub

Private Sub InitPageFoot()
    'ҳüҳ������
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
        Set cbrPopupBar = cbsThis.Add("�Ҽ��˵�", xtpBarPopup)
        cbrPopupBar.Title = "�Ҽ��˵�"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddLeft, "�����������"): cbrPopupItem.IconId = conMenu__Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddRight, "���Ҳ�������"): cbrPopupItem.IconId = conMenu__Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteCol, "ɾ����"): cbrPopupItem.IconId = conMenu__Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddUP, "���Ϸ�������"): cbrPopupItem.IconId = conMenu__Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddBottom, "���·�������"): cbrPopupItem.IconId = conMenu__Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteRow, "ɾ����"): cbrPopupItem.IconId = conMenu__Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddySingle, "�󶨵���"): cbrPopupItem.IconId = conMenu__Curve_BuddySingle
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddyDouble, "��˫��"): cbrPopupItem.IconId = conMenu__Curve_BuddyDouble
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
        Set cbrPopupBar = cbsThis.Add("�Ҽ��˵�", xtpBarPopup)
        cbrPopupBar.Title = "�Ҽ��˵�"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddLeft, "�����������"): cbrPopupItem.IconId = conMenu__Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddRight, "���Ҳ�������"): cbrPopupItem.IconId = conMenu__Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteCol, "ɾ����"): cbrPopupItem.IconId = conMenu__Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddUP, "���Ϸ�������"): cbrPopupItem.IconId = conMenu__Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddBottom, "���·�������"): cbrPopupItem.IconId = conMenu__Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteRow, "ɾ����"): cbrPopupItem.IconId = conMenu__Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddySingle, "�󶨵���"): cbrPopupItem.IconId = conMenu__Curve_BuddySingle
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddyDouble, "��˫��"): cbrPopupItem.IconId = conMenu__Curve_BuddyDouble
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
    '���,��������ڵ��ĸ���Ԫ���ֵ��ͬ,����������(�п�����Ҫ������,����,����,���½��м��)
    lngRow = udHeadRow.Value + 1
    lngCol = Me.udHeadCol.Value
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 1, lngCol) = strInput
    
    If lngRow <= 4 Then
        If (vfgThis.TextMatrix(3, lngCol) = vfgThis.TextMatrix(4, lngCol) And vfgThis.TextMatrix(3, lngCol) <> "") Then
            If lngCol > 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '����
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
                '����
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '����
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
    'ҳüҳ������
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
    '����һ��ļ�����������ʼʱ���ʱ���������ֵ
    Dim intHour As Integer
    Dim lngRow As Long, lngFrequency As Long
    Dim blnSetTab As Boolean
    
    '68649:������,2013-12-19,��ʼʱ���ʱ�������ֵ������
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
    
    '�����ĿƵ�μ��
    For lngRow = vfgTab.FixedRows To vfgTab.Rows - 1
        If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���"))) <> 0 Then
            lngFrequency = Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")))
            If Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ���"))) = 3 Then
                vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = Val(txtTabDayTime.Text)
            Else
                If InStr(1, GetTabFrequency(Val(txtTabDayTime.Text), Val(vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��Ŀ��ʾ")))), lngFrequency) = 0 Then
                    vfgTab.TextMatrix(lngRow, vfgTab.ColIndex("��¼Ƶ��")) = IIf(Val(txtTabDayTime.Text) > 2, 2, 1)
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
    
    'ҳüҳ������
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
    '68649:������,2013-12-19,��ʼʱ���ʱ�������ֵ������
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
'���ܣ���ʾ���õ�ֽ�ŵ�Ԥ��
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
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Private Function ReadPageFoot(objFoot As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

Private Function ReadPageHeadFile(ByVal StrKey As String) As String
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageHeadFile = strZip
    End If
End Function

Private Function ReadPageFootFile(ByVal StrKey As String) As String
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageFootFile = strZip
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
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
    'blnBuild=False:�����ļ���ѹ��;True:�Ѳ���ѹ���ļ�
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
    'blnBuild=False:�����ļ���ѹ��;True:�Ѳ���ѹ���ļ�
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
    '�����ϱ߾෵�ؼ�
    Dim fr As FORMATRANGE           '��ʽ�����ı���Χ
    Dim rcDrawTo As RECT            'Ŀ����������
    Dim rcPage As RECT              'Ŀ��ҳ������
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
        
        lngPageCount = lngPageCount + 1             ' ҳ����1
        '��¼��ҳ��Ϣ
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          'ʵ�ʴ�ӡ�߶�
        AllPages(lngPageCount).Start = lngTMP
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTMP Or lngNextPos >= lngLen Then Exit Do      ' �������ҳ��ķ�ҳ
        lngTMP = lngNextPos
    Loop
    Call SendMessage(rtbHead.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    If fr.rc.Bottom > rcDrawTo.Bottom Or lngPageCount > 1 Then
        MsgBox "��Ƶ�ҳü���ݳ����ϱ߾࣡", vbInformation, gstrSysName
        Exit Function
    End If
    PageHeadTest = True
End Function

Private Function OverRun() As Boolean
    Dim intPageMargin As Integer    '�߾�
    Dim lngPageWidth As Long, lngPageHeight As Long      'ֽ�ſ�ȡ��߶�
    Dim lngTrimWidth  As Long, lngTrimHeight As Long       '����ʵ��ռ�õĿ�Ⱥ͸߶�(���߾�)
    Dim lngTabWidth As Long, lngTabHeight As Long          'Ӥ�����µ����
    Dim i As Integer
    
    '������µ��Ŀ���Ƿ񳬳�ҳ����Ч��ӡ��Χ
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
        MsgBox "Ӥ�����µ���������ʵ�ʿ�ȴ�����ֽ�ŵĿ��,�����!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngTrimWidth > lngPageWidth - 100 Then
        MsgBox "���µ������ʵ�ʿ�ȴ�����ֽ�ŵĿ��,�����!", vbInformation, gstrSysName
        Exit Function
    End If
    
    If lngTrimHeight > lngPageHeight - 100 Then
        MsgBox "���µ������ʵ�ʸ߶ȴ�����ֽ�ŵĸ߶�,�����!", vbInformation, gstrSysName
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
        If InStr(1, strFont, "��") > 0 Then
            .FontBold = True
        Else
            .FontBold = False
        End If
        If InStr(1, strFont, "б") > 0 Then
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
            strFont = strFont & "," & IIf(.FontBold, "��", "") & IIf(.FontItalic, "б", "")
        End If
    End With
End Sub

Private Sub SetVsFlexGridChangeHead(ByVal strHead As String, ByRef vsgrid As VSFlexGrid, lngNO As Long)
    '���ܣ���ʼvsFlexGrid
    '           ��һ�̶��У���ʼ����ֻ��һ�м�¼���޹̶��С�
    'strHead��  �����ʽ��
    '           ����1,���,���뷽ʽ;����2,���,���뷽ʽ;.......
    '           ���뷽ʽȡֵ, * ��ʾ����ȡֵ
    '           FlexAlignLeftTop       0   ����
    '           flexAlignLeftCenter    1   ����  *
    '           flexAlignLeftBottom    2   ����
    '           flexAlignCenterTop     3   ����
    '           flexAlignCenterCenter  4   ����  *
    '           flexAlignCenterBottom  5   ����
    '           flexAlignRightTop      6   ����
    '           flexAlignRightCenter   7   ����  *
    '           flexAlignRightBottom   8   ����
    '           flexAlignGeneral       9   ����
    'vsGrid:    Ҫ��ʼ���Ŀؼ�

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
            .ColKey(i) = Split(arrHead(i), ",")(0) '��������ΪcolKeyֵ
            
            If UBound(Split(arrHead(i), ",")) > 0 Then
               'Ϊ��֧��zl9PrintMode
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
                    'Ϊ��֧��zl9PrintMode
                    .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                End If
            Else
                If .FixedCols > 0 Then
                    .ColHidden(i) = True
                    .ColWidth(i) = 0  'Ϊ��֧��zl9PrintMode
                Else
                    .ColHidden(.FixedCols + i) = True
                    .ColWidth(.FixedCols + i) = 0 'Ϊ��֧��zl9PrintMode
                End If
            End If
            .ColData(i) = Split(arrHead(i), ",")(3) '��������Ϊ����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
        Next
        
        '�̶������־���
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .ROWHEIGHT(0) = 300
        
        .WordWrap = True '�Զ�����
        .AutoSizeMode = flexAutoSizeRowHeight '�Զ��и�
        .AutoResize = True '�Զ�
        .Redraw = True
    End With
End Sub

Private Sub vfgCurve_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vfgCurve
        If Row >= .FixedRows And Col >= .FixedCols Then
            If .ColIndex("ѡ��") = Col And Val(.TextMatrix(Row, .ColIndex("��Ŀ���"))) <> 0 Then
                mblnRedraw = True
                DataChanged = True
            End If
        End If
    End With
End Sub

Private Sub vfgCurve_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vfgCurve.ColIndex("���") Or Col = vfgCurve.ColIndex("ѡ��") Then
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
                .Col = .ColIndex("ѡ��")
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
        If Val(.TextMatrix(.Row, .ColIndex("��Ŀ���"))) = 0 Then Exit Sub
        If KeyAscii = vbKeySpace Then
            If Val(.TextMatrix(.Row, .ColIndex("��Ŀ���"))) = 1 Then
                .TextMatrix(.Row, .ColIndex("ѡ��")) = 1
                Exit Sub
            Else
                strValue = .TextMatrix(.Row, .ColIndex("ѡ��"))
                .TextMatrix(.Row, .ColIndex("ѡ��")) = IIf(Val(strValue) = 1, 0, 1)
            End If
            DataChanged = True
            mblnRedraw = True
            On Error Resume Next
            If .Enabled And .Visible Then .SetFocus
        End If
    End With
End Sub

Private Sub vfgCurve_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
   If Col <> vfgCurve.ColIndex("ѡ��") Then Cancel = True: Exit Sub
   If Trim(vfgCurve.TextMatrix(Row, vfgCurve.ColIndex("��Ŀ���"))) = "" Then Cancel = True: Exit Sub
End Sub

Private Sub vfgCurve_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgCurve
        If Row >= .FixedRows And Col >= .FixedCols Then
            If .ColIndex("ѡ��") = Col And Val(.TextMatrix(Row, .ColIndex("��Ŀ���"))) = 1 Then
                .TextMatrix(Row, .ColIndex("ѡ��")) = 1
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub vfgTab_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vfgTab
        If Not (Row >= .FixedRows And Col = .ColIndex("��¼Ƶ��")) Then Exit Sub
        If .ComboIndex < 0 Then Exit Sub
        .TextMatrix(Row, Col) = .EditText
    End With
End Sub

Private Sub vfgTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    Dim StrKey As String
    With vfgTab
        If Not (NewRow >= .FixedRows And NewCol = .ColIndex("��¼Ƶ��")) Then
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
    If Col = vfgTab.ColIndex("���") Then
        Cancel = True
    End If
End Sub

Private Sub vfgTab_EnterCell()
    With vfgTab
        .ColEditMask(.Col) = ""
        .ColComboList(.Col) = ""
        .CellBorderRange .FixedRows, .FixedCols, .Rows - 1, .Cols - 1, .GridColor, 1, 1, 1, 1, 0, 0
        If (.Row >= .FixedRows And .Col = .ColIndex("��¼Ƶ��")) Then
            If Val(.TextMatrix(.Row, .ColIndex("��Ŀ���"))) = 3 Then Exit Sub
            .CellBorderRange .Row, .Col, .Row, .Col, &HFF0000, 1, 1, 1, 1, 0, 0
            '���ݼ�����������¼Ƶ��
            .ColComboList(.Col) = GetTabFrequency(Val(txtTabDayTime.Text), Val(.TextMatrix(.Row, .ColIndex("��Ŀ��ʾ"))))
        End If
    End With
End Sub

Private Sub vfgTab_KeyDown(KeyCode As Integer, Shift As Integer)
    With vfgTab
        If KeyCode = vbKeyReturn Then
            If Not .Col = .ColIndex("��¼Ƶ��") Then
                .Col = .ColIndex("��¼Ƶ��")
            Else
                If .Row >= .Rows - 1 Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    .Row = .Row + 1
                    .Col = .ColIndex("��¼Ƶ��")
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
        '��������ѡ��
        If .TextMatrix(Row, .ColIndex("��Ŀ���")) = 3 Then Cancel = True: Exit Sub
        If Col <> .ColIndex("��¼Ƶ��") Then Cancel = True: Exit Sub
        If Trim(.TextMatrix(Row, .ColIndex("��Ŀ���"))) = "" Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vfgTab_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgTab
        If Col <> .ColIndex("��¼Ƶ��") Then Exit Sub
        If .TextMatrix(Row, Col) = .EditText Then Exit Sub
    End With
    mblnRedraw = True
    DataChanged = True
End Sub


Private Function CalcScrollBarSize() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�
    '������
    '���أ� ���óɹ�����TRUE������FALSE
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
    
    'ֻ����û��ʾ�������ǲ��������㲽��
    msinHStep = (picDraw.Width - picPane(1).Width + IIf(vsb.Visible = True, vsb.Width, 0)) / 10
    msinVStep = (picDraw.Height + IIf(Picbaby.Visible = True, Picbaby.Height, 0) - picPane(1).Height + IIf(hsb.Visible = True, hsb.Height, 0)) / 10
    
    '�㶨Ϊ100,ֻ�ǲ��������仯
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
        Set cbrPopupBar = cbsThis.Add("�Ҽ��˵�", xtpBarPopup)
        cbrPopupBar.Title = "�Ҽ��˵�"
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddLeft, "�����������"): cbrPopupItem.IconId = conMenu__Curve_AddLeft
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddRight, "���Ҳ�������"): cbrPopupItem.IconId = conMenu__Curve_AddRight
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteCol, "ɾ����"): cbrPopupItem.IconId = conMenu__Curve_DeleteCol
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddUP, "���Ϸ�������"): cbrPopupItem.IconId = conMenu__Curve_AddUP
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_AddBottom, "���·�������"): cbrPopupItem.IconId = conMenu__Curve_AddBottom
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_DeleteRow, "ɾ����"): cbrPopupItem.IconId = conMenu__Curve_DeleteRow
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddySingle, "�󶨵���"): cbrPopupItem.IconId = conMenu__Curve_BuddySingle
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, conMenu__Curve_BuddyDouble, "��˫��"): cbrPopupItem.IconId = conMenu__Curve_BuddyDouble
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
    Dim intType As Integer, intFace As Integer, intLen As Integer                  '������Ŀ�ı�ʾ��ʽһ��������
    Dim bln��ʿ As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim StrText As String, strItem As String
    Dim lngCol As Long, lngCount As Long
    Dim intDo As Integer, intHead As Integer
    Dim intRow As Integer
    
    'ÿ�ֻ����¼��������Ҫ��һ�а󶨻�ʿ����
    lngCount = vfgThis.Cols - 1
   
    If Replace(vfgThis.TextMatrix(5, 1), " ", "") <> "{����}" Then
        MsgBox "��һ�б����������Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    'ֻ�а�������Ŀ���У���Ŀ�ķָ���Ҳ��/������Ŀ���ͱ�����ͬ,ֻ����¼�롢ѡ��������������������и�ʽΪ�Խ��ߣ��������������öԽ��߱�ʾ���ڼ�д��
    'ֻȡ���������ͷ��3�б�ͷȡ5��2�б�ͷȡ4��1�б�ͷȡ3)
    If optTabTiers(0).Value Then
        intHead = 3
    ElseIf optTabTiers(1).Value Then
        intHead = 4
    Else
        intHead = 5
    End If
    For lngCol = 1 To lngCount
        If Replace(vfgThis.TextMatrix(5, lngCol), " ", "") = "{ʱ��}" And lngCol <> 2 Then
            MsgBox "Ҫ��ʱ�������ڵڶ��У�", vbInformation, gstrSysName
            Exit Function
        End If
        If vfgThis.Cell(flexcpData, 5, lngCol) <> "" Then
            StrText = Val(Split(vfgThis.Cell(flexcpData, 5, lngCol), "`")(1))
            If InStr(1, vfgThis.Cell(flexcpData, 5, lngCol), "����ѹ") > 0 Then
                If Not InStr(1, vfgThis.Cell(flexcpData, 5, lngCol), "����ѹ") > 0 Then
                    MsgBox "Ҫ������ѹ���������ѹ����ʽΪ:����ѹ/����ѹ", vbInformation, gstrSysName
                End If
            End If
            If StrText = 1 Then
                '��ʽ��{A}{B}����}�ֽ⣬2�зֽ�����������=2
                StrText = vfgThis.TextMatrix(5, lngCol)
                If UBound(Split(StrText, "}")) <> 2 Then
                    If StrText <> "{����}" Then
                        MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �����������жԽ��ߣ�" & vbCrLf & "[ֻ�������С���������Ŀ���в����������жԽ���]", vbInformation, gstrSysName
                        Exit Function
                    Else
                        GoTo ntloop
                    End If
                End If
                
                '������Ŀ�ķָ���Ҳ������/
                If Trim(Mid(StrText, InStr(1, StrText, "}") + 1, InStr(InStr(1, StrText, "{") + 1, StrText, "{") - InStr(1, StrText, "}") - 1)) <> "/" Then
                    MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �������жԽ��ߣ�Ҫ��󶨵���Ŀ��ʽӦ����:A/B", vbInformation, gstrSysName
                    Exit Function
                End If
                
                '������Ŀ����Ŀ���ͱ���һ��
                For intDo = 0 To 1
                    strItem = Trim(GetItemName(StrText, intDo))
                    mrsItems.Filter = "��Ŀ����='" & strItem & "'"
                    '¼���ʱ�������,ϵͳ��������,ʱ��Ȳ�������̶����,����,��ʱ�������Ҳ��������
                    If mrsItems.RecordCount = 0 Then
                        MsgBox "��Ŀ:" & strItem & "�Ѹ������Ѿ���ɾ����", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If intDo > 0 Then
                        If Not (intFace = mrsItems!��Ŀ��ʾ And intType = mrsItems!��Ŀ����) Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵�������Ŀ�ı༭��ʽ����һ�£�", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ��ʾ = 0 And NVL(mrsItems!��Ŀ����, 1) > 3 Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵������ı���Ŀ���Ȳ��ܴ���3��", vbInformation, gstrSysName
                            Exit Function
                        End If
                    Else
                        intFace = mrsItems!��Ŀ��ʾ
                        intType = mrsItems!��Ŀ����
                        intLen = NVL(mrsItems!��Ŀ����, 1)
                        If InStr(1, "0,2,4,5", intFace) = 0 Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵�������Ŀ���붼����ֵ�͡�ѡ�����ѡ�������򳤶�С�ڵ���3���ı���Ŀ��", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If intType = 1 And intFace = 0 And intLen > 3 Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵������ı���Ŀ���Ȳ��ܴ���3��", vbInformation, gstrSysName
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
    '��ȡָ����ʽ����ָ����ŵ���Ŀ���ƣ���ʽ�磺{����ѹ}/ {����ѹ}mmHg
    
    intPos = InStr(1, strData, "{")
    If intOrder > 0 Then intPos = InStr(intPos + 1, strData, "{")
    strData = Mid(strData, intPos + 1)
    strData = Mid(strData, 1, InStr(1, strData, "}") - 1)
    GetItemName = strData
End Function

