VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmPathTrack 
   AutoRedraw      =   -1  'True
   Caption         =   "�ٴ�·������"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   16005
   FillColor       =   &H80000008&
   Icon            =   "frmPathTrack.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   16005
   Begin XtremeReportControl.ReportControl rptPath 
      Height          =   5970
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3870
      _Version        =   589884
      _ExtentX        =   6826
      _ExtentY        =   10530
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin MSComDlg.CommonDialog dlgPublic 
      Left            =   11520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "����ΪͼƬ"
      Filter          =   "Jpeg|*.jpg"
   End
   Begin VB.PictureBox picReason 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   3600
      ScaleHeight     =   4335
      ScaleWidth      =   8295
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3720
      Width           =   8295
      Begin VB.PictureBox picTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Left            =   3120
         ScaleHeight     =   4575
         ScaleWidth      =   7095
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   7095
         Begin VB.TextBox txtFindNum 
            Height          =   270
            Left            =   4800
            TabIndex        =   37
            Text            =   "5"
            Top             =   240
            Width           =   350
         End
         Begin VB.Frame fraGroupUD 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   45
            Left            =   360
            MousePointer    =   7  'Size N S
            TabIndex        =   46
            Top             =   1800
            Width           =   2000
         End
         Begin VB.Frame fraGroupLR 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1000
            Left            =   2280
            MousePointer    =   9  'Size W E
            TabIndex        =   45
            Top             =   240
            Width           =   45
         End
         Begin VSFlex8Ctl.VSFlexGrid vsgInfo 
            Height          =   1455
            Index           =   0
            Left            =   360
            TabIndex        =   42
            Top             =   360
            Width           =   1815
            _cx             =   3201
            _cy             =   2566
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
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
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
         Begin VSFlex8Ctl.VSFlexGrid vsgInfo 
            Height          =   1455
            Index           =   1
            Left            =   2520
            TabIndex        =   43
            Top             =   360
            Width           =   1815
            _cx             =   3201
            _cy             =   2566
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
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
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
         Begin VSFlex8Ctl.VSFlexGrid vsgInfo 
            Height          =   1455
            Index           =   2
            Left            =   360
            TabIndex        =   44
            Top             =   2280
            Width           =   4095
            _cx             =   7223
            _cy             =   2566
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
            GridColor       =   -2147483638
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
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
         Begin VB.Image imgFrom 
            Height          =   300
            Left            =   3360
            Picture         =   "frmPathTrack.frx":058A
            Stretch         =   -1  'True
            Top             =   1920
            Width           =   345
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "δ������Ŀ��ϸ��"
            Height          =   250
            Index           =   2
            Left            =   360
            TabIndex        =   49
            Top             =   2040
            Width           =   1440
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "δ������Ŀ���ܱ�(˫���鿴��Ӧҽ��)"
            Height          =   250
            Index           =   1
            Left            =   2520
            TabIndex        =   48
            Top             =   120
            Width           =   3060
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "δ����ԭ����ܱ�"
            Height          =   250
            Index           =   0
            Left            =   360
            TabIndex        =   47
            Top             =   120
            Width           =   1440
         End
      End
      Begin C1Chart2D8.Chart2D chtThis 
         Height          =   2295
         Left            =   600
         TabIndex        =   32
         Top             =   600
         Width           =   5295
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   9340
         _ExtentY        =   4048
         _StockProps     =   0
         ControlProperties=   "frmPathTrack.frx":6DDC
      End
      Begin VB.Label lblMSG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ʾ��Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   40
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.PictureBox picVariation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   4920
      ScaleHeight     =   7935
      ScaleWidth      =   11055
      TabIndex        =   14
      Top             =   2160
      Width           =   11055
      Begin VB.PictureBox picContrast 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   12735
         TabIndex        =   68
         Top             =   480
         Width           =   12735
         Begin VB.CheckBox chkContrast 
            BackColor       =   &H80000005&
            Caption         =   "��ָ���ڼ�Ƚ�"
            Height          =   255
            Left            =   4530
            TabIndex        =   73
            Top             =   53
            Width           =   1575
         End
         Begin VB.CommandButton cmdContrast 
            Caption         =   "�Ա�(&C)"
            Height          =   345
            Left            =   9600
            TabIndex        =   70
            Top             =   8
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cboYorM 
            Height          =   300
            Left            =   90
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   30
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpTwo 
            Height          =   300
            Left            =   6570
            TabIndex        =   71
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpOne 
            Height          =   300
            Left            =   1530
            TabIndex        =   72
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpFour 
            Height          =   300
            Left            =   8130
            TabIndex        =   82
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpThree 
            Height          =   300
            Left            =   3090
            TabIndex        =   85
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy��MM��"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin VB.Label lblFromToTwo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��               ��"
            Height          =   180
            Left            =   6330
            TabIndex        =   84
            Top             =   90
            Width           =   1710
         End
         Begin VB.Label lblFromToOne 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��               ��"
            Height          =   180
            Left            =   1290
            TabIndex        =   83
            Top             =   90
            Width           =   1710
         End
      End
      Begin VB.PictureBox picFind 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   7695
         TabIndex        =   61
         Top             =   480
         Width           =   7695
         Begin VB.ComboBox cboTimeType 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   30
            Width           =   1095
         End
         Begin VB.ComboBox cboPathTime 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   30
            Width           =   1245
         End
         Begin VB.CommandButton cmdVariation 
            Caption         =   "ͳ��(&T)"
            Height          =   345
            Left            =   5880
            TabIndex        =   62
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   4560
            TabIndex        =   65
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   2790
            TabIndex        =   66
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin VB.Label lblFromTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��                 ��"
            Height          =   180
            Left            =   2460
            TabIndex        =   67
            Top             =   90
            Width           =   1890
         End
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Index           =   1
         Left            =   6885
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   60
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.ComboBox cboPathEdition 
         Height          =   300
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   60
         Visible         =   0   'False
         Width           =   1040
      End
      Begin VB.Frame fraGroupLine 
         Height          =   30
         Left            =   120
         TabIndex        =   31
         Top             =   405
         Width           =   5175
      End
      Begin VB.OptionButton optThisPath 
         BackColor       =   &H80000005&
         Caption         =   "��ǰ·��"
         Height          =   180
         Left            =   1230
         TabIndex        =   29
         Top             =   120
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton optAllPath 
         BackColor       =   &H80000005&
         Caption         =   "ȫԺ·��"
         Height          =   180
         Left            =   2400
         TabIndex        =   30
         Top             =   120
         Width           =   1200
      End
      Begin XtremeSuiteControls.TabControl tbcVariation 
         Height          =   2775
         Left            =   0
         TabIndex        =   38
         Top             =   1080
         Width           =   4695
         _Version        =   589884
         _ExtentX        =   8281
         _ExtentY        =   4895
         _StockProps     =   64
      End
      Begin VB.PictureBox picTrend 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   375
         ScaleWidth      =   11805
         TabIndex        =   74
         Top             =   480
         Width           =   11800
         Begin VB.OptionButton optOut 
            BackColor       =   &H80000005&
            Caption         =   "��·������"
            Height          =   180
            Left            =   9210
            TabIndex        =   81
            Top             =   90
            Width           =   1200
         End
         Begin VB.OptionButton optIn 
            BackColor       =   &H80000005&
            Caption         =   "·������"
            Height          =   180
            Left            =   8160
            TabIndex        =   80
            Top             =   90
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.ComboBox cboInterval 
            Height          =   300
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   30
            Width           =   1095
         End
         Begin VB.ComboBox cboTrendTime 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   30
            Width           =   1095
         End
         Begin VB.CommandButton cmdTrend 
            Caption         =   "��ѯ(&Q)"
            Height          =   335
            Left            =   5520
            TabIndex        =   75
            Top             =   0
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpTrendStart 
            Height          =   300
            Left            =   2160
            TabIndex        =   77
            Top             =   30
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd��"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin VB.Label lblTrend 
            BackColor       =   &H80000005&
            Caption         =   "��ʼʱ��                   ֮��"
            Height          =   255
            Left            =   1380
            TabIndex        =   78
            Top             =   85
            Width           =   3975
         End
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·������"
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
         Index           =   1
         Left            =   6000
         TabIndex        =   60
         Top             =   120
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblPathEdition 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·���汾"
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
         Left            =   3780
         TabIndex        =   50
         Top             =   120
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblZY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ע��"
         Height          =   660
         Left            =   360
         TabIndex        =   39
         Top             =   6360
         Width           =   360
      End
      Begin VB.Label lblPathType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "·������"
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
         TabIndex        =   28
         Top             =   120
         Width           =   780
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   2535
      Left            =   240
      TabIndex        =   4
      Top             =   6840
      Width           =   4575
      _Version        =   589884
      _ExtentX        =   8070
      _ExtentY        =   4471
      _StockProps     =   64
   End
   Begin VB.Frame fraLR_S 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5400
      Left            =   4215
      MousePointer    =   9  'Size W E
      TabIndex        =   2
      Top             =   960
      Width           =   45
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7590
      TabIndex        =   1
      Top             =   210
      Width           =   2175
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   9870
      Width           =   16005
      _ExtentX        =   28231
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPathTrack.frx":743B
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12647
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12647
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
   Begin MSComctlLib.ImageList img16 
      Left            =   1260
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrack.frx":7CCD
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrack.frx":8267
            Key             =   "PatiMan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrack.frx":8801
            Key             =   "PatiWoman"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrack.frx":8D9B
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrack.frx":9335
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrack.frx":98CF
            Key             =   "������"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6615
      Left            =   2400
      ScaleHeight     =   6615
      ScaleWidth      =   11175
      TabIndex        =   5
      Top             =   600
      Width           =   11175
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   3360
         Left            =   -1200
         TabIndex        =   6
         Top             =   1560
         Width           =   5520
         _Version        =   589884
         _ExtentX        =   9737
         _ExtentY        =   5927
         _StockProps     =   0
         BorderStyle     =   2
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1545
         Left            =   0
         ScaleHeight     =   1545
         ScaleWidth      =   10545
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   10545
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Index           =   0
            Left            =   4725
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   795
            Width           =   2460
         End
         Begin VB.ComboBox cbo·����Χ 
            Height          =   300
            Left            =   7200
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   1140
            Width           =   2055
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "�����˳�"
            Height          =   180
            Index           =   3
            Left            =   4920
            TabIndex        =   52
            Top             =   495
            Width           =   1020
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "������"
            Height          =   180
            Index           =   0
            Left            =   1245
            TabIndex        =   13
            Top             =   495
            Width           =   840
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "ִ����"
            Height          =   180
            Index           =   1
            Left            =   2070
            TabIndex        =   15
            Top             =   495
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "��������"
            Height          =   180
            Index           =   2
            Left            =   2880
            TabIndex        =   17
            Top             =   495
            Width           =   1020
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "�������"
            Height          =   180
            Index           =   4
            Left            =   3885
            TabIndex        =   19
            Top             =   495
            Width           =   1020
         End
         Begin VB.ComboBox cboForDate 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   1155
            Width           =   1095
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1155
            Width           =   1245
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "����(&F)"
            Height          =   350
            Left            =   9360
            TabIndex        =   36
            Top             =   1125
            Visible         =   0   'False
            Width           =   1100
         End
         Begin VB.TextBox txtPerson 
            Height          =   270
            Left            =   1250
            TabIndex        =   10
            Top             =   90
            Width           =   1695
         End
         Begin VB.ComboBox cboTyping 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   795
            Width           =   1695
         End
         Begin VB.Frame fraGroup 
            Height          =   30
            Left            =   0
            TabIndex        =   8
            Top             =   370
            Width           =   5175
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   4725
            TabIndex        =   35
            Top             =   1155
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   3000
            TabIndex        =   34
            Top             =   1155
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   153812995
            CurrentDate     =   40256
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "·������"
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
            Index           =   0
            Left            =   3840
            TabIndex        =   57
            Top             =   840
            Width           =   780
         End
         Begin VB.Label lbl·����Χ 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "·����Χ"
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
            Left            =   6240
            TabIndex        =   53
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��·��״̬"
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
            TabIndex        =   11
            Top             =   495
            Width           =   975
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��                 ��"
            Height          =   180
            Left            =   2670
            TabIndex        =   33
            Top             =   1215
            Width           =   1890
         End
         Begin VB.Label lblPerson 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����˲�ѯ"
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
            TabIndex        =   9
            Top             =   135
            Width           =   975
         End
         Begin VB.Label lblTyping 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   330
            TabIndex        =   21
            Top             =   855
            Width           =   780
         End
      End
      Begin VB.Frame fraUD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   24
         Top             =   5000
         Width           =   6495
      End
      Begin VB.PictureBox picInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   240
         ScaleHeight     =   1575
         ScaleWidth      =   7455
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4680
         Width           =   7455
         Begin XtremeReportControl.ReportControl rptOper 
            Height          =   615
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   5415
            _Version        =   589884
            _ExtentX        =   9551
            _ExtentY        =   1085
            _StockProps     =   0
            BorderStyle     =   2
            MultipleSelection=   0   'False
            EditOnClick     =   0   'False
            AutoColumnSizing=   0   'False
         End
         Begin VB.Label lblMerge 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "δ����ԭ����ܱ�"
            Height          =   255
            Left            =   240
            TabIndex        =   56
            Top             =   600
            Visible         =   0   'False
            Width           =   5880
         End
         Begin VB.Label lblMerges 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ϲ�·����"
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
            TabIndex        =   55
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDiagInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   1200
            TabIndex        =   22
            Top             =   120
            Width           =   90
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ҫ���"
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
            TabIndex        =   20
            Top             =   120
            Width           =   780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTemp 
         Height          =   900
         Left            =   6390
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1935
         Visible         =   0   'False
         Width           =   1080
         _cx             =   1905
         _cy             =   1587
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   2000
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
   Begin VB.Frame fraFlag 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4080
      TabIndex        =   86
      Top             =   120
      Visible         =   0   'False
      Width           =   240
      Begin VB.Image imgFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "frmPathTrack.frx":10131
         Top             =   0
         Width           =   240
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   420
      Top             =   165
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModul As Long
Private mblnShowStoped As Boolean
Private mblnFirstLoad As Boolean '�ж��Ƿ��ǵ�һ�μ���
Private mrsTmp As Recordset
Private mobjCISJob As Object
Private mstrFlag As String     '��ǵ�ǰѡ�б��

Private Enum COL_PATH_LIST
    COL_ID = 0
    COL_ͼ�� = 1
    COL_���� = 2
    COL_���� = 3
    COL_���� = 4
    COL_�������� = 5
    COL_���ò��� = 6
    COL_�����Ա� = 7
    COL_�������� = 8
    COL_˵�� = 9
    COL_ͨ�� = 10
    COL_�����ڼ� = 11
    COL_��ϱ��� = 12
    COL_������� = 13
    COL_�������� = 14
    COL_�������� = 15
End Enum

Private Enum COL_PATI_LIST
    col_��ӡ = 0
    COL_����ID = 1
    COL_��ҳID = 2
    COL_���� = 3
    COL_���� = 4
    COL_�Ա� = 5
    COL_���� = 6
    COL_סԺ�� = 7
    COL_���� = 8
    COL_סԺҽʦ = 9
    COL_���� = 10
    COl_���� = 11
    COl_��׼סԺ�� = 12
    COL_��׼���� = 13
    COL_�汾�� = 14
    COL_������ = 15
    COl_����ʱ�� = 16
    COL_����ʱ�� = 17
    COL_����ID = 18
    COL_����ID = 19
    COL_����ת�� = 20
    COL_����״̬ = 21
    COL_������ԭ�� = 22
    COL_�����˳�ԭ�� = 23
    COL_��ӡ�� = 24
    COL_��ӡʱ�� = 25
    COL_��֧·�� = 26
    COL_��֧·����� = 27
    COL_�ϲ�·������ = 28
    COL_������ = 29
    COL_���߰��ӡ = 30
    col_����·��ID = 31
End Enum

Private Enum CONST_AREA
    Area_Cross = 0
    Area_Category = 1
    Area_Step = 2
    Area_Item = 3
End Enum

Private Enum COL_OPER_LIST
    COL_��¼ID = 0
    COL_�������� = 1
    COL_�������� = 2
    COL_����ҽʦ = 3
    COL_����ҽʦ = 4
End Enum

Private Enum VSG_Info
    vsg_ԭ�� = 0
    vsg_��Ŀ = 1
    VSG_��ϸ = 2
End Enum

Private Enum COL_VSG_Info
    VCol_���� = 0
    VCol_ԭ�� = 1
    VCol_�׶� = 0
    VCOL_���� = 0
    VCol_���� = 0
    VCol_ԭ������ = 2
    VCol_���� = 1
    VCol_סԺ�� = 1
    VCol_��Ŀ���� = 2
    VCOL_ҽ�� = 2
    VCol_δʹ��ԭ�� = 3
    VCol_����ʱ�� = 4
    VCOL_ҽ������ = 1
    VCOL_������ = 2
    vcol_�뾶���� = 3
    vcol_�뾶�� = 4
    vcol_�����˳��� = 5
    vcol_�����˳��� = 6
    vcol_��������� = 7
    vcol_��������� = 8
    VCOL_ҽ�����϶� = 9
    VCOL_ָ�� = 0
    VCOL_ͬ��һ = 1
    VCOL_ͬ�ڶ� = 2
    VCOL_��ֵ = 3
End Enum

Private Const conMenu_View_FindName = 7211                 '*��·�����Ʋ���(&F)
Private Const conMenu_View_FindIll = 7212                 '*��������ϲ���(&F)
Private mlng����ID As Long, mlng��ҳID As Long, mlng����·��ID As Long
Private mlngVariation As Long, mlngSurvey As Long, mlngTrend As Long
Private mblnIsPathTo As Boolean
Private mblnIsEdition As Boolean
Private mlngOldPathID As Long      '��һ�β�ѯ��·��id
Private mdateOldStart As Date      '��һ�εĿ�ʼʱ��
Private mdateOldEnd As Date       '��һ�εĽ���ʱ��
Private mstrDateType As String     '��һ�ε�ʱ������
Private mlng·��ID As Long   '�ϴ�ѡ���·��ID

Private Sub cboForDate_Click()
    If cboTime.ListIndex <> cboTime.ListCount - 1 Then
        Call rptPath_SelectionChanged
    End If
End Sub

Private Sub cboPathEdition_Click()
    mblnIsEdition = True
    If tbcSub.Selected.Tag <> "����·��" Then
        Call tbcVariation_SelectedChanged(tbcVariation.Selected)
    End If
End Sub

Private Sub cboPathTime_Click()
    Dim curDate As Date
    
    dtpStart.Enabled = cboPathTime.ListIndex = cboPathTime.ListCount - 1
    dtpEnd.Enabled = cboPathTime.ListIndex = cboPathTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpStart.MaxDate = curDate
    dtpEnd.MaxDate = curDate
    cmdVariation.Visible = False
    
    Select Case cboPathTime.ListIndex
    Case 0 '����
        dtpStart.Value = Format(curDate, "yyyy-MM-dd")
        dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
    Case 1 '���һ��
        dtpStart.Value = DateAdd("ww", -1, curDate)
        dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
    Case 2 '���һ��
        dtpStart.Value = DateAdd("m", -1, curDate)
        dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
    Case 3 '���һ��
        dtpStart.Value = DateAdd("q", -1, curDate)
        dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
    Case 4 '�������
        dtpStart.Value = DateAdd("m", -6, curDate)
        dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
    Case 5 '���һ��
        dtpStart.Value = DateAdd("yyyy", -1, curDate)
        dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
    Case 6 'ָ  ��
        dtpStart.SetFocus
        cmdVariation.Visible = True
    End Select
    
    If cboPathTime.ListIndex <> cboPathTime.ListCount - 1 Then
        Call rptPath_SelectionChanged
    End If
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zlDatabase.Currentdate
    dtpTime(0).MaxDate = curDate
    dtpTime(1).MaxDate = curDate
    cmdFind.Visible = False
    
    Select Case cboTime.ListIndex
    Case 0 '����
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
    Case 1 '���һ��
        dtpTime(0).Value = DateAdd("ww", -1, curDate)
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
    Case 2 '���һ��
        dtpTime(0).Value = DateAdd("m", -1, curDate)
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
    Case 3 '���һ��
        dtpTime(0).Value = DateAdd("q", -1, curDate)
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
    Case 4 '�������
        dtpTime(0).Value = DateAdd("m", -6, curDate)
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
    Case 5 '���һ��
        dtpTime(0).Value = DateAdd("yyyy", -1, curDate)
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
    Case 6 'ָ  ��
        dtpTime(0).SetFocus
        cmdFind.Visible = True
    End Select
    
    If cboTime.ListIndex <> cboTime.ListCount - 1 Then
        Call rptPath_SelectionChanged
    End If
End Sub

Private Sub cboTimeType_Click()
    If tbcSub.Selected.Tag <> "����·��" Then
        Call tbcVariation_SelectedChanged(tbcVariation.Selected)
    End If
End Sub

Private Sub cboTrendTime_Click()
    cboInterval.Clear
    If cboTrendTime.ListIndex = 0 Then
        '����
        cboInterval.AddItem "һ��"
        cboInterval.AddItem "һ��"
        cboInterval.AddItem "����"
        cboInterval.AddItem "һ����"
        dtpTrendStart.CustomFormat = "yyyy��MM��dd��"
    Else
        cboInterval.AddItem "����"
        cboInterval.AddItem "һ��"
        cboInterval.AddItem "����"
        cboInterval.AddItem "����"
        dtpTrendStart.CustomFormat = "yyyy��MM��"
    End If
    cboInterval.ListIndex = 1
End Sub

Private Sub cboTyping_Click()
    If cboTime.ListIndex <> cboTime.ListCount - 1 Then
        Call rptPath_SelectionChanged
    End If
End Sub

Private Sub cboYorM_Click()
    If cboYorM.ListIndex = 0 Then
        dtpOne.CustomFormat = "yyyy��MM��"
        dtpTwo.CustomFormat = "yyyy��MM��"
    ElseIf cboYorM.ListIndex = 1 Then
        dtpOne.CustomFormat = "yyyy��MM��"
        dtpTwo.CustomFormat = "yyyy��MM��"
        dtpThree.CustomFormat = "yyyy��MM��"
        dtpFour.CustomFormat = "yyyy��MM��"
    ElseIf cboYorM.ListIndex = 2 Then
        dtpOne.CustomFormat = "yyyy��"
        dtpTwo.CustomFormat = "yyyy��"
    End If
    If tbcSub.Selected.Tag <> "����·��" Then
        Call tbcVariation_SelectedChanged(tbcVariation.Selected)
    End If
End Sub

Private Sub cbo·����Χ_Click()
    If cboTime.ListIndex <> cboTime.ListCount - 1 Then
        Call rptPath_SelectionChanged
    End If
End Sub

Private Sub cbo����_Click(Index As Integer)
    
    If rptPath.SelectedRows.count <> 0 And cbo����(Index).ListIndex >= 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            With rptPath.SelectedRows(0)
                If Index = 0 Then
                    Call LoadPatiList(.Record(COL_ID).Value, , cbo����(Index).ItemData(cbo����(Index).ListIndex))
                Else
                    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
                End If
            End With
        End If
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim lng·��ID As Long
    Dim objPopup As CommandBarPopup
    
    If Control.ID <> 0 And Control.ID <> conMenu_View_FindNext Then
        If cbsMain.FindControl(, Control.ID, True, True) Is Nothing Then Exit Sub
    End If
    
    Select Case Control.ID
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview: Call zlRptPrint(2)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    Case conMenu_File_BatPrint: Call zlRptBatPrint
    Case conMenu_File_SaveJpeg: Call SaveImage
        
    Case conMenu_Tool_Report    '������ͳ�Ʊ�
         Call FuncShowReport
    Case conMenu_View_Show '�鿴·����
        Call FuncShowPath
    Case conMenu_Edit_OutLogView
        If rptPath.SelectedRows.count > 0 Then
            If Not rptPath.SelectedRows(0).GroupRow Then
                lng·��ID = rptPath.SelectedRows(0).Record(COL_ID).Value
            End If
        End If
        Call frmPathOutLog.ShowMe(Me, mlng����ID, mlng��ҳID, 1, Nothing, lng·��ID, mlng����·��ID)
        
    Case conMenu_View_ShowStoped
        mblnShowStoped = Not mblnShowStoped
        Control.Checked = mblnShowStoped
        Call LoadPathList
    Case conMenu_View_Find '����
        If Me.ActiveControl Is txtFind Then
            txtFind.SetFocus '��ʱ��Ҫ��λһ��
            If txtFind.Text <> "" Then
                Call FuncFindPath
            End If
        Else
            txtFind.SetFocus
        End If
    Case conMenu_View_FindNext '������һ��
        If txtFind.Text = "" Then
            txtFind.SetFocus
        Else
            Call FuncFindPath(True)
        End If
    Case conMenu_View_ToolBar_Button '������
        For i = 2 To cbsMain.count
            Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Text '��ť����
        For i = 2 To cbsMain.count
            For Each objControl In Me.cbsMain(i).Controls
                objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
        Next
        Me.cbsMain.RecalcLayout
    Case conMenu_View_ToolBar_Size '��ͼ��
        Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
        Me.cbsMain.RecalcLayout
    Case conMenu_View_StatusBar '״̬��
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsMain.RecalcLayout
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        If rptPath.SelectedRows.count > 0 Then
            If rptPath.SelectedRows(0).GroupRow Then
                rptPath.SelectedRows(0).Expanded = False
            ElseIf Not rptPath.SelectedRows(0).ParentRow Is Nothing Then
                If rptPath.SelectedRows(0).ParentRow.GroupRow Then
                    rptPath.SelectedRows(0).ParentRow.Expanded = False
                End If
            End If
        End If
        '���۵���λ��������,�����Զ�������¼�
        Call rptPath_SelectionChanged
    Case conMenu_Tool_Archive '���Ӳ�������
        If mobjCISJob Is Nothing Then
            On Error Resume Next
            Set mobjCISJob = CreateObject("ZL9CISJob.clsCISJOB")
            Err.Clear: On Error GoTo 0
        End If
        If mobjCISJob Is Nothing Then
                MsgBox "�ٴ�����δ����ȷ��װ���鿴�����������ܼ�����", vbInformation, gstrSysName
            Exit Sub
        Else
            Call mobjCISJob.ShowArchive(Me, mlng����ID, mlng��ҳID)
        End If
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        If rptPath.SelectedRows.count > 0 Then
            rptPath.SelectedRows(0).Expanded = True
        End If
    Case conMenu_View_Expend_AllCollapse '�۵�������
        For Each objRow In rptPath.Rows
            If objRow.GroupRow Then objRow.Expanded = False
        Next
        '���۵���λ��������,�����Զ�������¼�
        Call rptPath_SelectionChanged
    Case conMenu_View_Expend_AllExpend 'չ��������
        For Each objRow In rptPath.Rows
            If objRow.GroupRow Then objRow.Expanded = True
        Next
    Case conMenu_View_Refresh 'ˢ��
        Call LoadPathList
        
    Case conMenu_Help_Web_Home 'Web�ϵ�����
        Call zlHomePage(Me.Hwnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.Hwnd)
    Case conMenu_Help_Web_Mail '���ͷ���
        Call zlMailTo(Me.Hwnd)
    Case conMenu_Help_About '����
        Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
    Case conMenu_Help_Help '����
        Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_File_Exit '�˳�
        Unload Me
    Case conMenu_View_FindName '��·�����Ʋ���
        Set objPopup = cbsMain.FindControl(, Control.Parent.BarID)
        objPopup.Caption = Control.Caption
    Case conMenu_View_FindIll '��������ϲ���
        Set objPopup = cbsMain.FindControl(, Control.Parent.BarID)
        objPopup.Caption = Control.Caption
    Case Else
        If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
            If rptPath.SelectedRows.count > 0 Then
                If Not rptPath.SelectedRows(0).GroupRow Then
                    lng·��ID = rptPath.SelectedRows(0).Record(COL_ID).Value
                End If
            End If
            
            'ִ�з�������ǰģ��ı���
            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                   "·��=" & lng·��ID, "����ID=" & mlng����ID, "��ҳID=" & mlng��ҳID)
        End If
    
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    With Me.rptPath
        .Left = lngLeft: .Top = lngTop
        .Height = lngBottom - lngTop
        .Width = fraLR_S.Left
        If .Width > Me.Width - 3500 Then .Width = Me.Width - 3500
    End With
    
    With Me.fraLR_S
        If Not mblnFirstLoad Then .Left = Me.rptPath.Left + Me.rptPath.Width
        .Top = Me.rptPath.Top
        .Height = Me.rptPath.Height
    End With
        
    With Me.tbcSub
    
        .Left = fraLR_S.Left + fraLR_S.Width
        .Top = lngTop
        .Height = rptPath.Height
        .Width = lngRight - .Left
    
    End With

    Me.Refresh
End Sub

Private Sub SetControlVisible(ByVal Control As CommandBarControl)
'���ܣ�����Ȩ�����ò˵��͹������Ŀɼ�״̬
    Dim blnVisible As Boolean, strItem As String

    'Ȩ��ֻ���ж�һ��,�Ѿ��жϹ�����������ж�
    If Control.Category = "���ж�" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Tool_Report
            If InStr(";" & mstrPrivs & ";", ";������ͳ�Ʊ�;") = 0 Then blnVisible = False
        Case conMenu_Edit_OutLogView
            blnVisible = CheckPathOutLog
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "���ж�"
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
        
    '����Ȩ�����ð�ť�ɼ�״̬
    Call SetControlVisible(Control)
    If Not Control.Visible And Control.ID <> conMenu_File_SaveJpeg Then Exit Sub
        
    Select Case Control.ID
    Case conMenu_View_Show, conMenu_Edit_OutLogView '�鿴·����,�鿴�����ǼǱ�
        blnEnabled = mlng����ID > 0
        
        If Control.ID = conMenu_Edit_OutLogView And blnEnabled Then
            blnEnabled = optState(2).Value Or optState(3).Value Or optState(4).Value
        End If
        Control.Enabled = blnEnabled
    Case conMenu_File_SaveJpeg '����ͼƬ
        Control.Enabled = chtThis.Visible
        Control.Visible = chtThis.Visible
    Case conMenu_Tool_Report    '������ͳ�Ʊ�
        blnEnabled = rptPath.SelectedRows.count > 0
        If blnEnabled Then
            blnEnabled = Not rptPath.SelectedRows(0).GroupRow
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_ToolBar_Button '������
        If cbsMain.count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
        End If
    Case conMenu_View_ToolBar_Text 'ͼ������
        If cbsMain.count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        End If
    Case conMenu_View_ToolBar_Size '��ͼ��
        Control.Checked = Me.cbsMain.Options.LargeIcons
    Case conMenu_View_FindNext '������һ��
        Control.Visible = False
    Case conMenu_View_StatusBar '״̬��
        Control.Checked = Me.stbThis.Visible
    Case conMenu_View_Expend_CurExpend 'չ����ǰ��
        blnEnabled = False
        If rptPath.SelectedRows.count > 0 Then
            If rptPath.SelectedRows(0).GroupRow Then
                blnEnabled = Not rptPath.SelectedRows(0).Expanded
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend_CurCollapse '�۵���ǰ��
        blnEnabled = False
        If rptPath.SelectedRows.count > 0 Then
            If rptPath.SelectedRows(0).GroupRow Then
                blnEnabled = rptPath.SelectedRows(0).Expanded
            ElseIf Not rptPath.SelectedRows(0).ParentRow Is Nothing Then
                If rptPath.SelectedRows(0).ParentRow.GroupRow Then
                    blnEnabled = rptPath.SelectedRows(0).ParentRow.Expanded
                End If
            End If
        End If
        Control.Enabled = blnEnabled
    Case conMenu_View_Expend '�۵�/չ����
        Control.Enabled = rptPath.GroupsOrder.count > 0 And rptPath.Rows.count > 0
    Case conMenu_Tool_Archive '���Ӳ�������
        If GetInsidePrivs(p���Ӳ�������) = "" Then
            Control.Visible = False
        Else
            Control.Enabled = mlng����ID <> 0
        End If
    End Select
End Sub

Private Sub chkContrast_Click()
    dtpTwo.Enabled = chkContrast.Value = 1
    cmdContrast.Visible = chkContrast.Value = 1
    
    If dtpTwo.Enabled Then
        If cboYorM.ListIndex = 0 Then
            dtpTwo.Value = Format(CDate(Format(dtpOne.Value, "yyyy-mm")) - 1, "yyyy-MM-01")
        ElseIf cboYorM.ListIndex = 1 Then
            dtpTwo.Value = Format(DateAdd("M", -3, CDate(dtpOne.Value)), dtpTwo.CustomFormat)
        Else
            dtpTwo.Value = Format(Format(dtpOne.Value, "yyyy") - 1 & "-01-01", "yyyy-MM-dd 00:00:00")
        End If
        vsgInfo(vsg_ԭ��).ColHidden(VCOL_ͬ�ڶ�) = False
        vsgInfo(vsg_ԭ��).ColHidden(VCOL_��ֵ) = False
    Else
        vsgInfo(vsg_ԭ��).ColHidden(VCOL_ͬ�ڶ�) = True
        vsgInfo(vsg_ԭ��).ColHidden(VCOL_��ֵ) = True
    End If
    
    If cboYorM.ListIndex = 1 Then
        dtpThree.Value = Format(DateAdd("M", 2, CDate(dtpOne.Value)), dtpThree.CustomFormat)
        dtpFour.Value = Format(DateAdd("M", 2, CDate(dtpTwo.Value)), dtpFour.CustomFormat)
    End If
End Sub

Private Sub cmdContrast_Click()
    Dim lngPathID As Long
    If rptPath.SelectedRows.count > 0 Or optAllPath.Value Then
        If Not rptPath.SelectedRows(0).GroupRow Or optAllPath.Value Then
            If rptPath.SelectedRows.count > 0 And Not rptPath.SelectedRows(0).GroupRow Then lngPathID = Val(rptPath.SelectedRows(0).Record(COL_ID).Value)
        End If
    End If
    Call set��������ȶ�(lngPathID)
End Sub

Private Sub cmdFind_Click()
    Call rptPath_SelectionChanged
End Sub

Private Sub cmdTrend_Click()
    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
End Sub

Private Sub cmdVariation_Click()
    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
End Sub

Private Sub dtpOne_Change()
    If tbcSub.Selected.Tag <> "����·��" Then
        Call tbcVariation_SelectedChanged(tbcVariation.Selected)
    End If
End Sub

Private Sub dtpTime_Change(Index As Integer)
    If Index = 0 Then
        dtpTime(1).MinDate = dtpTime(0).Value
    End If
End Sub

Private Sub dtpTwo_Change()
    If cboYorM.ListIndex = 1 Then
        dtpFour.Value = Format(DateAdd("M", 2, CDate(dtpTwo.Value)), dtpFour.CustomFormat)
    End If
End Sub

Private Sub Form_Activate()
    mblnFirstLoad = False
End Sub

Private Sub InitTable(vsgInfo As VSFlexGrid, ByVal strHead As String)
    Dim arrHead As Variant, i As Long
    
    arrHead = Split(strHead, ";")
    With vsgInfo
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
        
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
    End With
End Sub

Private Sub Form_Load()
    Dim strHead As String
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    Call zlCommFun.SetWindowsInTaskBar(Me.Hwnd, False)
    mblnFirstLoad = True
    mblnIsPathTo = True
    mlng����ID = 0
    mlng��ҳID = 0
    mlng����·��ID = 0
    mstrFlag = ""
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "����·��", picPati.Hwnd, 0).Tag = "����·��"
        .InsertItem(1, "�������", picVariation.Hwnd, 0).Tag = "�������"
        .InsertItem(2, "�ſ�����", picVariation.Hwnd, 0).Tag = "�ſ�����"
        .InsertItem(3, "���Ʒ���", picVariation.Hwnd, 0).Tag = "���Ʒ���"
        
        .Item(0).Selected = True
    End With
    
     With tbcVariation
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        .InsertItem(0, "��ҽ��ͳ��", picReason.Hwnd, 0).Tag = "��ҽ��ͳ��"
        .InsertItem(1, "���ұ���������", picReason.Hwnd, 0).Tag = "���ұ���������"
        .InsertItem(2, "����֢����", picReason.Hwnd, 0).Tag = "����֢����"
        .InsertItem(3, "δ����ԭ��", picReason.Hwnd, 0).Tag = "δ����ԭ��"
        .InsertItem(4, "δ����ԭ��", picReason.Hwnd, 0).Tag = "δ����ԭ��"
        .InsertItem(5, "·������Ŀ", picReason.Hwnd, 0).Tag = "·������Ŀ"
        .InsertItem(6, "ʱ��������", picReason.Hwnd, 0).Tag = "ʱ��������"
        .InsertItem(7, "�����˳�����", picReason.Hwnd, 0).Tag = "�����˳�����"
        .InsertItem(8, "·��������", picReason.Hwnd, 0).Tag = "·��������"
        .InsertItem(9, "�׶�ƽ������", picReason.Hwnd, 0).Tag = "�׶�ƽ������"
        .InsertItem(10, "סԺ�շֲ�ͼ", picReason.Hwnd, 0).Tag = "סԺ�շֲ�ͼ"
        .InsertItem(11, "�������", picReason.Hwnd, 0).Tag = "�������"
        .InsertItem(12, "ƽ��סԺ����", picReason.Hwnd, 0).Tag = "ƽ��סԺ����"
        .InsertItem(13, "ƽ��סԺ����", picReason.Hwnd, 0).Tag = "ƽ��סԺ����"
        .InsertItem(14, "�뾶��", picReason.Hwnd, 0).Tag = "�뾶��"
        .InsertItem(15, "�����", picReason.Hwnd, 0).Tag = "�����"
        .InsertItem(16, "������", picReason.Hwnd, 0).Tag = "������"
    End With
    tbcVariation.Item(tbcVariation.ItemCount - 1).Selected = True
    tbcVariation.Item(0).Selected = True
    
    'vsFlexGrid
    '-----------------------------------------------------
    
    strHead = "��������,1500,1;סԺ��,1500,1;ҽ��,1500,1;δʹ��ԭ��,3200,1;����ʱ��,3000,1"
    Call InitTable(vsgInfo(VSG_��ϸ), strHead)
    vsgInfo(VSG_��ϸ).ExplorerBar = flexExSortShowAndMove
    
    'ReportControl
    '-----------------------------------------------------
    Call InitPathReportColumn
    Call InitPatiReportColumn
    Call InitOperReportColumn
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    '---cboTime
    cboTime.AddItem "��    ��"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "[ָ  ��]"
    cboTime.ListIndex = 2
    
    '---cboPathTime
    cboPathTime.AddItem "��    ��"
    cboPathTime.AddItem "���һ��"
    cboPathTime.AddItem "���һ��"
    cboPathTime.AddItem "���һ��"
    cboPathTime.AddItem "�������"
    cboPathTime.AddItem "���һ��"
    cboPathTime.AddItem "[ָ  ��]"
    cboPathTime.ListIndex = 2
    
    '---cboForDate
    cboForDate.AddItem "����ʱ��"
    cboForDate.AddItem "��Ժ����"
    cboForDate.AddItem "��Ժ����"
    cboForDate.ListIndex = 0
    
    '---cboTimeType
    cboTimeType.AddItem "����ʱ��"
    cboTimeType.AddItem "��Ժ����"
    cboTimeType.AddItem "��Ժ����"
    cboTimeType.ListIndex = 0
    
    '---cboYorM
    cboYorM.AddItem "����"
    cboYorM.AddItem "������"
    cboYorM.AddItem "����"
    cboYorM.ListIndex = 0
    dtpOne.Value = Format(zlDatabase.Currentdate, "yyyy-mm")
    dtpTwo.Value = Format(CDate(Format(dtpOne.Value, "yyyy-mm")) - 1, "yyyy-MM-01")
    
    '---cboTrendTime
    cboTrendTime.AddItem "����"
    cboTrendTime.AddItem "����"
    cboTrendTime.ListIndex = 0
    dtpTrendStart.Value = Format(CDate(Format(dtpOne.Value, "yyyy-mm")) - 1, "yyyy-MM-01")
    
    
    '---cboTyping ����������
    cboTyping.AddItem ""
    cboTyping.AddItem "������ͨ��"
    cboTyping.AddItem "������֢��"
    cboTyping.AddItem "����������"
    cboTyping.AddItem "����Σ����"
    cboTyping.ListIndex = 0
    
    '---
    Call RestoreWinState(Me, App.ProductName)
    Call LoadPathList
End Sub

Private Sub InitPathReportColumn()
    Dim objCol As ReportColumn

    With rptPath
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(COL_ID, "", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_ͼ��, "", 18, False)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_����, "����", 80, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 35, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(COL_����, "����", 150, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(COL_��������, "��������", 65, True)
        Set objCol = .Columns.Add(COL_���ò���, "���ò���", 55, True)
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_�����Ա�, "�����Ա�", 55, True)
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_��������, "��������", 55, True)
        Set objCol = .Columns.Add(COL_˵��, "", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_ͨ��, "", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_�����ڼ�, "�����ڼ�", 55, True)
            objCol.Alignment = xtpAlignmentCenter
        
        For Each objCol In .Columns
            objCol.Editable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ���ٴ�·��..."
            '.ShadeGroupHeadings = True
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(COL_����)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub InitPatiReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(col_��ӡ, "��ӡ", 50, True)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = img16.ListImages("Check").Index - 1
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_��ҳID, "��ҳID", 0, False)
            objCol.Visible = False
        
        Set objCol = .Columns.Add(COL_����, "����", 70, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_����, "����", 70, True)
        Set objCol = .Columns.Add(COL_�Ա�, "�Ա�", 45, True)
        Set objCol = .Columns.Add(COL_����, "����", 60, True)
        Set objCol = .Columns.Add(COL_סԺ��, "סԺ��", 62, True)
        Set objCol = .Columns.Add(COL_����, "����", 45, True)
        Set objCol = .Columns.Add(COL_סԺҽʦ, "סԺҽʦ", 62, True)
        Set objCol = .Columns.Add(COL_����, "����", 40, True)
        Set objCol = .Columns.Add(COl_����, "����", 40, True)
        Set objCol = .Columns.Add(COl_��׼סԺ��, "��׼סԺ��", 70, True)
        Set objCol = .Columns.Add(COL_��׼����, "��׼����", 80, True)
        Set objCol = .Columns.Add(COL_�汾��, "�汾��", 45, True)
        Set objCol = .Columns.Add(COL_������, "������", 55, True)
        Set objCol = .Columns.Add(COl_����ʱ��, "����ʱ��", 106, True)
        Set objCol = .Columns.Add(COL_����ʱ��, "����ʱ��", 106, True)
        
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_����ID, "����ID", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_����״̬, "����״̬", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_����ת��, "����ת��", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_������ԭ��, "������ԭ��", 200, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_�����˳�ԭ��, "�����˳�ԭ��", 200, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_��ӡ��, "��ӡ��", 55, True)
        Set objCol = .Columns.Add(COL_��ӡʱ��, "��ӡʱ��", 106, True)
        Set objCol = .Columns.Add(COL_��֧·��, "��֧·��", 106, True)
        Set objCol = .Columns.Add(COL_��֧·�����, "��֧·�����", 80, True)
        Set objCol = .Columns.Add(COL_�ϲ�·������, "�ϲ�·������", 80, True)
        Set objCol = .Columns.Add(COL_������, "������", 150, True)
        Set objCol = .Columns.Add(COL_���߰��ӡ, "���߰��ӡ", 70, True)
        Set objCol = .Columns.Add(col_����·��ID, "����·��ID", 0, False)
        For Each objCol In .Columns
            If objCol.Index <> col_��ӡ Then
                objCol.Editable = False
            End If
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û���ٴ�·��Ӧ�õĲ�������..."
            '.ShadeGroupHeadings = True
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
        .SetImageList Me.img16
        
        .GroupsOrder.Add .Columns(COL_����)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(COL_����)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_סԺ��)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub InitOperReportColumn()
    Dim objCol As ReportColumn

    With rptOper
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)��ItemIndex������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(COL_��¼ID, "��¼ID", 0, False)
            objCol.Visible = False
            
        Set objCol = .Columns.Add(COL_��������, "��������", 300, True)
        Set objCol = .Columns.Add(COL_��������, "��������", 200, True)
        Set objCol = .Columns.Add(COL_����ҽʦ, "����ҽʦ", 80, True)
        Set objCol = .Columns.Add(COL_����ҽʦ, "����ҽʦ", 80, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�в���������Ϣ..."
            '.ShadeGroupHeadings = True
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '������SelectionChanged�¼�
        .SetImageList Me.img16
    End With
End Sub

Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    
    Dim lngCount As Long
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveJpeg, "����ΪͼƬ(&J)")
            objControl.IconId = 8104
        Set objControl = .Add(xtpControlButton, conMenu_File_BatPrint, "������ӡ")
            objControl.IconId = 8128
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
            objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "չ��/�۵���(&X)"):
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "�۵�������(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "չ��������(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "�۵���ǰ��(&C)", -1, False)
                objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "չ����ǰ��(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Report, "������ͳ�Ʊ�(&V)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "�鿴·����(&P)")
            objControl.IconId = 126401202
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "�鿴�����ǼǱ�(&O)")
            objControl.IconId = 3032
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "���Ӳ�������(&I)")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "��ʾ��ͣ�õ�·����(&D)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
            objControl.BeginGroup = True
    End With

    '���˵��Ҳ�Ĳ���
    With cbsMain.ActiveMenuBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, 0, "�����Ʋ���")
        objPopup.ID = 0
        objPopup.Style = xtpButtonIconAndCaption
        objPopup.IconId = conMenu_View_Find
        objPopup.Flags = xtpFlagRightAlign
        With objPopup.CommandBar.Controls
        
            Set objControl = .Add(xtpControlButton, conMenu_View_FindName, "�����Ʋ���")
            Set objControl = .Add(xtpControlButton, conMenu_View_FindIll, "����ϲ���")
            
        End With
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFind.Hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��")
        Set objControl = .Add(xtpControlButton, conMenu_File_BatPrint, "������ӡ")
            objControl.IconId = 3903
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveJpeg, "����ΪͼƬ")
            objControl.IconId = 8104
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Report, "������ͳ�Ʊ�")
            objControl.BeginGroup = True
        
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "·����")
            objControl.IconId = 126401202
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "�����ǼǱ�")
            objControl.IconId = 3032
        Set objControl = .Add(xtpControlButton, conMenu_Tool_Archive, "����")
            objControl.BeginGroup = True
            
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '����
        .Add 0, vbKeyF3, conMenu_View_FindNext '������һ��
        
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend 'չ��������
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '�۵�������
        .Add FCONTROL, vbKeyP, conMenu_File_Print '��ӡ
        .Add FALT, vbKeyJ, conMenu_File_SaveJpeg   '����ΪͼƬ
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With

    '����һЩ�����Ĳ���������
    '-----------------------------------------------------
    With cbsMain.Options
        '.AddHiddenCommand conMenu_File_PrintSet '��ӡ����
        '.AddHiddenCommand conMenu_File_Excel '�����Excel
    End With
    
    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
    
    '��ȡ��������ģ��ı���(��������ģ���)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
End Sub

Private Function LoadPathList(Optional ByVal str���� As String, Optional ByVal str���� As String) As Boolean
'���ܣ����ݵ�ǰ���õ�������ȡ�ٴ�·��Ŀ¼����
'���������ڶ�λ
    Dim rsTmp As ADODB.Recordset
    Dim rsCode As ADODB.Recordset
    Dim strSql As String

    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objRow As ReportRow, i As Long, j As Long
    Dim lngPreID As Long, lngPreIdx As Long
    Dim strItem As String, strCodeA As String, strNameA As String, strCodeB As String, strNameB As String
    Dim colSQL As Collection
    Dim blnOK As Boolean
    
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    'SQL�в��������Ч��,ReportControl��������
    '34269 ��
    strSql = "Select ID, ����, ����, ����, ��������, ���ò���, �����Ա�, ��������, ˵��, ͨ��, �ڼ� " & vbNewLine & _
            "       From (Select a.Id, a.����, a.����, a.����, a.��������, a.���ò���, a.�����Ա�, a.��������, a.˵��, a.ͨ��, b.�ڼ�," & vbNewLine & _
            "                     Row_Number() Over(Partition By a.Id Order By b.�ڼ� Desc) As Top" & vbNewLine & _
            "              From �ٴ�·��Ŀ¼ A, ·�������ļ� B" & vbNewLine & _
            "              Where a.Id = b.·��id(+) And b.����id(+) = 1 And NVL(a.����,0)=0 And Exists" & vbNewLine & _
            "               (Select ·��id From �ٴ�·���汾 C Where a.Id = c.·��id And ����� Is Not Null" & _
            IIf(mblnShowStoped, "", " And ͣ���� Is Null") & ")) A " & vbNewLine & _
            "       Where Top = 1"
    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        'û��Ȩ��ʱ��ֻ�ܶ�ֻӦ���ڱ��Ƶ�·�����д���
        strSql = strSql & _
            " And ͨ��=2 And Not Exists(" & _
                " Select ����ID From �ٴ�·������ Where ·��ID=a.ID" & _
                " Minus Select ����ID From ������Ա Where ��ԱID=[1])"
        optThisPath.Value = True
        optAllPath.Enabled = False
        optThisPath.Enabled = False
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    
    Set rsTmp = zlDatabase.CopyNewRec(rsTmp, , , Array("��������", adVarChar, 50000, Empty, "��������", adVarChar, 50000, Empty, "��ϱ���", adVarChar, 50000, Empty, "�������", adVarChar, 50000, Empty))
    Set colSQL = New Collection
    strItem = ""
    For i = 1 To rsTmp.RecordCount
        If Len(strItem & "," & rsTmp!ID) > 4000 Then
            colSQL.Add Mid(strItem, 2)
            strItem = "," & rsTmp!ID
            blnOK = False
        Else
            strItem = strItem & "," & rsTmp!ID
            blnOK = True
        End If
        If i = rsTmp.RecordCount And blnOK Then
            colSQL.Add Mid(strItem, 2)
        End If
        rsTmp.MoveNext
    Next

    '��ȡ�������롢����-����:108062 �û����ݼ������Ƴ���4000���ַ�
    For i = 1 To colSQL.count
        strSql = "Select /*+cardinality(d,10)*/" & vbNewLine & _
                " d.Column_Value As ·��id, c.���� As ��������, c.���� As ��������, b.���� As ��ϱ���, b.���� As �������" & vbNewLine & _
                "From �ٴ�·������ A, �������Ŀ¼ B, ��������Ŀ¼ C, Table(f_Num2list([1])) D" & vbNewLine & _
                "Where d.Column_Value = a.·��id And b.Id(+) = a.���id And c.Id(+) = a.����id And a.���� = 0"
       Set rsCode = zlDatabase.OpenSQLRecord(strSql, Me.Caption, colSQL(i))
       strItem = "": strCodeA = "": strCodeB = "": strNameA = "": strNameB = ""
       For j = 1 To rsCode.RecordCount
            If rsCode!·��ID & "" <> strItem Then
                If j <> 1 Then
                    rsTmp.Filter = "ID =" & strItem
                    If Not rsTmp.EOF Then
                        rsTmp!�������� = strCodeA
                        rsTmp!��ϱ��� = strCodeB
                        rsTmp!�������� = strNameA
                        rsTmp!������� = strNameB
                    End If
                End If
                strItem = rsCode!·��ID
                strCodeA = rsCode!�������� & ""
                strCodeB = rsCode!��ϱ��� & ""
                strNameA = rsCode!�������� & ""
                strNameB = rsCode!������� & ""
            Else
                strCodeA = strCodeA & IIf(rsCode!�������� & "" <> "", "," & rsCode!��������, "")
                strCodeB = strCodeB & IIf(rsCode!��ϱ��� & "" <> "", "," & rsCode!��ϱ���, "")
                strNameA = strNameA & IIf(rsCode!�������� & "" <> "", "," & rsCode!��������, "")
                strNameB = strNameB & IIf(rsCode!������� & "" <> "", "," & rsCode!�������, "")
            End If
            rsCode.MoveNext
            If j = rsCode.RecordCount Then
                rsTmp.Filter = "ID =" & strItem
                If Not rsTmp.EOF Then
                    rsTmp!�������� = strCodeA
                    rsTmp!��ϱ��� = strCodeB
                    rsTmp!�������� = strNameA
                    rsTmp!������� = strNameB
                End If
            End If
       Next
    Next
    rsTmp.Filter = "" '�ָ���¼��
    '��¼����ѡ�еķ���
    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lngPreIdx = rptPath.SelectedRows(0).Index '���ڿ������¶�λ
            lngPreID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If
    
    rptPath.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPath.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!ID))
        Set objItem = objRecord.AddItem("")
            objItem.Icon = img16.ListImages("Path").Index - 1
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����, "<δָ������>")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��������)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!���ò���)))
        Set objItem = objRecord.AddItem(CStr(Decode(NVL(rsTmp!�����Ա�, 0), 0, "", 1, "��", 2, "Ů")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��������)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!˵��)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!ͨ��, 1)))
        Set objItem = objRecord.AddItem("" & rsTmp!�ڼ�)
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��ϱ���)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!�������)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��������)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!��������)))

        rsTmp.MoveNext
    Loop
    rptPath.Populate
    
    If rptPath.Rows.count = 0 Then
        Call ClearSubData
    Else
        If str���� <> "" And str���� <> "" Then
            For i = 0 To rptPath.Rows.count - 1
                If Not rptPath.Rows(i).GroupRow Then
                    If rptPath.Rows(i).Record(COL_����).Value = str���� _
                        And rptPath.Rows(i).Record(COL_����).Value = str���� Then
                        Set objRow = rptPath.Rows(i): Exit For
                    End If
                End If
            Next
        Else
            If lngPreID <> 0 Then
                '�ȿ��ٶ�λ
                If lngPreIdx <= rptPath.Rows.count - 1 Then
                    If Not rptPath.Rows(lngPreIdx).GroupRow Then
                        If rptPath.Rows(lngPreIdx).Record(COL_ID).Value = lngPreID Then
                            Set objRow = rptPath.Rows(lngPreIdx)
                        End If
                    End If
                End If
                '�ٽ��в���
                If objRow Is Nothing Then
                    For i = 0 To rptPath.Rows.count - 1
                        If Not rptPath.Rows(i).GroupRow Then
                            If rptPath.Rows(i).Record(COL_ID).Value = lngPreID Then
                                Set objRow = rptPath.Rows(i): Exit For
                            End If
                        End If
                    Next
                End If
            End If
            'ȡ��һ���Ƿ�����
            If objRow Is Nothing Then
                For i = 0 To rptPath.Rows.count - 1
                    If Not rptPath.Rows(i).GroupRow Then Set objRow = rptPath.Rows(i): Exit For
                Next
            End If
        End If
        
        Set rptPath.FocusedRow = objRow '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Me.stbThis.Panels(2).Text = "���� " & rptPath.Records.count & " ���ٴ�·��"
    End If
    
    
    Screen.MousePointer = 0
    LoadPathList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cbsMain_Resize
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    If Not mobjCISJob Is Nothing Then Set mobjCISJob = Nothing
End Sub

Private Sub fraFlag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fraFlag.Visible Then
       Call zlCommFun.ShowTipInfo(fraFlag.Hwnd, "����ñ������Ԥ������ӡ�������EXCEL", True)
    End If
End Sub

Private Sub fraGroupLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If vsgInfo(vsg_ԭ��).Width + X < 2000 Or vsgInfo(vsg_��Ŀ).Width - X < 2000 Then Exit Sub
        
        fraGroupLR.Left = fraGroupLR.Left + X
        vsgInfo(vsg_ԭ��).Width = vsgInfo(vsg_ԭ��).Width + X
        vsgInfo(vsg_��Ŀ).Width = vsgInfo(vsg_��Ŀ).Width - X
        vsgInfo(vsg_��Ŀ).Left = vsgInfo(vsg_��Ŀ).Left + X
        lblInfo(1).Left = lblInfo(1).Left + X
        imgFrom.Left = imgFrom.Left + X / 2
        txtFindNum.Left = txtFindNum.Left + X
        
        Me.Refresh
        Call SetFlagBySelectedTable(True, mstrFlag)
    End If
End Sub

Private Sub fraGroupUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = 1 Then
        If vsgInfo(vsg_ԭ��).Height + Y < 1740 Or vsgInfo(vsg_ԭ��).Height - Y < 1000 Then Exit Sub
        If vsgInfo(VSG_��ϸ).Height + Y < 1000 Or vsgInfo(VSG_��ϸ).Height - Y < 1740 Then Exit Sub

        fraGroupUD.Top = fraGroupUD.Top + Y
        fraGroupLR.Height = fraGroupLR.Height + Y
        vsgInfo(vsg_ԭ��).Height = vsgInfo(vsg_ԭ��).Height + Y
        vsgInfo(vsg_��Ŀ).Height = vsgInfo(vsg_��Ŀ).Height + Y
        vsgInfo(VSG_��ϸ).Top = vsgInfo(VSG_��ϸ).Top + Y
        vsgInfo(VSG_��ϸ).Height = vsgInfo(VSG_��ϸ).Height - Y
        lblInfo(2).Top = lblInfo(2).Top + Y
        imgFrom.Top = imgFrom.Top + Y

        Me.Refresh
        Call SetFlagBySelectedTable(True, mstrFlag)
    End If
End Sub

Private Sub fraLR_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If rptPath.Width + X < 2000 Or picPati.Width - X < 3000 Or tbcSub.Width - X < 3000 Then Exit Sub
        
        fraLR_S.Left = fraLR_S.Left + X
        rptPath.Width = rptPath.Width + X
        
        tbcSub.Left = tbcSub.Left + X
        tbcSub.Width = tbcSub.Width - X

        Me.Refresh
        Call SetFlagBySelectedTable(True, mstrFlag)
    End If
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next

    If Button = 1 Then
        If rptPati.Height + Y < 1740 Or rptPati.Height - Y < 1000 Then Exit Sub
        If picInfo.Height + Y < 1000 Or picInfo.Height - Y < 1740 Then Exit Sub

        fraUD.Top = fraUD.Top + Y
        rptPati.Height = rptPati.Height + Y
        picInfo.Top = picInfo.Top + Y
        picInfo.Height = picInfo.Height - Y

        Me.Refresh
    End If
End Sub

Private Sub optAllPath_Click()
    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
End Sub

Private Sub optIn_Click()
    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
End Sub

Private Sub optOut_Click()
    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
End Sub

Private Sub optState_Click(Index As Integer)
    Call rptPath_SelectionChanged
End Sub

Private Sub optThisPath_Click()
    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
End Sub

Private Sub picFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo picFilter.Hwnd, "�趨��������ִ��ˢ�¶�ȡ����(F5)"
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    fraGroup.Move 0, fraGroup.Top, picFilter.Width
End Sub

Private Sub picInfo_Resize()
    On Error Resume Next
    With Me.rptOper
        .Left = 0
        .Top = IIf(lblMerge.Visible, lblMerge.Top + lblMerge.Height, lblDiag.Top + lblDiag.Height) + 100
        .Width = picInfo.Width
        If picInfo.Height < 300 Then Exit Sub
        .Height = picInfo.Height - (IIf(lblMerge.Visible, lblMerge.Top + lblMerge.Height, lblDiag.Top + lblDiag.Height)) - 100
        lblMerge.Width = .Width
    End With
End Sub

Private Sub picPati_Resize()
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
     With Me.picFilter
        .Left = 0
        .Top = 0
        .Width = picPati.Width
    End With

    With Me.rptPati
        .Left = picFilter.Left
        .Top = picFilter.Top + picFilter.Height
        .Width = picFilter.Width
        .Height = picPati.Height - .Top - IIf(picInfo.Visible, picInfo.Height, 0)
    End With
    
    With Me.picInfo
        .Left = picFilter.Left
        .Top = picPati.Height - .Height
        If .Top < picPati.Height - picFilter.Height - 1750 Then .Top = picPati.Height - picFilter.Height - 1750
        .Width = picFilter.Width
        .Height = picPati.Height - .Top
        If .Height > picPati.Height - 1750 Then .Height = picPati.Height - 1750
    End With
    
    With Me.fraUD
        .Left = picFilter.Left
        .Top = rptPati.Top + rptPati.Height
        .Width = picFilter.Width
    End With
    
   
    
End Sub

Private Sub picReason_Resize()
    On Error Resume Next
    With chtThis
    
        .Top = 0
        .Left = 0
        .Width = picReason.Width
        .Height = picReason.Height
        lblMSG.Top = 50
        lblMSG.Left = 50
        picTable.Top = .Top
        picTable.Left = .Left
        picTable.Width = .Width
        picTable.Height = .Height
    
    End With
End Sub

Private Sub picTable_Resize()
    On Error Resume Next
    Dim lngWidth As Long, i As Long
    
    lblInfo(0).Height = 250
    lblInfo(1).Height = 250
    With vsgInfo(vsg_ԭ��)
        For i = 0 To .Cols - 1
            lngWidth = lngWidth + .ColWidth(i)
        Next
        lblInfo(0).Move 50, 50
        .Move 0, lblInfo(0).Top + lblInfo(0).Height, lngWidth + 100, IIf(vsgInfo(VSG_��ϸ).Visible Or tbcVariation.Selected.Tag = "δ����ԭ��", picTable.Height / 2, picTable.Height) - lblInfo(0).Top - lblInfo(0).Height
        fraGroupLR.Move .Width, 0, fraGroupLR.Width, .Height + lblInfo(0).Top + lblInfo(0).Height
        
        If vsgInfo(vsg_��Ŀ).Visible = False Then vsgInfo(vsg_ԭ��).Width = picTable.Width
    End With
    
    With vsgInfo(vsg_��Ŀ)
        lblInfo(1).Move fraGroupLR.Left + fraGroupLR.Width + 50, 50
        txtFindNum.Move lblInfo(1).Left + lblInfo(1).Width - 950, lblInfo(1).Top - 30
        .Move vsgInfo(vsg_ԭ��).Width + fraGroupLR.Width, vsgInfo(vsg_ԭ��).Top, picTable.Width - vsgInfo(vsg_ԭ��).Width - fraGroupLR.Width, vsgInfo(vsg_ԭ��).Height
        If Not vsgInfo(VSG_��ϸ).Visible Then Exit Sub
        fraGroupUD.Move 0, vsgInfo(vsg_ԭ��).Height + vsgInfo(vsg_ԭ��).Top, picTable.Width
    End With
    
    With vsgInfo(VSG_��ϸ)
        lblInfo(2).Move 50, fraGroupUD.Top + fraGroupUD.Height + 50
        imgFrom.Move vsgInfo(vsg_��Ŀ).Left + vsgInfo(vsg_��Ŀ).Width / 2, lblInfo(2).Top - 50
        .Move 0, lblInfo(2).Top + lblInfo(2).Height, picTable.Width, picTable.Height - lblInfo(2).Top - lblInfo(2).Height
        .ColWidth(VCol_δʹ��ԭ��) = .Width / 2.88
    End With

End Sub

Private Sub picVariation_Resize()
    On Error Resume Next
    With tbcVariation
    
        .Top = picFind.Top + picFind.Height
        .Left = 0
        .Width = picVariation.Width
        .Height = picVariation.Height - picFind.Top - cmdVariation.Height - 100 - 900
        fraGroupLine.Width = .Width
        lblZY.Top = .Top + .Height + 100
        
    End With
    
End Sub

Private Sub rptPath_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objHitTest As ReportHitTestInfo
    Dim objPopup As CommandBarPopup
        
    If Button = 2 Then
        Set objHitTest = rptPath.HitTest(X, Y)
        If objHitTest.ht = xtpHitTestReportArea And Not objHitTest.Row Is Nothing Then
            If objHitTest.Row.GroupRow Then
                Set objPopup = cbsMain.FindControl(, conMenu_View_Expend, , True)
            End If
        End If
        
        rptPath.SetFocus
        If Not objPopup Is Nothing Then objPopup.CommandBar.ShowPopup
    End If
End Sub

Private Sub LoadBranch()
'���ܣ�����·���ķ�֧·��
    Dim strSql As String, rsTmp As Recordset
    Dim lngPathID As Long
    
    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lngPathID = Val(rptPath.SelectedRows(0).Record(COL_ID).Value)
        End If
    End If
    If mlng·��ID <> lngPathID Then
        cbo·����Χ.Clear
        '---cbo·����Χ����֧·��
        If lngPathID <> 0 Then
            cbo·����Χ.AddItem "��·�������з�֧·��"
            Call cbo.SetIndex(cbo·����Χ.Hwnd, cbo·����Χ.NewIndex)
            cbo·����Χ.AddItem "��·��"
            On Error GoTo errH
            strSql = "Select ���� From �ٴ�·����֧ Where ·��ID=[1] Group by ����"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID)
            Do While Not rsTmp.EOF
                cbo·����Χ.AddItem "" & rsTmp!����
                rsTmp.MoveNext
            Loop
        End If
        If cbo·����Χ.ListCount <= 2 Then
            cbo·����Χ.Visible = False
            lbl·����Χ.Visible = False
            cmdFind.Left = lbl·����Χ.Left
        Else
            cbo·����Χ.Visible = True
            lbl·����Χ.Visible = True
            cmdFind.Left = cbo·����Χ.Left + cbo·����Χ.Width + 100
        End If
    End If
    mlng·��ID = lngPathID
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub rptPath_SelectionChanged()
    If rptPath.Tag = "1" Then Exit Sub
    Call LoadBranch
    If txtPerson.Tag = "" Then txtPerson.Text = ""
    If rptPath.SelectedRows.count = 0 Then
        Call ClearSubData
    ElseIf rptPath.SelectedRows(0).GroupRow Then
        Call ClearSubData
    Else
        With rptPath.SelectedRows(0)
            Call LoadPatiList(.Record(COL_ID).Value)
        End With
    End If
    picInfo.Visible = rptPati.Rows.count And Not rptPati.FocusedRow Is Nothing
    Call picPati_Resize
    Call tbcVariation_SelectedChanged(tbcVariation.Selected)
    If tbcSub.Selected.Tag <> "����·��" Then
        Me.stbThis.Panels(3).Visible = False
    Else
        Me.stbThis.Panels(3).Visible = True
    End If
    If Me.Visible Then
        Call SetFlagBySelectedTable(True, "RPTPATH")
    End If
End Sub

Private Sub ClearSubData()
    rptPati.Records.DeleteAll
    rptPati.Populate
    mlng����ID = 0
    mlng��ҳID = 0
    mlng����·��ID = 0
    
    Me.stbThis.Panels(3).Text = ""
End Sub

Private Sub rptPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim objControl As CommandBarControl
    
    If KeyCode = 13 Then
        Set objControl = cbsMain.FindControl(, conMenu_View_Show, True)
        If objControl.Visible And objControl.Enabled Then objControl.Execute
    End If
End Sub

Private Sub rptPati_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim objColumn As ReportColumn
    Dim i As Long
    
    '��������ͷ��ͼƬ����ѡ��ȫ��
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_��ӡ Then
                    If rptPati.Columns(col_��ӡ).Icon = img16.ListImages("Check").Index - 1 Then
                        rptPati.Columns(col_��ӡ).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPati.Records.count - 1
                            rptPati.Records(i)(col_��ӡ).Checked = True
                        Next
                    Else
                        rptPati.Columns(col_��ӡ).Icon = img16.ListImages("Check").Index - 1
                        For i = 0 To rptPati.Records.count - 1
                            rptPati.Records(i)(col_��ӡ).Checked = False
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call rptPati_KeyDown(13, 0)
End Sub

Private Sub rptPati_SelectionChanged()
    Dim rsTmp As ADODB.Recordset
    
    mlng����ID = 0
    mlng��ҳID = 0
    mlng����·��ID = 0
    
    If Me.Visible Then
        Call SetFlagBySelectedTable(True, "RPTPATI")
    End If
    If rptPati.FocusedRow Is Nothing Then Exit Sub
    If rptPati.FocusedRow.GroupRow Then Exit Sub
    cbsMain_Resize
    
    
    mlng����ID = Val(rptPati.FocusedRow.Record(COL_����ID).Value)
    mlng��ҳID = Val(rptPati.FocusedRow.Record(COL_��ҳID).Value)
    mlng����·��ID = Val(rptPati.FocusedRow.Record(col_����·��ID).Value)
    
    Set rsTmp = Get����ID(mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        lblDiagInfo.Caption = "" & rsTmp!�������
    End If
    If Val(rptPati.FocusedRow.Record(COL_�ϲ�·������).Value & "") > 0 Then
        '���غϲ�·����Ϣ
        Call SetMergeMsg(mlng����ID, mlng��ҳID)
        lblMerge.Visible = True
        lblMerges.Visible = True
    Else
        lblMerge.Visible = False
        lblMerges.Visible = False
    End If
    picInfo.Height = rptOper.Height + IIf(lblMerge.Visible, lblMerge.Height + lblMerge.Top, lblDiag.Height + lblDiag.Top) + 100
    Call LoadOperInfo(mlng����ID, mlng��ҳID)
    picInfo.Visible = rptPati.Rows.count And Not rptPati.FocusedRow Is Nothing
 
    Call picPati_Resize
    
End Sub

Private Sub SetMergeMsg(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
'���ܣ����úϲ�·����Ϣ
    Dim strSql As String, rsTmp As Recordset
    Dim lngCount As Long
    
    lblMerge.Caption = ""
    
    strSql = "Select c.����, b.�������,a.����ʱ��,a.����ʱ��,a.��ǰ����" & vbNewLine & _
            "From ���˺ϲ�·�� A, ������ϼ�¼ B, �ٴ�·��Ŀ¼ C" & vbNewLine & _
            "Where a.·��id = c.Id And a.����id = b.����id And a.��ҳid = b.��ҳid And a.������� = b.������� And a.�����Դ = b.��¼��Դ And a.����id = b.����id And" & vbNewLine & _
            "      a.����id = [1] and a.��ҳID=[2]"
    On Err GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳID)
    Do While Not rsTmp.EOF
        lngCount = lngCount + 1
        lblMerge.Caption = lblMerge.Caption & IIf(lblMerge.Caption = "", "", vbCrLf) & lngCount & "��" & rsTmp!���� & "  ��Ӧ��ϣ�" & rsTmp!������� & "  ����ʱ�䣺" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm") & "  ����ʱ�䣺" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm") & "  ��ǰ������" & rsTmp!��ǰ����
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long

    If Item.Tag = "����·��" Then
        cboForDate.ListIndex = cboTimeType.ListIndex
        cboTime.ListIndex = cboPathTime.ListIndex
        dtpTime(0).Value = dtpStart.Value
        dtpTime(1).Value = dtpEnd.Value
        mblnIsPathTo = True
        Me.stbThis.Panels(3).Visible = True
    ElseIf Item.Tag = "�������" Then
        '�ж��ϴ��Ƿ��ǲ���·��ҳ��ת������
        If mblnIsPathTo Then
            cboTimeType.ListIndex = cboForDate.ListIndex
            cboPathTime.ListIndex = cboTime.ListIndex
            dtpStart.Value = dtpTime(0).Value
            dtpEnd.Value = dtpTime(1).Value
        End If
        mblnIsPathTo = False
        '���ظſ�������ѡ�����ʾ����ԭ���ѡ�
        For i = 0 To tbcVariation.ItemCount - 1
            If i <= 7 Then
                tbcVariation.Item(i).Visible = True
            Else
                tbcVariation.Item(i).Visible = False
            End If
        Next
        '�����ϴ���������ѡ�
        If mlngVariation <= 7 Then
            tbcVariation.Item(mlngVariation).Selected = True
        Else
            tbcVariation.Item(0).Selected = True
        End If
        Me.stbThis.Panels(3).Visible = False
    ElseIf Item.Tag = "�ſ�����" Then
        If mblnIsPathTo Then
            cboTimeType.ListIndex = cboForDate.ListIndex
            cboPathTime.ListIndex = cboTime.ListIndex
            dtpStart.Value = dtpTime(0).Value
            dtpEnd.Value = dtpTime(1).Value
        End If
        mblnIsPathTo = False
        '��ʾ�ſ�������ѡ������ر���ԭ���ѡ�
        For i = 0 To tbcVariation.ItemCount - 1
            If i <= 7 Or i > 11 Then
                tbcVariation.Item(i).Visible = False
            Else
                tbcVariation.Item(i).Visible = True
            End If
        Next
        '�����ϴ���������ѡ�
        If mlngSurvey <= 7 Then
            tbcVariation.Item(8).Selected = True
        Else
            tbcVariation.Item(mlngSurvey).Selected = True
        End If
        Me.stbThis.Panels(3).Visible = False
    ElseIf Item.Tag = "���Ʒ���" Then
        If mblnIsPathTo Then
            cboTimeType.ListIndex = cboForDate.ListIndex
            cboPathTime.ListIndex = cboTime.ListIndex
            dtpStart.Value = dtpTime(0).Value
            dtpEnd.Value = dtpTime(1).Value
        End If
        mblnIsPathTo = False
        '��ʾ�ſ�������ѡ������ر���ԭ���ѡ�
        For i = 0 To tbcVariation.ItemCount - 1
            If i <= 11 Then
                tbcVariation.Item(i).Visible = False
            Else
                tbcVariation.Item(i).Visible = True
            End If
        Next
        '�����ϴ���������ѡ�
        If mlngTrend <= 11 Then
            tbcVariation.Item(12).Selected = True
        Else
            tbcVariation.Item(mlngTrend).Selected = True
        End If
        Me.stbThis.Panels(3).Visible = False
    End If
    If Me.Visible And InStr(";����·��;�������;", ";" & Item.Tag & ";") > 0 Then
        If rptPati.Records.count > 0 And Item.Tag = "����·��" Then
        
            Call SetFlagBySelectedTable(True, "RPTPATI")
        ElseIf Item.Tag = "�������" Then
            Call SetFlagBySelectedTable(True, "VSGINFO_0")
        End If
    End If
    
End Sub

Private Sub Setδ����ԭ��(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        lblMSG.Visible = False
        chtThis.Visible = True
        cbo����(1).Visible = False
        lbl����(1).Visible = False
        lblZY.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypePie
        .ChartArea.Border.Width = 4
        'δ����ԭ��
        .Header.Text = "δ����ԭ��ֲ�ͼ"
        .Header.Interior.ForegroundColor = vbBlack
        '.ChartArea.Pie.StartAngle = 90
        strSql = "Select b.����, b.����, Count(1) As δ��������, 100 * Round(Count(1) / Sum(Count(1)) Over(), 4) ����" & vbNewLine & _
                "From �����ٴ�·�� A, ���쳣��ԭ�� B,������ҳ D" & vbNewLine & _
                "Where a.δ����ԭ�� = b.���� And d.����id=a.����id and a.��ҳid=d.��ҳid And b.���� = 0"
        strSql = strSql & IIf(optAllPath.Value, "", " And a.·��id=[1]")
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.����, b.����"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.NumSeries = rsTmp.RecordCount
        If .ChartGroups.Item(1).Data.NumSeries <> 0 Then .ChartGroups.Item(1).Data.NumPoints(1) = 1
        i = 1
        Do Until rsTmp.EOF
            .ChartGroups.Item(1).SeriesLabels.Add (rsTmp!���� & Space(2) & "��" & rsTmp!δ�������� & "��(" & Val(rsTmp!���� & "") & "%)")
            .ChartGroups.Item(1).Data.Y(i, 1) = Val(rsTmp!���� & "")
            rsTmp.MoveNext
            i = i + 1
            
        Loop
        'ע����Ϣ
        lblZY.Caption = "ע����ͼ�ļ���������һ�����˵�һ��סԺ��ÿ��סԺ����������Ϊһ�Σ���" & vbCrLf & _
                        "���У�û��ʹ�ù���δ����ԭ����ʾ������"
        If rsTmp.RecordCount = 0 Then
            lblMSG.Caption = "����ָ����ʱ�䷶Χ��δ����δ����Ĳ��ˡ�"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        '�����ϴ������ͼ
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Set�����˳�����(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        cbo����(1).Visible = False
        lbl����(1).Visible = False
        lblMSG.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        '�����˳�����
        .ChartGroups.Item(1).ChartType = oc2dTypePie
        .ChartArea.Border.Width = 4
        .Header.Text = "�����˳�ԭ��ֲ�ͼ"
        .Header.Interior.ForegroundColor = vbBlack
        '.ChartArea.Pie.StartAngle = 90
        strSql = "Select g.����, g.����, Count(1) As �����˳�����, 100 * Round(Count(1) / Sum(Count(1)) Over(), 4) ""����""" & vbNewLine & _
                "From �����ٴ�·�� A, ����·������ B,����·������ C ," & IIf(strDateTmp = "A.����ʱ��", "", "������ҳ D,") & " ���쳣��ԭ�� G" & vbNewLine & _
                "Where " & IIf(strDateTmp = "A.����ʱ��", "", "a.����id = d.����id And a.��ҳid = d.��ҳid And ") & " b.·����¼id = a.Id And b.���� = a.��ǰ���� And  " & vbNewLine & _
                " b.·����¼Id=C.·����¼ID(+) And b.�׶�ID=C.�׶�ID(+) and b.����=c.����(+) " & vbNewLine & _
                " And g.���� = NVl(c.����ԭ��,b.����ԭ��) And a.״̬ = 3 And G.����=2"
                '������·����������� ������·�����족������������Ϊ�˼��ݲ�ѯ��ǰ���ݣ�����·������Ϊ 10.34.0������
        '�Ƿ�ǰ·��ͳ��
        strSql = strSql & IIf(optAllPath.Value, "", " And a.·��id=[1]")
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By g.����, g.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.NumSeries = rsTmp.RecordCount
        If .ChartGroups.Item(1).Data.NumSeries <> 0 Then .ChartGroups.Item(1).Data.NumPoints(1) = 1
        i = 1
        Do Until rsTmp.EOF
            .ChartGroups.Item(1).SeriesLabels.Add (rsTmp!���� & Space(2) & "��" & rsTmp!�����˳����� & "��(" & Val(rsTmp!���� & "") & "%)")
            .ChartGroups.Item(1).Data.Y(i, 1) = Val(rsTmp!���� & "")
            rsTmp.MoveNext
            i = i + 1
        Loop
        'ע����Ϣ
        lblZY.Caption = "ע����ͼ�ļ���������һ�����˵�һ��סԺ��ÿ��סԺ����������Ϊһ�Σ���" & vbCrLf & _
                        "���У�û��ʹ�ù��ı����˳�ԭ����ʾ������"
        If rsTmp.RecordCount = 0 Then
            lblMSG.Caption = "����ָ����ʱ�䷶Χ��δ���ֱ����˳��Ĳ��ˡ�"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        '�����ϴ������ͼ
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub setʱ��������(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        cbo����(1).Visible = False
        lbl����(1).Visible = False
        lblMSG.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '����ͼ�δ�С
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 40
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 80
        .ChartArea.Border.Width = 4
        .Header.Text = "ʱ���������ֲ�ͼ"
        .Header.Interior.ForegroundColor = vbBlack
        '���������ɫ������
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartGroups.Item(1).Data.NumPoints(1) = 5
        .ChartArea.Bar.ClusterWidth = 35
        '������Ӱ
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3DЧ��
        .ChartArea.View3D.depth = 10
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 25
        .ChartGroups.Item(1).SeriesLabels.Add ("����")
        '��������
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 45
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        '��������
        .ChartGroups.Item(1).PointLabels.Add ("����")
        .ChartGroups.Item(1).PointLabels.Add ("�׶���ǰ")
        .ChartGroups.Item(1).PointLabels.Add ("�׶��Ӻ�")
        .ChartGroups.Item(1).PointLabels.Add ("���ڱ�׼סԺ��")
        .ChartGroups.Item(1).PointLabels.Add ("������׼סԺ��")
        strSql = "Select ����,����, 100 * Round(���� / Decode(Sum(����) Over(), 0, 1,Sum(����) Over()), 4) ���� From (With Test As" & vbNewLine & _
                " (Select Distinct b.·����¼id, Decode(b.ʱ�����, 0, '����', 1, '�׶���ǰ',2,'�׶���ǰ', -1, '�׶��Ӻ�') As ����" & vbNewLine & _
                "  From �����ٴ�·�� A, ����·������ B, ������ҳ D" & vbNewLine & _
                "  Where a.����id = d.����id And a.��ҳid = d.��ҳid And b.ʱ����� <> 0 And a.id=b.·����¼ID"
        '�Ƿ�ǰ·��ͳ��
        strSql = strSql & IIf(optAllPath.Value, "", " And a.·��id=[1]")
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & "Group By b.·����¼id, b.ʱ�����" & vbNewLine & _
                " Union All " & vbNewLine & _
                " Select ·����¼id, ����" & vbNewLine & _
                " From (Select a.Id As ·����¼id," & vbNewLine & _
                "              Decode(Sign(a.��ǰ���� -" & vbNewLine & _
                "                           Nvl(Substr(c.��׼סԺ��, 0, Instr(c.��׼סԺ��, '-') - 1), Substr(c.��׼סԺ��, Instr(c.��׼סԺ��, '-') + 1))), 0," & vbNewLine & _
                "                      '����', -1, '���ڱ�׼סԺ��', 1," & vbNewLine & _
                "                      Decode(Sign(a.��ǰ���� - Substr(c.��׼סԺ��, Instr(c.��׼סԺ��, '-') + 1)), 1, '������׼סԺ��', '����')) As ����" & vbNewLine & _
                "       From �����ٴ�·�� A, �ٴ�·���汾 C, ������ҳ D" & vbNewLine & _
                "       Where a.·��id = c.·��id And a.�汾�� = c.�汾�� And a.����id = d.����id And a.��ҳid = d.��ҳid And a.����ʱ�� Is Not Null And a.��ǰ���� is not null"
        '�Ƿ�ǰ·��ͳ��
        strSql = strSql & IIf(optAllPath.Value, "", " And a.·��id=[1]")
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & ") Where ���� <> '����')" & vbNewLine & _
                "Select '����' As ����, Count(1) As ����" & vbNewLine & _
                "From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
                "Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.��ǰ���� is not null And Not Exists (Select 1 From Test Where a.Id = Test.·����¼id)"
        '�Ƿ�ǰ·��ͳ��
        strSql = strSql & IIf(optAllPath.Value, "", " And a.·��id=[1]")
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & "Union All" & vbNewLine & _
                "Select ����, Count(1) As ���� From Test Group By ����) group by ����,����"
                
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.Y(1, 1) = 0
        .ChartGroups.Item(1).Data.Y(1, 2) = 0
        .ChartGroups.Item(1).Data.Y(1, 3) = 0
        .ChartGroups.Item(1).Data.Y(1, 4) = 0
        .ChartGroups.Item(1).Data.Y(1, 5) = 0
        If rsTmp.RecordCount = 1 And Val(rsTmp!���� & "") = 0 Then
            lblMSG.Caption = "����ָ����ʱ�䷶Χ��δ����ʱ�����Ĳ��ˡ�"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        Do Until rsTmp.EOF
            
            Select Case rsTmp!���� & ""
            
                Case "����"
                    .ChartGroups.Item(1).Data.Y(1, 1) = Val(rsTmp!���� & "")
                    i = 1
                Case "�׶���ǰ"
                    .ChartGroups.Item(1).Data.Y(1, 2) = Val(rsTmp!���� & "")
                    i = 2
                Case "�׶��Ӻ�"
                    .ChartGroups.Item(1).Data.Y(1, 3) = Val(rsTmp!���� & "")
                    i = 3
                Case "���ڱ�׼סԺ��"
                    .ChartGroups.Item(1).Data.Y(1, 4) = Val(rsTmp!���� & "")
                    i = 4
                Case "������׼סԺ��"
                    .ChartGroups.Item(1).Data.Y(1, 5) = Val(rsTmp!���� & "")
                    i = 5
                
            End Select
            '����ÿ�������ı�ǩ
            Set chtLabel = .ChartLabels.Add()
            chtLabel.Offset = 15
            chtLabel.Border.Type = oc2dBorderShadow
            chtLabel.Border.Width = 2
            chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
            chtLabel.AttachMethod = oc2dAttachDataIndex
            chtLabel.AttachDataIndex.Point = i
            chtLabel.IsConnected = True
            chtLabel.Anchor = oc2dAnchorNorthEast
            chtLabel.Name = rsTmp!���� & ""
            chtLabel.Text = "��" & rsTmp!���� & "��(" & Val(rsTmp!���� & "") & "%)"
            chtLabel.Font.Size = 8
            rsTmp.MoveNext
        Loop
        For i = 1 To 5
            If .ChartGroups.Item(1).Data.Y(1, i) = 0 Then
                 'û�������Ľ���ǩ=0����
                Set chtLabel = .ChartLabels.Add()
                chtLabel.Offset = 15
                chtLabel.Border.Type = oc2dBorderShadow
                chtLabel.Border.Width = 2
                chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
                chtLabel.AttachMethod = oc2dAttachDataIndex
                chtLabel.AttachDataIndex.Point = i
                chtLabel.IsConnected = True
                chtLabel.Anchor = oc2dAnchorNorthEast
                chtLabel.Name = .ChartGroups.Item(1).PointLabels(i).Text
                chtLabel.Text = "��0��(0%)"
                chtLabel.Font.Size = 8
            End If
            If i <> 1 Then
                If .ChartGroups.Item(1).Data.Y(1, i) > .ChartArea.Axes.Item(2).Max Then
                    .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 5
                End If
            Else
                .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 5
            End If
        Next
        'ע����Ϣ
        lblZY.Caption = "��һ�����˵�һ��·�������У�" & vbCrLf & _
                        "������ָδ��������4�ֱ���������" & vbCrLf & _
                        "�׶���ǰ\�׶��Ӻ�һ��������һ��סԺ��·��������ֻҪ�����˾����ҽ���һ�Ρ�(�����������˷ֱ���һ��)" & vbCrLf & _
                        "���ڱ�׼סԺ��\������׼סԺ��:һ������һ��סԺ�Ѿ�������·�����������ڻ���ڱ�׼סԺ����һ�Ρ�"
        '�����ϴ������ͼ
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Setδ����ԭ��(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    With chtThis
        cbo����(1).Visible = False
        lbl����(1).Visible = False
        chtThis.Visible = False
        picTable.Visible = True
        lblZY.Visible = True
        vsgInfo(vsg_��Ŀ).Visible = True
        strHead = "����,1500,1;ԭ��,2000,1;����,800,7"
        Call InitTable(vsgInfo(vsg_ԭ��), strHead)
        
        strHead = "�׶�����,1500,1;��Ŀ����,5000,1;����,800,7"
        Call InitTable(vsgInfo(vsg_��Ŀ), strHead)
        '��ͬ�ϲ���Ԫ��
        vsgInfo(vsg_��Ŀ).MergeCells = flexMergeRestrictColumns
        vsgInfo(vsg_��Ŀ).MergeCol(VCol_�׶�) = True
        vsgInfo(vsg_ԭ��).Rows = 1
        vsgInfo(vsg_��Ŀ).Rows = 1
        fraGroupLR.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        vsgInfo(vsg_��Ŀ).TextMatrix(0, VCol_����) = "��Ŀ����"
        lblInfo(0).Caption = "δ����ԭ����ܱ�"
        lblInfo(1).Caption = "δ������Ŀ���ܱ�(˫���鿴��Ӧҽ��)"
        txtFindNum.Visible = False
        'ԭ���
        strSql = "Select b.����,e.���� as �ϼ�����, b.����, Count(1) As ����" & vbNewLine & _
                " From �����ٴ�·�� A, ���쳣��ԭ�� B, ����·��ִ�� C, ������ҳ D,���쳣��ԭ�� E" & vbNewLine & _
                " Where c.����ԭ�� = b.���� And c.·����¼id = a.Id And d.����id = a.����id And a.��ҳid = d.��ҳid and e.����=b.�ϼ� And b.���� = 1 And c.��Ŀid Is Not Null"
        strSql = strSql & " And a.·��id=[1]"
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.����, b.����,e.���� order by ���� desc"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        
        With vsgInfo(vsg_ԭ��)
        For i = 1 To rsTmp.RecordCount
                .AddItem ""
                .RowData(i) = rsTmp!���� & ""
                .TextMatrix(i, VCol_����) = rsTmp!�ϼ�����
                .TextMatrix(i, VCol_ԭ��) = rsTmp!���� & ""
                .TextMatrix(i, VCol_ԭ������) = rsTmp!���� & ""
                
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(vsg_ԭ��).Rows = 1 Then vsgInfo(vsg_ԭ��).Rows = 2
        'δ����·����Ŀ��
        strSql = "Select b.Id, b.��Ŀ����, b.�׶�id, e.���� As �׶�����, Count(1) As ����, Nvl(f.���, e.���) ���" & vbNewLine & _
                " From ����·��ִ�� C, �ٴ�·����Ŀ B, �����ٴ�·�� A, ������ҳ D, �ٴ�·���׶� E, �ٴ�·���׶� F" & vbNewLine & _
                " Where c.��Ŀid = b.Id And c.·����¼id = a.Id And d.����id = a.����id And a.��ҳid = d.��ҳid And e.Id = b.�׶�id And" & vbNewLine & _
                "      c.��Ŀid Is Not Null And c.����ԭ�� Is Not Null And e.��id = f.Id(+)"
        strSql = strSql & " And a.·��id=[1]"
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.Id, b.��Ŀ����, b.�׶�id, e.����,Nvl(f.���, e.���) Order By ���,���� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        
        With vsgInfo(vsg_��Ŀ)
        For i = 1 To rsTmp.RecordCount
                .AddItem ""
                .RowData(i) = rsTmp!ID & ""
                .TextMatrix(i, VCol_�׶�) = rsTmp!�׶����� & ""
                .Cell(flexcpData, i, VCol_�׶�) = Val(rsTmp!�׶�id & "")
                .TextMatrix(i, VCol_����) = rsTmp!��Ŀ���� & ""
                .TextMatrix(i, VCol_��Ŀ����) = rsTmp!���� & ""
                
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(vsg_��Ŀ).Rows = 1 Then vsgInfo(vsg_��Ŀ).Rows = 2
        'ע����Ϣ
        lblZY.Caption = "ע����ҳ����Ϊ��ͳ�Ƶ��������У��������ɵ�����û������·����Ŀ�ı�����Ϣ��"
        '�����ϴ������ͼ
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Set·������Ŀ(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    With chtThis
        cbo����(1).Visible = False
        lbl����(1).Visible = False
        chtThis.Visible = False
        picTable.Visible = True
        lblZY.Visible = True
        vsgInfo(vsg_��Ŀ).Visible = True
        strHead = "����,1300,1;ԭ��,1800,1;����,800,7"
        Call InitTable(vsgInfo(vsg_ԭ��), strHead)
        
        strHead = "�׶�����,1300,1;��Ŀ����,3050,1;����,800,7"
        Call InitTable(vsgInfo(vsg_��Ŀ), strHead)
        '��ͬ�ϲ���Ԫ��
        vsgInfo(vsg_��Ŀ).MergeCells = flexMergeRestrictColumns
        vsgInfo(vsg_��Ŀ).MergeCol(VCol_�׶�) = True
        vsgInfo(VSG_��ϸ).Visible = False
        fraGroupUD.Visible = False
        fraGroupLR.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        imgFrom.Visible = False
        txtFindNum.Visible = False
        vsgInfo(vsg_ԭ��).Rows = 1
        vsgInfo(vsg_��Ŀ).TextMatrix(0, VCol_����) = "ҽ������"
        lblInfo(0).Caption = "·������Ŀ����ԭ����ܱ�"
        lblInfo(1).Caption = "·������Ŀ��Ӧҽ�����ܱ�   ��ʾÿ���׶�ǰ     ��ҽ��"
        txtFindNum.Visible = True
        txtFindNum.Tag = "OK"
        'ԭ���
        strSql = "Select b.����, b.����,e.���� as �ϼ�����, Count(1) As ����" & vbNewLine & _
                " From �����ٴ�·�� A, ���쳣��ԭ�� B, ����·��ִ�� C, ������ҳ D,���쳣��ԭ�� E" & vbNewLine & _
                " Where c.����ԭ�� = b.���� And c.·����¼id = a.Id And d.����id = a.����id And a.��ҳid = d.��ҳid And e.����=b.�ϼ� And b.���� = 1 And c.��Ŀid is Null"
        strSql = strSql & " And a.·��id=[1]"
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.����, b.����,e.���� order by ���� desc"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        
        With vsgInfo(vsg_ԭ��)
            
        For i = 1 To rsTmp.RecordCount
                .AddItem ""
                .RowData(i) = rsTmp!���� & ""
                .TextMatrix(i, VCol_����) = rsTmp!�ϼ�����
                .TextMatrix(i, VCol_ԭ��) = rsTmp!���� & ""
                .TextMatrix(i, VCol_ԭ������) = rsTmp!���� & ""
                
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(vsg_ԭ��).Rows = 1 Then vsgInfo(vsg_ԭ��).Rows = 2
        '���·������Ŀ��Ӧ��ҽ��
        Call GetPathOutAdvice
        If vsgInfo(vsg_��Ŀ).Rows = 1 Then vsgInfo(vsg_��Ŀ).Rows = 2
        'ע����Ϣ
        lblZY.Caption = "ע����ҳ����Ϊ��ͳ�Ƶ��������У������׶���ӵ�·������Ŀ�ı�����Ϣ��"
        '�����ϴ������ͼ
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Set·��������(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        cbo����(1).Visible = False
        lbl����(1).Visible = False
        '�����ϴ������ͼ
        mlngSurvey = tbcVariation.Selected.Index
        '·��������
        lblMSG.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '����ͼ�δ�С
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 40
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 80
        .ChartArea.Border.Width = 4
        .Header.Text = "·���������ֲ�ͼ"
        .Header.Interior.ForegroundColor = vbBlack
        '���������ɫ������
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartGroups.Item(1).Data.NumPoints(1) = 5
        .ChartArea.Bar.ClusterWidth = 30
        '������Ӱ
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3DЧ��
        .ChartArea.View3D.depth = 10
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 25
        .ChartGroups.Item(1).SeriesLabels.Add ("����")
        '��������
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 45
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        '��������
        .ChartGroups.Item(1).PointLabels.Add ("δ����")
        .ChartGroups.Item(1).PointLabels.Add ("����ִ��")
        .ChartGroups.Item(1).PointLabels.Add ("�������")
        .ChartGroups.Item(1).PointLabels.Add ("�������")
        .ChartGroups.Item(1).PointLabels.Add ("�����˳�")

        
        strSql = "Select ������, δ��������, Round(δ�������� / ������, 4) * 100 As δ�������, ����ִ������, Round(����ִ������ / ������, 4) * 100 As ����ִ�б���, �����������," & vbNewLine & _
                "       Round(����������� / ������, 4) * 100 As ������ɱ���, �����˳�����, Round(�����˳����� / ������, 4) * 100 As �����˳�����, �����������," & vbNewLine & _
                "       Round(����������� / ������, 4) * 100 As ������ɱ���" & vbNewLine & _
                "From (Select Count(1) As ������, Sum(Decode(a.״̬, 0, 1, 0)) As δ��������, Sum(Decode(a.״̬, 1, 1, 0)) As ����ִ������," & vbNewLine & _
                "              Sum(Decode(a.״̬, 2, 1, 0)) As �����������, Sum(Decode(a.״̬, 3, 1, 0)) As �����˳�����," & vbNewLine & _
                "              Sum(Decode(a.״̬, 100, 1, 0)) As �����������" & vbNewLine & _
                "       From (Select a.Id, Decode(a.״̬, 2, Decode(Sign(Sum(Decode(p.�������, -1, 1, 0))), 0, 2, 1, 100), a.״̬) As ״̬" & vbNewLine & _
                "              From �����ٴ�·�� A, ������ҳ D, ����·������ P" & vbNewLine & _
                "              Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.Id = p.·����¼id(+) "
        '�Ƿ�ǰ·��ͳ��
        strSql = strSql & IIf(optAllPath.Value, "", " And a.·��id=[1]")
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
        "              Group By a.Id, a.״̬) A)"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.Y(1, 1) = 0
        .ChartGroups.Item(1).Data.Y(1, 2) = 0
        .ChartGroups.Item(1).Data.Y(1, 3) = 0
        .ChartGroups.Item(1).Data.Y(1, 4) = 0
        .ChartGroups.Item(1).Data.Y(1, 5) = 0
        If rsTmp.RecordCount = 1 And Val(rsTmp!������ & "") = 0 Then
            lblMSG.Caption = "����ָ����ʱ�䷶Χ��δ�����ٴ�·�����ˡ�"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        
        If Not rsTmp.EOF Then
            For i = 1 To 5
                '����ÿ�������ı�ǩ
                Set chtLabel = .ChartLabels.Add()
                chtLabel.Offset = 15
                chtLabel.Border.Type = oc2dBorderShadow
                chtLabel.Border.Width = 2
                chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
                chtLabel.AttachMethod = oc2dAttachDataIndex
                chtLabel.AttachDataIndex.Point = i
                chtLabel.IsConnected = True
                chtLabel.Anchor = oc2dAnchorNorthEast
                chtLabel.Font.Size = 8
                If i = 1 Then
                    .ChartGroups.Item(1).Data.Y(1, 1) = Val(rsTmp!δ�������� & "")
                    chtLabel.Name = "δ��������"
                    chtLabel.Text = "��" & rsTmp!δ�������� & "��(" & Val(rsTmp!δ������� & "") & "%)"
                ElseIf i = 2 Then
                    .ChartGroups.Item(1).Data.Y(1, 2) = Val(rsTmp!����ִ������ & "")
                    chtLabel.Name = "����ִ������"
                    chtLabel.Text = "��" & rsTmp!����ִ������ & "��(" & Val(rsTmp!����ִ�б��� & "") & "%)"
                ElseIf i = 3 Then
                    .ChartGroups.Item(1).Data.Y(1, 3) = Val(rsTmp!����������� & "")
                    chtLabel.Name = "�����������"
                    chtLabel.Text = "��" & rsTmp!����������� & "��(" & Val(rsTmp!������ɱ��� & "") & "%)"
                ElseIf i = 4 Then
                    .ChartGroups.Item(1).Data.Y(1, 4) = Val(rsTmp!����������� & "")
                    chtLabel.Name = "�����������"
                    chtLabel.Text = "��" & rsTmp!����������� & "��(" & Val(rsTmp!������ɱ��� & "") & "%)"
                Else
                    .ChartGroups.Item(1).Data.Y(1, 5) = Val(rsTmp!�����˳����� & "")
                    chtLabel.Name = "�����˳�����"
                    chtLabel.Text = "��" & rsTmp!�����˳����� & "��(" & Val(rsTmp!�����˳����� & "") & "%)"
                End If
                If i <> 1 Then
                    If .ChartGroups.Item(1).Data.Y(1, i) > .ChartArea.Axes.Item(2).Max Then
                        .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 5
                    End If
                Else
                    .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 5
                End If
            Next
            
        End If
        'ע����Ϣ
        lblZY.Caption = "ע����ͼ�ļ���������һ�����˵�һ��סԺ��ÿ��סԺ����·����Ϊһ�Σ���" & vbCrLf & _
                        "���У�δ����--����ʱ�����ϵ��������Ĳ���       ����ִ��--����·���еĲ���      �������--��������·���Ĳ��ˡ�" & vbCrLf & _
                        "      �������--������������·���Ĳ���         �����˳�--���������û������·���Ĳ��ˡ�"

    End With
End Sub

Private Sub Set�׶�ƽ������(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim lng�ϲ�·��ID As Long
    Dim lngEdition As Long
    
    With chtThis
        cbo����(1).Visible = True
        lbl����(1).Visible = True
        lng�ϲ�·��ID = cbo����(1).ItemData(cbo����(1).ListIndex)
        '�����ϴ������ͼ
        mlngSurvey = tbcVariation.Selected.Index
        '·��������
        lblMSG.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        lblPathEdition.Visible = True
        cboPathEdition.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '����ͼ�δ�С
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 60
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 100
        .ChartArea.Border.Width = 4
        .Header.Text = "�׶�ƽ�����÷ֲ�ͼ"
        .Header.Interior.ForegroundColor = vbBlack
        '���������ɫ������
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartArea.Bar.ClusterWidth = 15
        '������Ӱ
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3DЧ��
        .ChartArea.View3D.depth = 5
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 15
        
        '��������
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 45
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        If Not mblnIsEdition And (mlngOldPathID <> lngPathID Or mdateOldStart <> CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00")) Or _
                                    mdateOldEnd <> CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")) Or mstrDateType <> cboTimeType.Text) Then
            strSql = "Select Distinct �汾��" & vbNewLine & _
                    " From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
                    " Where d.����id = a.����id And a.��ҳid = d.��ҳid"
                    
            strSql = strSql & " And a.·��id=[1] "
            'ʱ�䷶Χ
            strSql = strSql & " And " & strDateTmp & _
                    " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
            strSql = strSql & " Order By �汾�� Desc"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                        Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
            cboPathEdition.Clear
            Do Until rsTmp.EOF
            
                cboPathEdition.AddItem "�� " & rsTmp!�汾�� & " ��"
                cboPathEdition.ItemData(cboPathEdition.NewIndex) = Val(rsTmp!�汾�� & "")
                rsTmp.MoveNext
            Loop
            If cboPathEdition.ListCount > 0 Then Call cbo.SetIndex(cboPathEdition.Hwnd, 0)
            
        End If
        mblnIsEdition = False
        If lng�ϲ�·��ID > 0 Then
            strSql = "Select /*+ rule*/h.���� �׶�����, a.�汾��, Avg(a.����) As ƽ������ ,Nvl(g.���, h.���) ���" & vbNewLine & _
                    "From (Select f.����id, b.�ϲ�·���׶�id as �׶�id, a.�汾��, Sum(f.ʵ�ս��) As ����" & vbNewLine & _
                    "       From ����·��ִ�� B, ���˺ϲ�·�� A, ����·��ҽ�� C, סԺ���ü�¼ F, ������ҳ D, �����ٴ�·�� G" & vbNewLine & _
                    "       Where b.�ϲ�·����¼id = a.Id and g.id=a.��Ҫ·����¼ID And c.·��ִ��id = b.Id And c.����ҽ��id = f.ҽ����� And d.����id = a.����id And a.��ҳid = d.��ҳid And" & vbNewLine & _
                    "             f.��¼״̬ <> 0 And g.״̬=2 "
            '��ǰ·��ͳ��
            strSql = strSql & " And a.·��id=[1] And g.�汾��=[4]"
        Else
            strSql = "Select /*+ rule*/h.���� �׶�����, a.�汾��, Avg(a.����) As ƽ������ ,Nvl(g.���, h.���) ���" & vbNewLine & _
                    "From (Select f.����id, b.�׶�id, a.�汾��, Sum(f.ʵ�ս��) As ����" & vbNewLine & _
                    "       From ����·��ִ�� B, �����ٴ�·�� A, ����·��ҽ�� C, סԺ���ü�¼ F, ������ҳ D" & vbNewLine & _
                    "       Where b.·����¼id = a.Id And c.·��ִ��id = b.Id And c.����ҽ��id = f.ҽ����� And d.����id = a.����id And a.��ҳid = d.��ҳid And" & vbNewLine & _
                    "             f.��¼״̬ <> 0 And a.״̬=2 " & _
                    IIf(lng�ϲ�·��ID = -1, " And b.�ϲ�·����¼ID is null", "")
            '��ǰ·��ͳ��
            strSql = strSql & " And a.·��id=[1] And a.�汾��=[4]"
        End If
        
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
                
        strSql = strSql & "Group By f.����id, a.�汾��, b." & IIf(lng�ϲ�·��ID > 0, "�ϲ�·���׶�ID", "�׶�id") & vbNewLine & _
                "       Having Sum(f.ʵ�ս��) <> 0) A, �ٴ�·���׶� H ,�ٴ�·���׶� G" & vbNewLine & _
                "Where h.Id = a.�׶�id and h.��id = g.Id(+)" & vbNewLine & _
                "Group By nvl(g.id,h.Id), h.����, a.�汾��,Nvl(g.���, h.���) Order By ���"

        If cboPathEdition.ListIndex = -1 Or cboPathEdition.ListCount = 0 Then
            lngEdition = 0
        Else
            lngEdition = Val(cboPathEdition.ItemData(cboPathEdition.ListIndex))
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng�ϲ�·��ID > 0, lng�ϲ�·��ID, lngPathID), _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"), lngEdition)

        If rsTmp.RecordCount = 0 Then
            lblMSG.Caption = "����ָ����ʱ�䷶Χ��δ���ֲ������õ�·�����ˡ�"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        Else
            .ChartGroups.Item(1).Data.NumPoints(1) = rsTmp.RecordCount
        End If
        i = 1
        Do While Not rsTmp.EOF
            '��������
            .ChartGroups.Item(1).PointLabels.Add (Mid(rsTmp!�׶����� & "", 1, 10) & IIf(Len(rsTmp!�׶����� & "") > 10, "...", ""))
            .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!ƽ������ & "")
                
            '����ÿ�������ı�ǩ
            Set chtLabel = .ChartLabels.Add()
            chtLabel.Offset = 15
            chtLabel.Border.Type = oc2dBorderShadow
            chtLabel.Border.Width = 2
            chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
            chtLabel.AttachMethod = oc2dAttachDataIndex
            chtLabel.AttachDataIndex.Point = i
            chtLabel.IsConnected = True
            chtLabel.Anchor = oc2dAnchorNorthEast
            chtLabel.Name = rsTmp!�׶����� & ""
            chtLabel.Text = Format(rsTmp!ƽ������, "##.00") & "Ԫ"
            chtLabel.Font.Size = 8
            
            If i <> 1 Then
                If .ChartGroups.Item(1).Data.Y(1, i) > .ChartArea.Axes.Item(2).Max Then
                    .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 10
                End If
            Else
                .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 10
            End If
            i = i + 1
            rsTmp.MoveNext
        Loop
        'ע����Ϣ
        lblZY.Caption = "ע����ͼ�ǵ�ǰ·������ѡ���·���汾��Ӧ�Ľ׶��˾�����ͼ��" & vbCrLf & _
                        "���У�1����ͼͳ�ƵĶ������Ѿ��������굱ǰ·���Ĳ��ˡ�" & vbCrLf & _
                        "       2��Ĭ����ʾ���°汾�Ľ׶��˾����ã���ѡ��鿴����汾������Ϣ��" & vbCrLf & _
                        "       3����ѡ��İ汾Ϊ��ǰѡ���ʱ���������ù���·���汾��"
        mlngOldPathID = lngPathID
        mdateOldStart = CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"))
        mdateOldEnd = CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        mstrDateType = cboTimeType.Text
    
    End With
End Sub

Private Sub SetסԺ�շֲ�ͼ(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim lng�ϲ�·��ID As Long
    
    With chtThis
        '�����ϴ������ͼ
        mlngSurvey = tbcVariation.Selected.Index
        '·��������
        lblMSG.Visible = False
        chtThis.Visible = True
        cbo����(1).Visible = True
        lbl����(1).Visible = True
        lng�ϲ�·��ID = cbo����(1).ItemData(cbo����(1).ListIndex)
        lblZY.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '����ͼ�δ�С
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 60
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 50
        .ChartArea.Border.Width = 4
        .Header.Interior.ForegroundColor = vbBlack
        '���������ɫ������
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartArea.Bar.ClusterWidth = 15
        '������Ӱ
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3DЧ��
        .ChartArea.View3D.depth = 5
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 15
        
        '��������
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        .ChartGroups.Item(1).SeriesLabels.Add ("����")
        
        If lng�ϲ�·��ID > 0 Then
            strSql = "Select ��ǰ����, ����, ��׼סԺ��, Round(���� / Sum(����) Over(), 4) * 100 As ����" & vbNewLine & _
                    "From (Select a.��ǰ����, c.��׼סԺ��, Count(1) As ����" & vbNewLine & _
                    "       From ���˺ϲ�·�� A, ������ҳ D, �ٴ�·��Ŀ¼ B, �ٴ�·���汾 C,�����ٴ�·�� E" & vbNewLine & _
                    "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.��Ҫ·����¼ID=e.ID And b.Id = a.·��id And b.Id = c.·��id And b.���°汾 = c.�汾�� And e.״̬ = 2 And a.��ǰ���� Is Not Null"
        Else
            strSql = "Select ��ǰ����, ����, ��׼סԺ��, Round(���� / Sum(����) Over(), 4) * 100 As ����" & vbNewLine & _
                    "From (Select a.��ǰ����, c.��׼סԺ��, Count(1) As ����" & vbNewLine & _
                    "       From �����ٴ�·�� A, ������ҳ D, �ٴ�·��Ŀ¼ B, �ٴ�·���汾 C" & vbNewLine & _
                    "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And b.Id = a.·��id And b.Id = c.·��id And b.���°汾 = c.�汾�� And a.״̬ = 2 And a.��ǰ���� Is Not Null"
        End If
        '��ǰ·��ͳ��
        strSql = strSql & " And a.·��id=[1]"
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
                
        strSql = strSql & " Group By a.��ǰ����, c.��׼סԺ��" & vbNewLine & _
                        " Order By a.��ǰ����)"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, IIf(lng�ϲ�·��ID > 0, lng�ϲ�·��ID, lngPathID), _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))

        If rsTmp.RecordCount = 0 Then
            lblMSG.Caption = "����ָ����ʱ�䷶Χ��δ�������·���Ĳ��ˡ�"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        Else
            .ChartGroups.Item(1).Data.NumPoints(1) = rsTmp.RecordCount
            .Header.Text = "סԺ�շֲ�ͼ " & vbCrLf & "(��׼סԺ�գ�" & IIf(InStr(rsTmp!��׼סԺ�� & "", "-") > 0, "", "��") & rsTmp!��׼סԺ�� & "��)"
        End If
        i = 1
        Do While Not rsTmp.EOF
            '��������
            .ChartGroups.Item(1).PointLabels.Add (rsTmp!��ǰ���� & "��")
            .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!���� & "")
                
            '����ÿ�������ı�ǩ
            Set chtLabel = .ChartLabels.Add()
            chtLabel.Offset = 5
            chtLabel.Border.Type = oc2dBorderShadow
            chtLabel.Border.Width = 2
            chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
            chtLabel.AttachMethod = oc2dAttachDataIndex
            chtLabel.AttachDataIndex.Point = i
            chtLabel.IsConnected = True
            chtLabel.Anchor = oc2dAnchorAuto
            chtLabel.Name = rsTmp!��ǰ���� & ""
            chtLabel.Text = "��" & rsTmp!���� & "��(" & rsTmp!���� & "%)"
            chtLabel.Font.Size = 8
            
            If i <> 1 Then
                If .ChartGroups.Item(1).Data.Y(1, i) > .ChartArea.Axes.Item(2).Max Then
                    .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 5
                End If
            Else
                .ChartArea.Axes.Item(2).Max = .ChartGroups.Item(1).Data.Y(1, i) + 5
            End If
            i = i + 1
            rsTmp.MoveNext
        Loop
        'ע����Ϣ
        lblZY.Caption = "ע����ͼ�ǵ�ǰ·���¶�Ӧ��ʱ�䷶Χ�����в��˵�סԺ�����ֲ�ͼ��" & vbCrLf & _
                        "���У�1����ͼͳ�ƵĶ������Ѿ��������굱ǰ·���Ĳ��ˡ�" & vbCrLf & _
                        "       2��ͳ�Ƶ�סԺ�ձ�ʾ������·���е�סԺ������"
    End With
End Sub

Private Sub Set��ҽ��ͳ��(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    cbo����(1).Visible = False
    lbl����(1).Visible = False
    chtThis.Visible = False
    lblZY.Visible = True
    picTable.Visible = True
    strHead = "����,2500,1;ҽ��,1400,1;������,1100,7;�뾶����,1100,7;�뾶��,1100,7;�����˳���,1100,7;�����˳���,1100,7;���������,1100,7;���������,1100,7;ҽ�����϶�,1100,7"
    Call InitTable(vsgInfo(vsg_ԭ��), strHead)
    vsgInfo(vsg_ԭ��).Width = picTable.Width
    vsgInfo(vsg_��Ŀ).Visible = False
    vsgInfo(VSG_��ϸ).Visible = False
    fraGroupLR.Visible = False
    fraGroupUD.Visible = False
    imgFrom.Visible = False
    txtFindNum.Visible = False
    lblInfo(1).Caption = ""
    lblInfo(0).Caption = "��ҽ��ͳ��·��������Ϣ"
    vsgInfo(vsg_ԭ��).MergeCells = flexMergeRestrictColumns
    vsgInfo(vsg_ԭ��).MergeCol(VCOL_����) = True
    
    '��strDateTmp =����ʱ�� ��ѯʱ  ������ҳ D ������\����Ժ����\��Ժ���ڲ�ѯʱ,������ҳ D ��Ч
    strSql = "Select a.����id, b.���� As ����, a.������, Sum(a.������) As ������, Sum(a.�뾶����) As �뾶����, Sum(a.�뾶��) || '%' As �뾶��," & vbNewLine & _
        "       Sum(a.�����˳���) As �����˳���, Sum(a.�����˳���) || '%' As �����˳���, Sum(a.���������) As ���������, Sum(a.���������) || '%' As ���������," & vbNewLine & _
        "       Round(Decode(Nvl(Sum(a.ҽ����), 0), 0, '0', (Nvl(Sum(a.ҽ����), 0) - Nvl(Sum(a.·����ҽ��), 0)) / Nvl(Sum(a.ҽ����), 0)) * 100," & vbNewLine & _
        "              2) || '%' As ҽ�����϶�" & vbNewLine & _
        "From (Select a.����id, a.������, Count(1) As ������, Sum(Decode(a.״̬, 0, 0, 1)) As �뾶����," & vbNewLine & _
        "              Round(Sum(Decode(a.״̬, 0, 0, 1)) / Count(1) * 100, 2) As �뾶��, Sum(Decode(a.״̬, 3, 1, 0)) As �����˳���," & vbNewLine & _
        "              Decode(Sum(Decode(a.״̬, 0, 0, 1)), 0, '0'," & vbNewLine & _
        "                      Round(Sum(Decode(a.״̬, 3, 1, 0)) / Sum(Decode(a.״̬, 0, 0, 1)) * 100, 2)) As �����˳���," & vbNewLine & _
        "              Sum(Decode(a.״̬, 100, 1, 0)) As ���������," & vbNewLine & _
        "              Decode(Sum(Decode(a.״̬, 0, 0, 1)), 0, '0'," & vbNewLine & _
        "                      Round(Sum(Decode(a.״̬, 100, 1, 0)) / Sum(Decode(a.״̬, 0, 0, 1)) * 100, 2)) As ���������, Null As ҽ����," & vbNewLine & _
        "              Null As ·����ҽ��" & vbNewLine & _
        "       From (Select a.Id, a.����id, a.������," & vbNewLine & _
        "                     Decode(a.״̬, 2, Decode(Sign(Sum(Decode(p.�������, -1, 1, 0))), 0, 2, 1, 100), a.״̬) As ״̬" & vbNewLine & _
        "              From �����ٴ�·�� A, ������ҳ D, ����·������ P" & vbNewLine & _
        "              Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.Id = p.·����¼id(+) And a.״̬ <> 1 " & IIf(optAllPath.Value, "", " And a.·��id=[1]") & vbNewLine & _
        " And " & strDateTmp & " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
        "              Group By a.Id, a.����id, a.������, a.״̬) A" & vbNewLine & _
        "       Group By a.����id, a.������ "
    strSql = strSql & vbNewLine & _
        "   Union All   " & vbNewLine & _
        "       Select ����id, ������, Null, Null, Null, Null, Null, Null, Null, Count(1) As ҽ����, Sum(·����ҽ��) As ·����ҽ��" & vbNewLine & _
        "       From (Select Distinct a.����id, a.������, c.Id, Decode(b.·��ִ��id, Null, 1, 0) As ·����ҽ��" & vbNewLine & _
        "              From �����ٴ�·�� A, ����ҽ����¼ C, ����·��ҽ�� B, ������ҳ D" & vbNewLine & _
        "              Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.����id = c.����id And a.��ҳid = c.��ҳid And c.Id = b.����ҽ��id(+) And" & vbNewLine & _
        "                    c.���id Is Null And c.ǰ��id Is Null And c.��ʼִ��ʱ�� Between a.��ʼʱ�� And" & vbNewLine & _
        "                    Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.״̬ = 2 " & IIf(optAllPath.Value, "", " And a.·��id=[1]") & vbNewLine & _
                    " And " & strDateTmp & " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS'))" & vbNewLine & _
        "       Group By ����id, ������) A, ���ű� B" & vbNewLine & _
        "Where a.����id = b.Id" & vbNewLine & _
        "Group By a.����id, a.������, b.����" & vbNewLine & _
        "Order By b.����, a.����id, Sum(a.�����˳���) Desc"

    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
                    
    With vsgInfo(vsg_ԭ��)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, VCOL_����) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, VCOL_ҽ������) = rsTmp!������ & ""
                .TextMatrix(.Rows - 1, VCOL_������) = rsTmp!������ & ""
                .TextMatrix(.Rows - 1, vcol_�뾶����) = rsTmp!�뾶���� & ""
                .TextMatrix(.Rows - 1, vcol_�뾶��) = rsTmp!�뾶�� & ""
                .TextMatrix(.Rows - 1, vcol_�����˳���) = rsTmp!�����˳��� & ""
                .TextMatrix(.Rows - 1, vcol_�����˳���) = rsTmp!�����˳��� & ""
                .TextMatrix(.Rows - 1, vcol_���������) = rsTmp!��������� & ""
                .TextMatrix(.Rows - 1, vcol_���������) = rsTmp!��������� & ""
                .TextMatrix(.Rows - 1, VCOL_ҽ�����϶�) = rsTmp!ҽ�����϶� & ""
                rsTmp.MoveNext
            Loop
        Else
            .Rows = 2
        End If
    End With
    
    'ע����Ϣ
    lblZY.Caption = _
                    "˵����1��ҽ�����϶�=��·��ģ�������ҽ����/ҽ���������·���Ĳ���·���ڼ��ҽ������" & vbCrLf & _
                    "       2��ҽ�����϶��е�ҽ��������ҽ�������´��ҽ����·���ڼ����⣨����ǰ����ɺ�ģ���ҽ����" & vbCrLf & _
                    "       3��ҽ����ָ·���ĵ����ˣ���һ���ǲ��˵�סԺҽʦ��"
    '�����ϴ������ͼ
    mlngVariation = tbcVariation.Selected.Index
End Sub

Private Sub set���ұ���������(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    cbo����(1).Visible = False
    lbl����(1).Visible = False
    chtThis.Visible = False
    lblZY.Visible = True
    picTable.Visible = True
    fraGroupUD.Visible = False
    fraGroupLR.Visible = True
    vsgInfo(vsg_��Ŀ).Visible = True
    imgFrom.Visible = False
    txtFindNum.Visible = False
    vsgInfo(VSG_��ϸ).Visible = False
    lblInfo(1).Caption = "���ұ��������ʮ��"
    lblInfo(0).Caption = "���ұ��������ʮ��"
    
    strHead = "����,3000,1;�����˳���,1500,7;���������,1500,7"
    Call Grid.Init(vsgInfo(vsg_ԭ��), strHead)
    
    strHead = "����,3000,1;�����˳���,1500,7;���������,1500,7"
    Call Grid.Init(vsgInfo(vsg_��Ŀ), strHead)
    
    vsgInfo(vsg_��Ŀ).MergeCells = flexMergeRestrictColumns
    vsgInfo(vsg_��Ŀ).MergeCol(vsgInfo(vsg_��Ŀ).ColIndex("�����˳���")) = False
    vsgInfo(vsg_��Ŀ).MergeCol(vsgInfo(vsg_��Ŀ).ColIndex("���������")) = False
    vsgInfo(vsg_ԭ��).MergeCol(vsgInfo(vsg_ԭ��).ColIndex("�����˳���")) = False
    vsgInfo(vsg_ԭ��).MergeCol(vsgInfo(vsg_ԭ��).ColIndex("���������")) = False
            
    strSql = "Select a.����id, a.���� As ����, Count(1), Round(Sum(Decode(a.״̬, 3, 1, 0)) / Count(1) * 100, 2) As �����˳���," & vbNewLine & _
            "       Round(Sum(Decode(a.״̬, 100, 1, 0)) / Count(1) * 100, 2) As ���������" & vbNewLine & _
            "From (Select a.Id, a.����id, b.����, Decode(a.״̬, 2, Decode(Sign(Sum(Decode(p.�������, -1, 1, 0))), 0, 2, 1, 100), a.״̬) As ״̬" & vbNewLine & _
            "       From �����ٴ�·�� A, ������ҳ D, ���ű� B, ����·������ P" & vbNewLine & _
            "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.����id = b.Id And a.Id = p.·����¼id  And a.״̬ In (2, 3) " & vbNewLine
    '��ǰ·��ͳ��
    strSql = strSql & IIf(optAllPath.Value, "", " And a.·��id=[1]")
    'ʱ�䷶Χ
    strSql = strSql & " And " & strDateTmp & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
    strSql = strSql & _
            "Group By a.Id, a.����id, b.����, a.״̬) A" & vbNewLine & _
            "Group By a.����id, a.����" & vbNewLine & _
            "Order By �����˳��� Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
        Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
    
    With vsgInfo(vsg_ԭ��)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, .ColIndex("�����˳���")) = rsTmp!�����˳��� & "%"
                .TextMatrix(.Rows - 1, .ColIndex("���������")) = rsTmp!��������� & "%"
                If .Rows = 11 Then Exit Do
                rsTmp.MoveNext
            Loop
        Else
            .Rows = 2
        End If
    End With
    
    With vsgInfo(vsg_��Ŀ)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            rsTmp.Sort = "�����˳���"
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, .ColIndex("�����˳���")) = rsTmp!�����˳��� & "%"
                 .TextMatrix(.Rows - 1, .ColIndex("���������")) = rsTmp!��������� & "%"
                If .Rows = 11 Then Exit Do
                rsTmp.MoveNext
            Loop
        Else
            .Rows = 2
        End If
            
    End With
    
    'ע����Ϣ
    lblZY.Caption = "˵���������ʽ����������˳��ġ�"
    '�����ϴ������ͼ
    mlngVariation = tbcVariation.Selected.Index
End Sub

Private Sub set����֢����(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    Dim dblTmp As Double
    
    cbo����(1).Visible = False
    lbl����(1).Visible = False
    lblMSG.Visible = False
    chtThis.Visible = True
    lblZY.Visible = True
    optAllPath.Enabled = False
    optThisPath.Enabled = False
    '�����˳�����
    With chtThis
        .ChartGroups.Item(1).ChartType = oc2dTypePie
        .ChartArea.Border.Width = 4
        .Header.Text = "����֢���ɷֲ�ͼ"
        .Header.Interior.ForegroundColor = vbBlack
        '.ChartArea.Pie.StartAngle = 90
        strSql = "Select b.������� as ����֢,decode(count(1) over(),0,'0',round(count(1)/sum(count(1)) over() * 100,2)) as ����" & vbNewLine & _
            "From �����ٴ�·�� A, ������ҳ D, ������ϼ�¼ B" & vbNewLine & _
            "Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.����id=b.����id and a.��ҳid=b.��ҳid and b.�������=10" & vbNewLine & _
            "And a.״̬ In (2, 3) "
        
        '�Ƿ�ǰ·��ͳ��
        strSql = strSql & " And a.·��id=[1]"
        'ʱ�䷶Χ
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.�������,b.����id,b.���id Order by ����*100 desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
                    
        .ChartGroups.Item(1).Data.NumSeries = rsTmp.RecordCount
        If .ChartGroups.Item(1).Data.NumSeries <> 0 Then .ChartGroups.Item(1).Data.NumPoints(1) = 1
        i = 1
        Do Until rsTmp.EOF
            
            If i = 6 Then
                .ChartGroups.Item(1).SeriesLabels.Add ("����" & Space(2) & "��" & 100 - dblTmp & "%)")
                .ChartGroups.Item(1).Data.Y(i, 1) = 100 - dblTmp
                Exit Do
            Else
                .ChartGroups.Item(1).SeriesLabels.Add (rsTmp!����֢ & Space(2) & "��" & rsTmp!���� & "%)")
                .ChartGroups.Item(1).Data.Y(i, 1) = Val(rsTmp!���� & "")
                dblTmp = dblTmp + Val(rsTmp!���� & "")
            End If
            rsTmp.MoveNext
            i = i + 1
        Loop
        'ע����Ϣ
        lblZY.Caption = "˵����1�����ֻ��ʾ�������������ǰ��λ����֢������Ĺ��������ࡣ" & vbCrLf & _
                        "       2������֢����������ɺͱ����˳��Ĳ��˵Ĳ���֢��"
        If rsTmp.RecordCount = 0 Then
            lblMSG.Caption = "����ָ����ʱ�䷶Χ��δ���ֲ���֢�Ĳ��ˡ�"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
    End With
    '�����ϴ������ͼ
    mlngVariation = tbcVariation.Selected.Index
End Sub

Private Function Get�������·�����(ByVal varTime As Variant, ByVal lngPathID As Long) As Recordset
'���ܣ������������ļ�¼�����ݲ�ͬ��ʱ��
    Dim strSql As String
    Dim lngTmp As Long
      
    strSql = "Select Sum(a.������) As ������, Sum(a.�뾶����) As �뾶����, Nvl(Round(Sum(a.�뾶����) / Sum(a.������) * 100, 2), 0) || '%' As �뾶��," & vbNewLine & _
        "       Sum(a.�����˳���) As �����˳���," & vbNewLine & _
        "       Decode(Sum(a.�뾶����), 0, '0', Nvl(Round(Sum(a.�����˳���) / Sum(a.�뾶����) * 100, 2), 0)) || '%' As �����˳���," & vbNewLine & _
        "       Sum(a.���������) As ���������," & vbNewLine & _
        "       Decode(Sum(a.�뾶����), 0, '0', Nvl(Round(Sum(a.���������) / Sum(a.�뾶����) * 100, 2), 0)) || '%' As ���������," & vbNewLine & _
        "       Nvl(Round(Decode(Nvl(Sum(a.ҽ����), 0), 0, '0', (Nvl(Sum(a.ҽ����), 0) - Nvl(Sum(a.·����ҽ��), 0)) / Nvl(Sum(a.ҽ����), 0)) * 100," & vbNewLine & _
        "                  2), 0) || '%' As ҽ�����϶�, Round(Sum(סԺ����) / Sum(a.������), 2) As ƽ��סԺ��," & vbNewLine & _
        "       Round(Sum(��Ԥ��) / Sum(a.������), 2) As ƽ��סԺ����," & vbNewLine & _
        "       Nvl(Decode(Sum(a.�뾶����), 0, '0', 100 - Round(Sum(a.�����˳���) / Sum(a.�뾶����) * 100, 2)), 0) || '%' As �����"
strSql = strSql & vbNewLine & _
        " From (Select a.����id, a.·��id, Count(1) As ������, Sum(�뾶����) As �뾶����, Sum(�����˳���) As �����˳���, Sum(���������) As ���������, Sum(a.סԺ����) As סԺ����," & vbNewLine & _
                "       Sum(a.��Ԥ��) As ��Ԥ��, Null As ҽ����, Null As ·����ҽ��" & vbNewLine & _
                "From (Select a.����id, a.·��id, a.����id, Decode(a.״̬, 0, 0, 1) As �뾶����, Decode(a.״̬, 3, 1, 0) As �����˳���," & vbNewLine & _
                "              Decode(a.״̬, 2, Decode(Sign(Sum(Decode(p.�������, -1, 1, 0))), 0, 0, 1, 1), 0) As ���������, a.סԺ����, a.��Ԥ��" & vbNewLine & _
                "       From (Select a.Id, a.����id, a.·��id, a.����id, d.סԺ����, a.״̬, Sum(b.��Ԥ��) As ��Ԥ��" & vbNewLine & _
                "              From �����ٴ�·�� A, ������ҳ D, ����Ԥ����¼ B" & vbNewLine & _
                "              Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.����id = b.����id(+) And b.��ҳid(+) = a.��ҳid And a.״̬ <> 1 And" & vbNewLine & _
                "                    b.��¼���� In (1, 11, 2, 12) " & vbNewLine & _
                IIf(optAllPath.Value, "", " And a.·��id=[1]") & _
                " And d.��Ժ���� Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
                "              Group By a.Id, a.����id, a.·��id, a.����id, d.סԺ����, a.״̬) A, ����·������ P" & vbNewLine & _
                "       Where a.Id = p.·����¼id(+)" & vbNewLine & _
                "       Group By a.����id, a.·��id, a.����id, a.סԺ����, a.��Ԥ��, a.״̬) A" & vbNewLine & _
                "Group By a.����id, a.·��id"
    strSql = strSql & vbNewLine & _
            "   Union All   " & vbNewLine & _
            "Select ����id, ·��id, Null, Null, Null, Null, Null, Null, Count(1) As ҽ����, Sum(·����ҽ��) As ·����ҽ��" & vbNewLine & _
            "From (Select Distinct a.����id, a.·��id, c.Id, Decode(b.·��ִ��id, Null, 1, 0) As ·����ҽ��" & vbNewLine & _
            "       From �����ٴ�·�� A, ����ҽ����¼ C, ����·��ҽ�� B, ������ҳ D" & vbNewLine & _
            "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.����id = c.����id And a.��ҳid = c.��ҳid And c.Id = b.����ҽ��id(+) And" & vbNewLine & _
            "             c.���id Is Null And c.ǰ��id Is Null And c.��ʼִ��ʱ�� Between a.��ʼʱ�� And" & vbNewLine & _
            "             Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.״̬ = 2 " & IIf(optAllPath.Value, "", " And a.·��id=[1]") & vbNewLine & _
             " And D.��Ժ���� Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS'))" & vbNewLine & _
            "Group By ����id, ·��id) A"


    lngTmp = cboYorM.ListIndex
        
    Set Get�������·����� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
        IIf(lngTmp = 0 Or lngTmp = 1, Format(varTime, "yyyy-MM-01 00:00:00"), Format(varTime, "yyyy-01-01 00:00:00")), _
        IIf(lngTmp = 0, Format(DateAdd("M", 1, CDate(varTime)), "yyyy-MM-01 00:00:00"), IIf(lngTmp = 1, Format(DateAdd("M", 3, CDate(varTime)), "yyyy-MM-01 00:00:00"), Format(Format(varTime, "yyyy") + 1 & "-01-01", "yyyy-MM-dd 00:00:00"))))
End Function

Private Function Get����������Ҳ�����(ByVal varTime As Variant, ByVal lngPathID As Long) As Recordset
'���ܣ������������Ŀ��Ҳ����������ݲ�ͬ��ʱ��
    Dim strSql As String
    
    strSql = "Select /*+ rule*/Sum(������) As ������, Sum(������) As ������" & vbNewLine & _
        "From (Select 1 ������, Null As ������" & vbNewLine & _
        "       From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
        "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.״̬ <> 1" & vbNewLine & _
        IIf(optAllPath.Value, "", " And a.·��id=[1]") & _
        " And D.��Ժ����" & _
        " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
        "       Group By a.����id" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select Null, 1" & vbNewLine & _
        "       From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
        "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.״̬ <> 1" & vbNewLine & _
        IIf(optAllPath.Value, "", " And a.·��id=[1]") & _
        " And D.��Ժ����" & _
        " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
        "       Group By a.·��id)"

    Set Get����������Ҳ����� = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
        IIf(cboYorM.ListIndex = 0, Format(varTime, "yyyy-MM-01 00:00:00"), Format(varTime, "yyyy-01-01 00:00:00")), IIf(cboYorM.ListIndex = 0, Format(CDate(Format(varTime, "yyyy-mm")) + 32, "yyyy-MM-01 00:00:00"), Format(Format(varTime, "yyyy") + 1 & "-01-01", "yyyy-MM-dd 00:00:00")))

      
End Function

Private Sub set��������ȶ�(ByVal lngPathID As Long)
    Dim rsTmp As Recordset
    
    Set rsTmp = Get�������·�����(dtpTwo.Value, lngPathID)
    
    With vsgInfo(vsg_ԭ��)
        .TextMatrix(3, VCOL_ͬ�ڶ�) = Val(rsTmp!������ & "")
        .TextMatrix(4, VCOL_ͬ�ڶ�) = Val(rsTmp!�뾶���� & "")
        .TextMatrix(5, VCOL_ͬ�ڶ�) = Val(rsTmp!�뾶���� & "") - Val(rsTmp!�����˳��� & "")
        .TextMatrix(6, VCOL_ͬ�ڶ�) = rsTmp!�뾶�� & ""
        .TextMatrix(7, VCOL_ͬ�ڶ�) = rsTmp!����� & ""
        .TextMatrix(8, VCOL_ͬ�ڶ�) = rsTmp!��������� & ""
        .TextMatrix(9, VCOL_ͬ�ڶ�) = rsTmp!�����˳��� & ""
        .TextMatrix(10, VCOL_ͬ�ڶ�) = rsTmp!ҽ�����϶� & ""
        .TextMatrix(11, VCOL_ͬ�ڶ�) = Val(rsTmp!ƽ��סԺ�� & "")
        .TextMatrix(12, VCOL_ͬ�ڶ�) = Val(rsTmp!ƽ��סԺ���� & "")
        
        
        .TextMatrix(1, VCOL_��ֵ) = Val(.TextMatrix(1, VCOL_ͬ��һ)) - Val(.TextMatrix(1, VCOL_ͬ�ڶ�))
        .TextMatrix(2, VCOL_��ֵ) = Val(.TextMatrix(2, VCOL_ͬ��һ)) - Val(.TextMatrix(2, VCOL_ͬ�ڶ�))
        .TextMatrix(3, VCOL_��ֵ) = Val(.TextMatrix(3, VCOL_ͬ��һ)) - Val(.TextMatrix(3, VCOL_ͬ�ڶ�))
        .TextMatrix(4, VCOL_��ֵ) = Val(.TextMatrix(4, VCOL_ͬ��һ)) - Val(.TextMatrix(4, VCOL_ͬ�ڶ�))
        .TextMatrix(5, VCOL_��ֵ) = Val(.TextMatrix(5, VCOL_ͬ��һ)) - Val(.TextMatrix(5, VCOL_ͬ�ڶ�))
        
        'val(89.3%) -����С����Ͱٷֺ�ͬʱ������val�����б�ʵʱ���� �����⴦��
        '------------------------------------------
        .TextMatrix(6, VCOL_��ֵ) = Val(Replace(.TextMatrix(6, VCOL_ͬ��һ), "%", "")) - Val(Replace(.TextMatrix(6, VCOL_ͬ�ڶ�), "%", "")) & "%"
        .TextMatrix(7, VCOL_��ֵ) = Val(Replace(.TextMatrix(7, VCOL_ͬ��һ), "%", "")) - Val(Replace(.TextMatrix(7, VCOL_ͬ�ڶ�), "%", "")) & "%"
        .TextMatrix(8, VCOL_��ֵ) = Val(Replace(.TextMatrix(8, VCOL_ͬ��һ), "%", "")) - Val(Replace(.TextMatrix(8, VCOL_ͬ�ڶ�), "%", "")) & "%"
        .TextMatrix(9, VCOL_��ֵ) = Val(Replace(.TextMatrix(9, VCOL_ͬ��һ), "%", "")) - Val(Replace(.TextMatrix(9, VCOL_ͬ�ڶ�), "%", "")) & "%"
        .TextMatrix(10, VCOL_��ֵ) = Val(Replace(.TextMatrix(10, VCOL_ͬ��һ), "%", "")) - Val(Replace(.TextMatrix(10, VCOL_ͬ�ڶ�), "%", "")) & "%"
        '------------------------------------------
        .TextMatrix(11, VCOL_��ֵ) = Val(.TextMatrix(11, VCOL_ͬ��һ)) - Val(.TextMatrix(11, VCOL_ͬ�ڶ�))
        .TextMatrix(12, VCOL_��ֵ) = Val(.TextMatrix(12, VCOL_ͬ��һ)) - Val(.TextMatrix(12, VCOL_ͬ�ڶ�))
        
        If optAllPath.Value Then
            Set rsTmp = Get����������Ҳ�����(dtpTwo.Value, lngPathID)
            .RowHidden(1) = False
            .RowHidden(2) = False
            .TextMatrix(1, VCOL_ͬ�ڶ�) = Val(rsTmp!������ & "")
            .TextMatrix(2, VCOL_ͬ�ڶ�) = Val(rsTmp!������ & "")
        Else
            .RowHidden(1) = True
            .RowHidden(2) = True
        End If
        If cboYorM.ListIndex = 1 Then
            .TextMatrix(0, VCOL_ͬ�ڶ�) = Format(dtpTwo.Value, dtpTwo.CustomFormat) & "-" & Format(dtpFour.Value, dtpFour.CustomFormat)
        Else
            .TextMatrix(0, VCOL_ͬ�ڶ�) = Format(dtpTwo.Value, dtpTwo.CustomFormat)
        End If
    End With
End Sub

Private Sub set�������(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String

    cbo����(1).Visible = False
    lbl����(1).Visible = False
    chtThis.Visible = False
    lblZY.Visible = True
    
    picContrast.Visible = True
    Call SetPicContrastFace  '�������
    
    picFind.Visible = False
    picTable.Visible = True
    strHead = "ָ��,3000,1;2012��10��,1500,7;2012��11��,1500,7;��ֵ,1500,7"
    Call InitTable(vsgInfo(vsg_ԭ��), strHead)
    vsgInfo(vsg_ԭ��).Width = picTable.Width
    vsgInfo(vsg_��Ŀ).Visible = False
    vsgInfo(VSG_��ϸ).Visible = False
    fraGroupLR.Visible = False
    fraGroupUD.Visible = False
    imgFrom.Visible = False
    txtFindNum.Visible = False
    lblInfo(1).Caption = ""
    lblInfo(0).Caption = "ͳ��ҽԺ�ٴ�·���������"
    
    vsgInfo(vsg_ԭ��).MergeCells = flexMergeRestrictColumns
    vsgInfo(vsg_ԭ��).MergeCol(VCOL_����) = False
    vsgInfo(vsg_ԭ��).Rows = 11
    chkContrast_Click
     With vsgInfo(vsg_ԭ��)
        .Rows = 13
        
        Set rsTmp = Get�������·�����(dtpOne.Value, lngPathID)
        .TextMatrix(1, VCOL_ָ��) = "������"
        .TextMatrix(2, VCOL_ָ��) = "������"
        .TextMatrix(3, VCOL_ָ��) = "����������"
        .TextMatrix(4, VCOL_ָ��) = "�뾶������"
        .TextMatrix(5, VCOL_ָ��) = "���������"
        .TextMatrix(6, VCOL_ָ��) = "�뾶��"
        .TextMatrix(7, VCOL_ָ��) = "�����"
        .TextMatrix(8, VCOL_ָ��) = "���������"
        .TextMatrix(9, VCOL_ָ��) = "�����˳���"
        .TextMatrix(10, VCOL_ָ��) = "ҽ�����϶�"
        .TextMatrix(11, VCOL_ָ��) = "ƽ��סԺ����"
        .TextMatrix(12, VCOL_ָ��) = "ƽ��סԺ�ܷ���"
        
        .TextMatrix(3, VCOL_ͬ��һ) = Val(rsTmp!������ & "")
        .TextMatrix(4, VCOL_ͬ��һ) = Val(rsTmp!�뾶���� & "")
        .TextMatrix(5, VCOL_ͬ��һ) = Val(rsTmp!�뾶���� & "") - Val(rsTmp!�����˳��� & "")
        .TextMatrix(6, VCOL_ͬ��һ) = rsTmp!�뾶�� & ""
        .TextMatrix(7, VCOL_ͬ��һ) = rsTmp!����� & ""
        .TextMatrix(8, VCOL_ͬ��һ) = rsTmp!��������� & ""
        .TextMatrix(9, VCOL_ͬ��һ) = rsTmp!�����˳��� & ""
        .TextMatrix(10, VCOL_ͬ��һ) = rsTmp!ҽ�����϶� & ""
        .TextMatrix(11, VCOL_ͬ��һ) = Val(rsTmp!ƽ��סԺ�� & "")
        .TextMatrix(12, VCOL_ͬ��һ) = Val(rsTmp!ƽ��סԺ���� & "")
        
        If optAllPath.Value Then
            Set rsTmp = Get����������Ҳ�����(dtpOne.Value, lngPathID)
            .RowHidden(1) = False
            .RowHidden(2) = False
            .TextMatrix(1, VCOL_ͬ��һ) = Val(rsTmp!������ & "")
            .TextMatrix(2, VCOL_ͬ��һ) = Val(rsTmp!������ & "")
        Else
            .RowHidden(1) = True
            .RowHidden(2) = True
        End If
        
        If cboYorM.ListIndex = 1 Then
            .TextMatrix(0, VCOL_ͬ��һ) = Format(dtpOne.Value, dtpOne.CustomFormat) & "-" & Format(dtpThree.Value, dtpThree.CustomFormat)
            .TextMatrix(0, VCOL_ͬ�ڶ�) = Format(dtpTwo.Value, dtpTwo.CustomFormat) & "-" & Format(dtpFour.Value, dtpFour.CustomFormat)
            Call .AutoSize(VCOL_ͬ��һ, VCOL_ͬ�ڶ�)
        Else
            .TextMatrix(0, VCOL_ͬ��һ) = Format(dtpOne.Value, dtpOne.CustomFormat)
            .TextMatrix(0, VCOL_ͬ�ڶ�) = Format(dtpTwo.Value, dtpTwo.CustomFormat)
        End If
    End With
    
    'ע����Ϣ
    lblZY.Caption = _
    "˵����1���ñ�ֻͳ�Ƴ�Ժ���ˡ�" & vbCrLf & _
    "       2��ҽ�����϶�=��·��ģ�������ҽ����/ҽ���������·���Ĳ���·���ڼ��ҽ������" & vbCrLf & _
    "       3����ȫԺ·��ͳ��ʱ����ͳ��ʹ���ٴ�·���Ŀ������Ͳ�������"
    '�����ϴ������ͼ
    mlngSurvey = tbcVariation.Selected.Index
End Sub

Private Function GetXNum() As Long
'���ܣ��������ͼX����ĵ���
    Dim lngXNum As Long
    
    If cboTrendTime.ListIndex = 0 Then
        '����
        If cboInterval.List(cboInterval.ListIndex) = "һ��" Then
            lngXNum = 7
        ElseIf cboInterval.List(cboInterval.ListIndex) = "һ��" Then
            lngXNum = DateAdd("M", 1, Format(dtpTrendStart.Value, "yyyy-MM-dd")) - CDate(Format(dtpTrendStart.Value, "yyyy-MM-dd"))
        ElseIf cboInterval.List(cboInterval.ListIndex) = "����" Then
            lngXNum = DateAdd("M", 2, Format(dtpTrendStart.Value, "yyyy-MM-dd")) - CDate(Format(dtpTrendStart.Value, "yyyy-MM-dd"))
        Else
            lngXNum = DateAdd("M", 3, Format(dtpTrendStart.Value, "yyyy-MM-dd")) - CDate(Format(dtpTrendStart.Value, "yyyy-MM-dd"))
        End If
    Else
        If cboInterval.List(cboInterval.ListIndex) = "����" Then
            lngXNum = 6
        ElseIf cboInterval.List(cboInterval.ListIndex) = "һ��" Then
            lngXNum = 12
        ElseIf cboInterval.List(cboInterval.ListIndex) = "����" Then
            lngXNum = 24
        Else
            lngXNum = 36
        End If
    End If
    GetXNum = lngXNum
End Function

Private Sub setƽ��סԺ����(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '����������
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
    
     lblMSG.Visible = False
     chtThis.Visible = True
     cbo����(1).Visible = False
     lbl����(1).Visible = False
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = False
     optAllPath.Enabled = False
     optIn.Visible = True
     optOut.Visible = True
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '����ͼ�δ�С
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '���������ɫ������
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '������Ӱ
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '����Ϊ3DЧ��
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '��������
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("����")
         .ChartGroups.Item(1).SeriesLabels.Add ("��׼ֵ")
         '���������ǩ
         
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
         If optIn.Value Then
            strSql = "select /*+ rule*/ƽ��סԺ����,��Ժ����,sum(ƽ��סԺ����) over() as ���� from " & _
            "(Select round(sum(d.סԺ����)/count(1),2) as ƽ��סԺ����,trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as ��Ժ����" & vbNewLine & _
            "       From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
            "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.״̬ =2" & vbNewLine & _
            "        And a.·��ID=[1]  " & _
            " And D.��Ժ����" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & "))"
         Else
            strSql = "Select /*+ rule*/ƽ��סԺ����, ��Ժ����, Sum(ƽ��סԺ����) Over() As ����" & vbNewLine & _
                "From (Select round(Sum(d.סԺ����) / Count(1),2) As ƽ��סԺ����, Trunc(d.��Ժ����, " & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") As ��Ժ����" & vbNewLine & _
                "       From (Select d.����id, d.��ҳid, Max(d.סԺ����) As סԺ����, Max(d.��Ժ����) As ��Ժ����" & vbNewLine & _
                "              From ������ϼ�¼ B, ������ҳ D" & vbNewLine & _
                "              Where b.����id = d.����id And b.��ҳid = d.��ҳid And b.��ϴ��� = 1 And b.������� In (1, 2, 11, 12) And" & vbNewLine & _
                "                    (Exists (Select 1 From �ٴ�·������ C Where b.����id = c.����id And Nvl(c.����, 0) = 0 And c.·��id = [1]) Or Exists" & vbNewLine & _
                "                     (Select 1 From �ٴ�·������ C Where b.���id = c.���id And Nvl(c.����, 0) = 0 And c.·��id = [1])) And" & vbNewLine & _
                "                    d.��Ժ���� Between To_Date([2], 'YYYY-MM-DD HH24:MI:SS') And" & vbNewLine & _
                "                    To_Date([3], 'YYYY-MM-DD HH24:MI:SS') And not exists(select 1 from �����ٴ�·�� E where b.����id=e.����id and b.��ҳid=e.��ҳid and e.״̬ =2)" & vbNewLine & _
                "              Group By d.����id, d.��ҳid) D" & vbNewLine & _
                "       Group By Trunc(d.��Ժ����, " & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & "))"

         End If
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!���� & "")
         For i = 1 To lngXNum
             '�����ʾ19����ǩ
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM��"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "��Ժ����=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!ƽ��סԺ���� & "")
                 If lngMax < Val(rsTmp!ƽ��סԺ���� & "") Then lngMax = Val(rsTmp!ƽ��סԺ���� & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
             
         Next
         .ChartArea.Axes(2).Max = lngMax + lngMax / 5
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "ƽ��סԺ��������ͼ"
         
        
         'ע����Ϣ
         lblZY.Caption = "˵����1����ͼֻͳ�Ƴ�Ժ���ˡ�" & vbCrLf & _
                         "       2��ͳ�Ƶ�סԺ�ձ�ʾ����ʵ�ʵ�סԺ������" & vbCrLf & _
                         "       3����׼ֵ��ָͳ��ʱ���ڼ��ƽ��ֵ��"
         '�����ϴ������ͼ
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub setƽ��סԺ����(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '����������
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
    
     lblMSG.Visible = False
     chtThis.Visible = True
     cbo����(1).Visible = False
     lbl����(1).Visible = False
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = False
     optAllPath.Enabled = False
     optIn.Visible = True
     optOut.Visible = True
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '����ͼ�δ�С
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '���������ɫ������
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '������Ӱ
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '����Ϊ3DЧ��
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '��������
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("����(Ԫ)")
         .ChartGroups.Item(1).SeriesLabels.Add ("��׼ֵ")
         '���������ǩ
         
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
         If optIn.Value Then
            strSql = "Select /*+ rule*/ƽ��סԺ����, ��Ժ����, Sum(ƽ��סԺ����) Over() As ���� From " & _
            "(select sum(��Ԥ��) as ��Ԥ��,��Ժ����,round(sum(��Ԥ��)/sum(����),2) as ƽ��סԺ���� from " & _
            "(Select sum(b.��Ԥ��) as ��Ԥ��,1 as ����,trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as ��Ժ����" & vbNewLine & _
            "       From �����ٴ�·�� A, ������ҳ D,����Ԥ����¼ B" & vbNewLine & _
            "       Where a.����id = d.����id And a.��ҳid = d.��ҳid And a.����id=b.����id and a.��ҳid=b.��ҳid And a.״̬ =2 And b.��¼���� in(1,11,2,12) " & vbNewLine & _
            "        And a.·��ID=[1]  " & _
            " And D.��Ժ����" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ,a.ID ) group by ��Ժ����)"
         Else
            strSql = "Select /*+ rule*/ƽ��סԺ����, ��Ժ����, Sum(ƽ��סԺ����) Over() As ����" & vbNewLine & _
                "From (Select round(sum(��Ԥ��)/Count(1),2) As ƽ��סԺ����, Trunc(d.��Ժ����, " & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") As ��Ժ����" & vbNewLine & _
                "       From (Select d.����id, d.��ҳid, sum(c.��Ԥ��) As ��Ԥ��, Max(d.��Ժ����) As ��Ժ����" & vbNewLine & _
                "              From ������ϼ�¼ B, ������ҳ D,����Ԥ����¼ C" & vbNewLine & _
                "              Where b.����id = d.����id And b.��ҳid = d.��ҳid  And c.����id=b.����id and c.��ҳid=b.��ҳid And b.��ϴ��� = 1 And b.������� In (1, 2, 11, 12) And c.��¼���� in(1,11,2,12) And" & vbNewLine & _
                "                    (Exists (Select 1 From �ٴ�·������ C Where b.����id = c.����id And Nvl(c.����, 0) = 0 And c.·��id = [1]) Or Exists" & vbNewLine & _
                "                     (Select 1 From �ٴ�·������ C Where b.���id = c.���id And Nvl(c.����, 0) = 0 And c.·��id = [1])) And" & vbNewLine & _
                "                    d.��Ժ���� Between To_Date([2], 'YYYY-MM-DD HH24:MI:SS') And" & vbNewLine & _
                "                    To_Date([3], 'YYYY-MM-DD HH24:MI:SS') And not exists(select 1 from �����ٴ�·�� E where b.����id=e.����id and b.��ҳid=e.��ҳid and e.״̬=2)" & vbNewLine & _
                "              Group By d.����id, d.��ҳid) D" & vbNewLine & _
                "       Group By Trunc(d.��Ժ����, " & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & "))"

         End If
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!���� & "")
         For i = 1 To lngXNum
             '�����ʾ19����ǩ
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM��"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "��Ժ����=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!ƽ��סԺ���� & "")
                 If lngMax < Val(rsTmp!ƽ��סԺ���� & "") Then lngMax = Val(rsTmp!ƽ��סԺ���� & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
             
         Next
         .ChartArea.Axes(2).Max = lngMax + lngMax / 5
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "ƽ��סԺ��������ͼ"
         
        
         'ע����Ϣ
        lblZY.Caption = "˵����1����ͼֻͳ�Ƴ�Ժ���ˡ�" & vbCrLf & _
                        "       2��סԺ����ֻ�������˵�Ԥ����ͽ��ʲ��" & vbCrLf & _
                        "       3����׼ֵ��ָͳ��ʱ���ڼ��ƽ��ֵ��"


         '�����ϴ������ͼ
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub set�뾶��(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '����������
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
    
     lblMSG.Visible = False
     chtThis.Visible = True
     cbo����(1).Visible = False
     lbl����(1).Visible = False
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = True
     optAllPath.Enabled = True
     optIn.Visible = False
     optOut.Visible = False
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '����ͼ�δ�С
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '���������ɫ������
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '������Ӱ
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '����Ϊ3DЧ��
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '��������
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("�뾶��(%)")
         .ChartGroups.Item(1).SeriesLabels.Add ("��׼ֵ")
         '���������ǩ
         
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
        strSql = "Select �뾶��, ��Ժ����, Sum(�뾶��) Over() As ���� From " & _
            "(Select round(sum(decode(a.״̬,0,0,1))/count(1) *100,2) as �뾶��,trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as ��Ժ����" & vbNewLine & _
            "       From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
            "       Where a.����id = d.����id And a.��ҳid = d.��ҳid " & vbNewLine & _
            IIf(optAllPath.Value, "", " And a.·��id=[1] ") & _
            " And D.��Ժ����" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ) "
         
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!���� & "")
         For i = 1 To lngXNum
             '�����ʾ19����ǩ
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM��"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "��Ժ����=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!�뾶�� & "")
                 If lngMax < Val(rsTmp!�뾶�� & "") Then lngMax = Val(rsTmp!�뾶�� & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             '��׼��
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
             
         Next
         .ChartArea.Axes(2).Max = IIf(lngMax + lngMax / 5 > 100, 100, lngMax + lngMax / 5)
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "·���뾶��"
         
        
         'ע����Ϣ
        lblZY.Caption = "˵����1����ͼֻͳ�Ƴ�Ժ���ˡ�" & vbCrLf & _
                        "       2����׼ֵ��ָͳ��ʱ���ڼ��ƽ��ֵ��"
         '�����ϴ������ͼ
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub set�����(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '����������
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
    
     lblMSG.Visible = False
     chtThis.Visible = True
     cbo����(1).Visible = False
     lbl����(1).Visible = False
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = True
     optAllPath.Enabled = True
     optIn.Visible = False
     optOut.Visible = False
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '����ͼ�δ�С
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '���������ɫ������
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '������Ӱ
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '����Ϊ3DЧ��
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '��������
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("�����(%)")
         .ChartGroups.Item(1).SeriesLabels.Add ("��׼ֵ")
         '���������ǩ
         
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
        strSql = "Select �����, ��Ժ����, Sum(�����) Over() As ���� From " & _
            "(Select round(sum(decode(a.״̬,2,1,0))/count(1) *100,2) as �����,trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as ��Ժ����" & vbNewLine & _
            "       From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
            "       Where a.����id = d.����id And a.��ҳid = d.��ҳid and a.״̬ in(2,3) " & vbNewLine & _
            IIf(optAllPath.Value, "", " And a.·��id=[1] ") & _
            " And D.��Ժ����" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ) "
         
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!���� & "")
         For i = 1 To lngXNum
             '�����ʾ19����ǩ
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM��"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "��Ժ����=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!����� & "")
                 If lngMax < Val(rsTmp!����� & "") Then lngMax = Val(rsTmp!����� & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             '��׼��
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
             
         Next
         .ChartArea.Axes(2).Max = IIf(lngMax + lngMax / 5 > 100, 100, lngMax + lngMax / 5)
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "·�������"
         
        
         'ע����Ϣ
        lblZY.Caption = "˵����1����ͼֻͳ�Ƴ�Ժ���ˡ�" & vbCrLf & _
                        "       2����׼ֵ��ָͳ��ʱ���ڼ��ƽ��ֵ��"
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub set������(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '����������
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
    
     lblMSG.Visible = False
     chtThis.Visible = True
     cbo����(1).Visible = False
     lbl����(1).Visible = False
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = True
     optAllPath.Enabled = True
     optIn.Visible = False
     optOut.Visible = False
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '����ͼ�δ�С
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '���������ɫ������
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '������Ӱ
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '����Ϊ3DЧ��
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '��������
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("������(%)")
         .ChartGroups.Item(1).SeriesLabels.Add ("��׼ֵ")
         '���������ǩ
         
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
        strSql = "Select ������, ��Ժ����, Sum(������) Over() As ���� From " & _
            "(Select round(sum(decode(a.״̬,3,1,0))/count(1) *100,2) as ������,trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as ��Ժ����" & vbNewLine & _
            "       From �����ٴ�·�� A, ������ҳ D" & vbNewLine & _
            "       Where a.����id = d.����id And a.��ҳid = d.��ҳid and a.״̬ in(2,3) " & vbNewLine & _
            IIf(optAllPath.Value, "", " And a.·��id=[1] ") & _
            " And D.��Ժ����" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(d.��Ժ����," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ) "
         
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!���� & "")
         For i = 1 To lngXNum
             '�����ʾ19����ǩ
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM��"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "��Ժ����=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!������ & "")
                 If lngMax < Val(rsTmp!������ & "") Then lngMax = Val(rsTmp!������ & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             '��׼��
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
             
         Next
         .ChartArea.Axes(2).Max = IIf(lngMax + lngMax / 5 > 100, 100, lngMax + lngMax / 5)
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "·��������"
         
        
         'ע����Ϣ
        lblZY.Caption = "˵����1����ͼֻͳ�Ƴ�Ժ���ˡ�" & vbCrLf & _
                        "       2����׼ֵ��ָͳ��ʱ���ڼ��ƽ��ֵ��"
         '�����ϴ������ͼ
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub tbcVariation_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strSql As String, rsTmp As Recordset
    Dim lngPathID As Long
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strDateTmp As String
    Dim lng�ϲ�·��ID As Long
    Dim strHead As String
    Dim dblTmp As Double
    
    If mblnFirstLoad And Item.Tag <> "��ҽ��ͳ��" Then Exit Sub
    
    strDateTmp = cboTimeType.List(cboTimeType.ListIndex)
    
    If strDateTmp = "����ʱ��" Then strDateTmp = "A.����ʱ��"
    If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժ����"
    If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժ����"
    On Error GoTo errH
    With chtThis
        '��ʼͼ�θ�ʽ
        '��Ļ��ֹˢ�£�������ɺ�����Ϊfalse
        .IsBatched = True
        picTable.Visible = False
        picTrend.Visible = False
        vsgInfo(VSG_��ϸ).Visible = True
        fraGroupUD.Visible = True
        imgFrom.Visible = True
        lblInfo(2).Visible = True
        If InStr(mstrPrivs, "ȫԺ·��") <> 0 Then
            optThisPath.Enabled = True
            optAllPath.Enabled = True
        End If
        picContrast.Visible = False
        picFind.Visible = True
        lblPathEdition.Visible = False
        cboPathEdition.Visible = False
        .Reset
        .AllowUserChanges = False
        .ChartGroups.Item(1).Data.NumSeries = 0
        .ChartArea.Border = oc2dBorderShadow
        .Border = oc2dBorderEtchedIn
        '�ұߵı�ǩ����
        .Legend.Border = oc2dBorder3DOut
        .Legend.Border.Width = 4
        'ͼ�α�ͷ
        .Header.IsShowing = True
        .Header.Font.Size = 18
        .Header.Font.Name = "����"
        .Header.Font.Bold = True
        '����Ϊ3DЧ��
        .ChartArea.View3D.depth = 20
        .ChartArea.View3D.Elevation = 20
        '����ͼ�δ�С
        .ChartArea.PlotArea.Top = 60
        .ChartArea.PlotArea.Left = 55
        .ChartArea.PlotArea.Right = 60
        .ChartArea.PlotArea.Bottom = 35
        
        If rptPath.SelectedRows.count > 0 Or optAllPath.Value Then
            If Not rptPath.SelectedRows(0).GroupRow Or optAllPath.Value Then
                If rptPath.SelectedRows.count > 0 And Not rptPath.SelectedRows(0).GroupRow Then lngPathID = Val(rptPath.SelectedRows(0).Record(COL_ID).Value)
                Select Case Item.Tag
                    
                    Case "δ����ԭ��"
                        Call Setδ����ԭ��(strDateTmp, lngPathID)
                    Case "�����˳�����"
                        Call Set�����˳�����(strDateTmp, lngPathID)
                    Case "ʱ��������"
                        Call setʱ��������(strDateTmp, lngPathID)
                    Case "δ����ԭ��"
                        Call Setδ����ԭ��(strDateTmp, lngPathID)
                    Case "·������Ŀ"
                        Call Set·������Ŀ(strDateTmp, lngPathID)
                    Case "·��������"
                        Call Set·��������(strDateTmp, lngPathID)
                    Case "�׶�ƽ������"
                        Call Set�׶�ƽ������(strDateTmp, lngPathID)
                    Case "סԺ�շֲ�ͼ"
                        Call SetסԺ�շֲ�ͼ(strDateTmp, lngPathID)
                    Case "��ҽ��ͳ��"
                        Call Set��ҽ��ͳ��(strDateTmp, lngPathID)
                    Case "���ұ���������"
                        Call set���ұ���������(strDateTmp, lngPathID)
                    Case "����֢����"
                        Call set����֢����(strDateTmp, lngPathID)
                    Case "�������"
                        Call set�������(strDateTmp, lngPathID)
                    Case "ƽ��סԺ����"
                        Call setƽ��סԺ����(lngPathID)
                    Case "ƽ��סԺ����"
                        Call setƽ��סԺ����(lngPathID)
                    Case "�뾶��"
                        Call set�뾶��(lngPathID)
                    Case "�����"
                        Call set�����(lngPathID)
                    Case "������"
                        Call set������(lngPathID)
                End Select
                
    
            Else
                lblMSG.Caption = "����ǰ·��ͳ����Ҫѡ��һ��·����"
                lblMSG.Visible = True
                .Visible = False
                lblZY.Visible = False
                .ChartArea.Border.Width = 0
                cbo����(1).Visible = False
                lbl����(1).Visible = False
            End If
        Else
            lblMSG.Caption = "����ǰ·��ͳ����Ҫѡ��һ��·����"
            lblMSG.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
            cbo����(1).Visible = False
            lbl����(1).Visible = False
        End If
        .IsBatched = False
        .Refresh
        Call picTable_Resize
        If Me.Visible And InStr(";��ҽ��ͳ��;���ұ���������;δ����ԭ��;·������Ŀ;�������;", ";" & Item.Tag & ";") > 0 Then
            Call SetFlagBySelectedTable(True, "VSGINFO_0")
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call FuncFindPath
    End If
End Sub

Private Sub FuncFindPath(Optional ByVal blnNext As Boolean)
'������blnNext=�Ƿ������һ��
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long, strFindTmp As String
    
    Call zlControl.TxtSelAll(txtFind)
            
    '��ʼ������
    If rptPath.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl����������0��ʼ
    Else
        i = rptPath.SelectedRows(0).Index + 1
    End If
    
    '����·��
    strFindTmp = txtFind.Text
    For i = i To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If Not .GroupRow Then
                If cbsMain.FindControl(, 0).Caption = "�����Ʋ���" Then
                    If .Record(COL_����).Value Like "*" & strFindTmp & "*" Then Exit For
                Else
                    If .Record(COL_��ϱ���).Value Like "*" & Trim(strFindTmp) & "*" Or _
                       .Record(COL_��������).Value Like "*" & Trim(strFindTmp) & "*" Or _
                       .Record(COL_�������).Value Like "*" & strFindTmp & "*" Or _
                       .Record(COL_��������).Value Like "*" & strFindTmp & "*" _
                       Then Exit For
                End If
            End If
        End With
    Next

    If i <= rptPath.Rows.count - 1 Then
        blnReStart = False
        '����ѡ������ʾ�ڿɼ�����,������SelectionChanged�¼�
        Set rptPath.FocusedRow = rptPath.Rows(i)
        
        If rptPath.Visible Then rptPath.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "������", "") & "�Ҳ��������������ٴ�·����", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "�������ٴ�·�����ơ���ϻ��߼���" & vbCrLf & "����(Ctrl+F)" & vbCrLf & "������һ��(F3)"
    
    zlCommFun.ShowTipInfo txtFind.Hwnd, strTip, True
End Sub

Private Function LoadPatiList(ByVal lng·��ID As Long, Optional ByVal lngPersonID As Long, Optional ByVal lng�ϲ�·��ID As Long) As Boolean
'���ܣ���ȡ·��Ӧ�õĲ����嵥
'������lng�ϲ�·��ID<>0������ϲ�·���б�
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intState As Integer
    
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strDateTmp As String
    Dim strTyping As String
    Dim strIsVariation As String
    Dim strBranch As String
    Dim strBranchName As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    mlng����ID = 0
    mlng��ҳID = 0
    mlng����·��ID = 0
    
    If optState(0).Value Then intState = 0
    If optState(1).Value Then intState = 1
    If optState(2).Value Or optState(4).Value Then
        strIsVariation = " And " & IIf(optState(2).Value, "Not", "") & " Exists (Select 1 From ����·������ Where ·����¼id = a.Id And ������� = -1) "
        intState = 2
    End If
    If optState(3).Value Then intState = 3
    
    '�ϲ�·��
    If lng�ϲ�·��ID = 0 Then
        cbo����(1).Clear
        cbo����(0).Clear
        cbo����(0).AddItem "ȫ��·��"
        cbo����(0).AddItem "��Ҫ·��"
        cbo����(0).ItemData(cbo����(0).NewIndex) = -1
        cbo����(1).AddItem "ȫ��·��"
        cbo����(1).AddItem "��Ҫ·��"
        cbo����(1).ItemData(cbo����(1).NewIndex) = -1
        Call cbo.SetIndex(cbo����(0).Hwnd, 0)
        Call cbo.SetIndex(cbo����(1).Hwnd, 0)
        strSql = "Select Distinct a.Id, a.����" & vbNewLine & _
                "From �ٴ�·��Ŀ¼ a, �ٴ�·������ b" & vbNewLine & _
                "Where b.·��id(+) = a.Id And a.���� = 1 And" & vbNewLine & _
                "           (a.ͨ�� = 1 Or a.ͨ�� = 2 And Exists (Select 1 From �ٴ�·������ c Where c.·��id = [1] And b.����id = c.����id))"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                cbo����(0).AddItem rsTmp!���� & ""
                cbo����(1).AddItem rsTmp!���� & ""
                cbo����(0).ItemData(cbo����(0).NewIndex) = rsTmp!ID & ""
                cbo����(1).ItemData(cbo����(1).NewIndex) = rsTmp!ID & ""
                rsTmp.MoveNext
            Loop
        Else
            cbo����(1).Enabled = False
            cbo����(0).Enabled = False
        End If
    End If
                
    
    strDateTmp = cboForDate.List(cboForDate.ListIndex)
    
    If strDateTmp = "����ʱ��" Then strDateTmp = "A.����ʱ��"
    If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժ����"
    If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժ����"
    
    '�����Ϻͱ����˳���ʾԭ��
    rptPati.Columns(COL_������ԭ��).Visible = optState(0)
    rptPati.Columns(COL_�����˳�ԭ��).Visible = optState(3)
        
    strTyping = cboTyping.List(cboTyping.ListIndex)
    If strTyping <> "" Then
        strTyping = " And " & "��������='" & strTyping & "'"
    End If
    If cbo·����Χ.ListIndex > 0 Then
        If cbo·����Χ.List(cbo·����Χ.ListIndex) = "��·��" Then
            strBranch = " And k.���� is null"
        Else
            strBranch = " And k.���� = [5]"
        End If
        strBranchName = cbo·����Χ.List(cbo·����Χ.ListIndex)
    End If
    
    strSql = "Select Distinct A.ID,a.����id, a.��ҳid, f.���� As ����,NVL(D.����, e.����) ����,NVL(D.�Ա�, e.�Ա�) �Ա� ,NVL(D.����, e.����) ���� , d.סԺ��, d.סԺҽʦ, d.��Ժ���� As ����, d.��ǰ���� As ����, a.״̬, a.��ǰ����, a.�汾��," & vbNewLine & _
            "       b.���°汾, c.��׼סԺ��, c.��׼����, a.������, a.����ʱ��, a.����ʱ��, d.��ǰ����id As ����id, d.��Ժ����id As ����id, d.״̬ As ����״̬, d.����ת��,j.��ӡ��,j.��ӡʱ��,a.�ϲ�·������,d.������,z.���� as ����������," & vbNewLine & _
            "       i.���� As ������ԭ��, " & _
            IIf(intState = 2, "''", "decode(a.״̬,3,g.����,'')") & _
            " As �����˳�ԭ��,NVL(k.����,'δ���֧') as ��֧����" & vbNewLine & _
            IIf(cbo·����Χ.ListIndex <> 1, ",Min(Decode(n.��֧id, Null, 9999, m.����)) As ��֧���", "") & _
            ",Decode(Q.Id,Null,0,1) as ���߰��ӡ" & vbNewLine & _
            " From �����ٴ�·�� A, �ٴ�·��Ŀ¼ B, �ٴ�·���汾 C," & _
            IIf(intState = 2, "", " ����·������ H, ���쳣��ԭ�� G,") & _
            "�ٴ�·���׶� L,�ٴ�·����֧ K," & _
            IIf(cbo·����Χ.ListIndex <> 1, "����·��ִ�� M,�ٴ�·���׶� N,", "") & _
            " ������ҳ D, ������Ϣ E, ���ű� F, ���쳣��ԭ�� I,���Ӳ�����ӡ J,������Ŀ¼ Z , ���Ӳ�����ӡ Q " & vbNewLine & _
            " Where a.·��id = b.Id And a.·��id = c.·��id And d.������ = z.����(+) And a.�汾�� = c.�汾�� And a.����id = d.����id And a.��ҳid = d.��ҳid And a.����id = e.����id And" & vbNewLine & _
            "      a.����id = f.Id And j.�ļ�id(+) = a.Id And j.����(+) = 11 And (j.Id = (Select Max(ID) From ���Ӳ�����ӡ Where �ļ�id(+) = a.Id And ���� = 11) Or j.Id Is Null)" & _
            "And Q.�ļ�id(+) = a.Id And Q.����(+) = 12 And (Q.Id = (Select Max(ID) From ���Ӳ�����ӡ Where �ļ�id(+) = a.Id And ���� = 12) Or Q.Id Is Null )" & vbNewLine & _
            IIf(intState = 2, "", " And h.·����¼id(+) = a.Id And h.����(+) = a.��ǰ���� And g.����(+) = h.����ԭ�� ") & _
            " And i.����(+) = a.δ����ԭ��" & vbNewLine & _
            " And A.·��ID=[1]" & _
            " And NVL(a.��ǰ�׶�ID,a.ǰһ�׶�ID)=l.ID(+) And l.��֧ID=k.id(+) " & _
            IIf(cbo·����Χ.ListIndex <> 1, " And m.·����¼ID(+)=a.id And m.�׶�ID=n.ID(+) ", "") & _
            IIf(lng�ϲ�·��ID > 0, " And", IIf(lng�ϲ�·��ID = 0, "", " And Not")) & _
            IIf(lng�ϲ�·��ID <> 0, "  exists(Select 1 from ���˺ϲ�·�� O Where o.��Ҫ·����¼ID=a.ID " & _
            IIf(lng�ϲ�·��ID > 0, " And o.·��ID=[6])", ")"), "")
    If lngPersonID = 0 Then
        strSql = strSql & strTyping & " And A.״̬=[2]" & _
        " And " & strDateTmp & _
        " Between To_Date([3],'YYYY-MM-DD HH24:MI:SS') And To_Date([4],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & strIsVariation
        strSql = strSql & strBranch
        
        If intState = 3 Then
            strSql = strSql & " And g.����=2"
        ElseIf intState = 0 Then
            strSql = strSql & " And i.����=0"
        End If
        
        If cbo·����Χ.ListIndex <> 1 Then
            strSql = strSql & " Group By A.ID,a.����id, a.��ҳid, f.����, NVL(D.����, e.����),NVL(D.�Ա�, e.�Ա�) ,NVL(D.����, e.����), d.סԺ��, d.סԺҽʦ, d.��Ժ����, d.��ǰ����, a.״̬, a.��ǰ����, a.�汾��, b.���°汾, c.��׼סԺ��," & vbNewLine & _
                    "         c.��׼����, a.������, a.����ʱ��, a.����ʱ��, d.��ǰ����id, d.��Ժ����id,d.������, d.״̬, d.����ת��, j.��ӡ��, j.��ӡʱ��,a.�ϲ�·������,z.����, i.����, " & IIf(intState = 2, " ", "decode(a.״̬,3,g.����,''),") & vbNewLine & _
                    "         Nvl(k.����, 'δ���֧'),Decode(Q.Id,Null,0,1) "

        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID, intState, _
        Format(dtpTime(0).Value, "yyyy-MM-dd 00:00:00"), Format(dtpTime(1).Value, "yyyy-MM-dd 23:59:59"), strBranchName, lng�ϲ�·��ID)
    Else
        '���Ҳ��ˣ�����ʱ�����Ϣ
        strSql = strSql & " And e.����id=[2]"
        If cbo·����Χ.ListIndex <> 1 Then
            strSql = strSql & " Group By A.ID,a.����id, a.��ҳid, f.����, NVL(D.����, e.����),NVL(D.�Ա�, e.�Ա�) ,NVL(D.����, e.����), d.סԺ��, d.סԺҽʦ, d.��Ժ����, d.��ǰ����, a.״̬, a.��ǰ����, a.�汾��, b.���°汾, c.��׼סԺ��," & vbNewLine & _
                    "         c.��׼����, a.������, a.����ʱ��, a.����ʱ��, d.��ǰ����id, d.��Ժ����id,d.������, d.״̬, d.����ת��, j.��ӡ��, j.��ӡʱ��,a.�ϲ�·������,z.����, i.����, " & IIf(intState = 2, " ", "decode(a.״̬,3,g.����,''),") & vbNewLine & _
                    "         Nvl(k.����, 'δ���֧'),Decode(Q.Id,Null,0,1) "

        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng·��ID, lngPersonID)
    End If
    
    '��¼��ˢ�º�Ĳ��˼�¼��������ӡʹ��
    Set mrsTmp = rsTmp
    
    rptPati.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPati.Records.Add()
        
        Set objItem = objRecord.AddItem("")
        objItem.HasCheckbox = True
            If rptPati.Columns(col_��ӡ).Icon = img16.ListImages("UnCheck").Index - 1 Then
                objItem.Checked = True
            Else
                objItem.Checked = False
            End If
        Set objItem = objRecord.AddItem(Val(rsTmp!����ID))
        Set objItem = objRecord.AddItem(Val(rsTmp!��ҳID))
        Set objItem = objRecord.AddItem(CStr(rsTmp!����))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        If rsTmp!������ & "" <> "" Then
            objItem.Icon = img16.ListImages("������").Index - 1
        End If
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!�Ա�)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!סԺ��)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!סԺҽʦ)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!����)))
        
        If NVL(rsTmp!״̬, 0) = 1 And Not IsNull(rsTmp!��ǰ����) Then
            If InStr(rsTmp!��׼סԺ��, "-") > 0 Then
                Set objItem = objRecord.AddItem(CInt(Val(rsTmp!��ǰ����) / Val(Split(rsTmp!��׼סԺ��, "-")(1)) * 100) & "%")
            Else
                Set objItem = objRecord.AddItem(CInt(Val(rsTmp!��ǰ����) / Val(rsTmp!��׼סԺ��) * 100) & "%")
            End If
        Else
            Set objItem = objRecord.AddItem("")
        End If
        
        Set objItem = objRecord.AddItem(NVL(rsTmp!��׼סԺ��) & IIf(Not IsNull(rsTmp!��׼סԺ��), "��", ""))
        Set objItem = objRecord.AddItem(NVL(rsTmp!��׼����) & IIf(Not IsNull(rsTmp!��׼����), "Ԫ", ""))
        Set objItem = objRecord.AddItem(rsTmp!�汾�� & "/" & rsTmp!���°汾)
        
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!������)))
        Set objItem = objRecord.AddItem(Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm"))
        Set objItem = objRecord.AddItem(Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm"))
        
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!����ID, 0)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!����ID, 0)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!����״̬, 0)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!����ת��, 0)))
        Set objItem = objRecord.AddItem(NVL(rsTmp!������ԭ��))
        Set objItem = objRecord.AddItem(NVL(rsTmp!�����˳�ԭ��))
        Set objItem = objRecord.AddItem(NVL(rsTmp!��ӡ��))
        Set objItem = objRecord.AddItem(NVL(Format(rsTmp!��ӡʱ��, "yyyy-MM-dd HH:mm")))
        Set objItem = objRecord.AddItem(NVL(rsTmp!��֧����))
        If cbo·����Χ.ListIndex <> 1 Then
            Set objItem = objRecord.AddItem(NVL(IIf(rsTmp!��֧��� & "" = "9999", "", "��" & rsTmp!��֧��� & "��"), ""))
        End If
        Set objItem = objRecord.AddItem(NVL(rsTmp!�ϲ�·������))
        Set objItem = objRecord.AddItem(NVL(rsTmp!����������))
        Set objItem = objRecord.AddItem(IIf(rsTmp!���߰��ӡ = 0, "", " ��"))
        Set objItem = objRecord.AddItem(rsTmp!ID & "")
        rsTmp.MoveNext
    Loop
    
    If cbo·����Χ.ListIndex = 0 Then
        Me.rptPati.Columns(COL_��֧·�����).Visible = True
        Me.rptPati.Columns(COL_��֧·��).Visible = True
    Else
        Me.rptPati.Columns(COL_��֧·��).Visible = False
        If cbo·����Χ.ListIndex = 1 Then
            Me.rptPati.Columns(COL_��֧·�����).Visible = False
        Else
            Me.rptPati.Columns(COL_��֧·�����).Visible = True
        End If
    End If
    rptPati.Populate
    
    If rptPati.Rows.count = 0 Then
        Me.stbThis.Panels(3).Text = ""
    Else
        Me.stbThis.Panels(3).Text = "��ǰ·������ " & rptPati.Records.count & " ��Ӧ�ò���"
    End If
    '���ô���ߴ�
    cbsMain_Resize

    Screen.MousePointer = 0
    LoadPatiList = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadOperInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '���ܣ���ȡ·��Ӧ�õĲ����嵥
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intSource As Integer
    
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    On Error GoTo errH
    Screen.MousePointer = 11
    intSource = -1
    strSql = "Select Id,��¼��Դ,��������,�������� As ��������,����ҽʦ,����ҽʦ From ���������¼ Where ����ID=[1] And ��ҳID=[2] Order By ��¼��Դ"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID, lng��ҳID)
    
    rptOper.Records.DeleteAll
    Do While Not rsTmp.EOF
        If intSource = -1 Then intSource = Val("" & rsTmp!��¼��Դ)
        If intSource = Val("" & rsTmp!��¼��Դ) Then
            Set objRecord = Me.rptOper.Records.Add()
            
            Set objItem = objRecord.AddItem("" & rsTmp!ID)
            Set objItem = objRecord.AddItem("" & rsTmp!��������)
            Set objItem = objRecord.AddItem("" & rsTmp!��������)
            Set objItem = objRecord.AddItem("" & rsTmp!����ҽʦ)
            Set objItem = objRecord.AddItem("" & rsTmp!����ҽʦ)
        Else
            Exit Do
        End If
        rsTmp.MoveNext
    Loop
    rptOper.Populate
    
    Screen.MousePointer = 0
    LoadOperInfo = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub zlRptPrint(ByVal bytMode As Byte)
'����:��¼���ӡ
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objReport As ReportControl
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    Dim objVSF As VSFlexGrid
    Dim blnIsRPT As Boolean   'True-��ReportControl������Ҫת����VSF����
    Dim blnPath As Boolean    'True-��� �ٴ�·���嵥
    Dim strTmp As String
    Dim objTable As Object

    If rptPath.SelectedRows.count = 1 Then
        Select Case tbcSub.Selected.Caption
        Case "����·��"
            If rptPati.Records.count > 0 And mstrFlag = "RPTPATI" Then
                Set objTable = rptPati
                strSubhead = rptPath.SelectedRows(0).Record(COL_����).Value & "Ӧ�ò����嵥"
            Else
                blnPath = True  '�ٴ�·���嵥
            End If
            blnIsRPT = True
        Case "�������", "�ſ�����"
             '����ҽ��ͳ�ơ���"���ұ���������"����δ����ԭ�򡱡���·������Ŀ���������������
            If mstrFlag = "RPTPATH" Then
                blnPath = True: blnIsRPT = True
            Else
                If optAllPath.Value And optAllPath.Enabled Then
                    strTmp = "ȫԺ·��"
                Else
                    If Not rptPath.SelectedRows(0).GroupRow Then
                        strTmp = rptPath.SelectedRows(0).Record(COL_����).Value
                    End If
                End If
                Select Case tbcVariation.Selected.Caption
                Case "��ҽ��ͳ��"
                   If mstrFlag <> "" And mstrFlag = "VSGINFO_0" Then
                       Set objTable = vsgInfo(vsg_ԭ��)
                       strSubhead = strTmp & "_��ҽ��ͳ��·��������Ϣ"
                   End If
                Case "���ұ���������"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_ԭ��)
                        strSubhead = strTmp & "_���ұ��������ʮ��"
                    ElseIf mstrFlag = "VSGINFO_1" Then
                        Set objTable = vsgInfo(vsg_��Ŀ)
                        strSubhead = strTmp & "_���ұ��������ʮ��"
                    End If
                Case "δ����ԭ��"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_ԭ��)
                        strSubhead = strTmp & "_δ����ԭ����ܱ�"
                    ElseIf mstrFlag = "VSGINFO_1" Then
                        Set objTable = vsgInfo(vsg_��Ŀ)
                        strSubhead = strTmp & "_δ������Ŀ���ܱ�"
                    ElseIf mstrFlag = "VSGINFO_2" Then
                        Set objTable = vsgInfo(VSG_��ϸ)
                        strSubhead = strTmp & "_δ������Ŀ��ϸ��"
                    End If
                Case "·������Ŀ"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_ԭ��)
                        strSubhead = strTmp & "_·������Ŀ����ԭ����ܱ�"
                    ElseIf mstrFlag = "VSGINFO_1" Then
                        Set objTable = vsgInfo(vsg_��Ŀ)
                        strSubhead = strTmp & "_·������Ŀ��Ӧҽ�����ܱ�"
                    End If
                   
                Case "�������"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_ԭ��)
                        strSubhead = "ҽԺ�ٴ�·���������"
                    End If
                Case Else
                    blnPath = True: blnIsRPT = True '�ٴ�·���嵥
                End Select
            End If
        Case Else
            blnPath = True: blnIsRPT = True '�ٴ�·���嵥
        End Select
    End If
    
    If blnPath Then
        Set objTable = rptPath
        strSubhead = "�ٴ�·���嵥"  '��� �ٴ�·���嵥
    End If
    '-------------------------------------------------
    '�������ݱ��
    If blnIsRPT Then
        Set objReport = objTable
        If objReport.Records.count = 0 Then Exit Sub
        If zlControl.RPTCopyToVSF(objReport, vsTemp) Is Nothing Then Exit Sub
    Else
        Set objVSF = objTable
        If Grid.CopyTo(objVSF, vsTemp) Is Nothing Then Exit Sub
    End If

    '���ô�ӡ��������
    
    Set objPrint.Body = Me.vsTemp
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("��ӡ��:" & UserInfo.����)
    Call objAppRow.Add("��ӡʱ��:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)

    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    
End Sub

Private Sub zlRptBatPrint()
'���ܣ�������ӡ·����
    Dim i As Long
    
    With rptPati
        If rptPati.Rows.count < 1 Then MsgBox "��ǰ�����б�û�пɴ�ӡ��·����", vbInformation, Me.Caption: Exit Sub
        If optState(0).Value Then MsgBox "��ǰѡ��Ĳ���Ϊ[������]��·�����ˣ�û�п��õ�·����", vbInformation, Me.Caption: Exit Sub
        If tbcSub.Selected.Tag <> "����·��" Then MsgBox "��ѡ��[����·��]��Ƭ���ٽ��д�ӡ������", vbInformation, Me.Caption: Exit Sub
        mrsTmp.Filter = 0
        For i = 1 To .Rows.count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(col_��ӡ).Checked Then
                    '������Ҫ��ӡ�Ĳ���
                    mrsTmp.Filter = IIf(mrsTmp.Filter = 0, "", mrsTmp.Filter) & IIf(mrsTmp.Filter = 0, "", " or ") & " (����ID =" & .Rows(i).Record(COL_����ID).Value & " And ��ҳID=" & .Rows(i).Record(COL_��ҳID).Value & ") "
                End If
            End If
        Next
        
        frmBatPrint.ShowMe Me, mrsTmp
    End With
End Sub

Private Sub FuncShowPath()
    Dim vPati As TYPE_Pati
    
    With rptPati.SelectedRows(0)
        vPati.����ID = .Record(COL_����ID).Value
        vPati.��ҳID = .Record(COL_��ҳID).Value
        vPati.����ID = .Record(COL_����ID).Value
        vPati.����ID = .Record(COL_����ID).Value
        vPati.����״̬ = .Record(COL_����״̬).Value
        
        frmPathTrackView.ShowMe Me, vPati, .Record(COL_����ת��).Value = 1
    End With
End Sub

Private Sub FuncShowReport()
    Dim lng·��ID As Long, str���� As String
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    If rptPath.SelectedRows.count <= 0 Then Exit Sub
    
    lng·��ID = rptPath.SelectedRows(0).Record(COL_ID).Value
    If lng·��ID <> 0 Then
        str���� = rptPath.SelectedRows(0).Record(COL_����).Value
        Call frmReport1.ShowMe(gfrmMain, lng·��ID, str����)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtFindNum_GotFocus()
    txtFindNum.Tag = "OK"
    txtFindNum.SelStart = 0
    txtFindNum.SelLength = Len(txtFindNum.Text)
End Sub

Private Sub txtFindNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call GetPathOutAdvice: vsgInfo(vsg_��Ŀ).SetFocus
End Sub

Private Sub txtFindNum_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) <= 0 And KeyAscii <> vbKeyBack Then KeyAscii = 0: Exit Sub
    If Len(txtFindNum.Text) >= 2 And KeyAscii <> vbKeyBack And txtFindNum.SelLength = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtFindNum_LostFocus()
    Call GetPathOutAdvice
End Sub

Private Sub txtPerson_GotFocus()
    txtPerson.SelStart = 0
    txtPerson.SelLength = Len(txtPerson.Text)
End Sub

Private Sub txtPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPerson.Tag = "�����"
        Call FindPerson
        txtPerson.Tag = ""
    End If
End Sub

Private Sub txtPerson_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "�����벡��������סԺ�š�"
    
    zlCommFun.ShowTipInfo txtPerson.Hwnd, strTip, True
End Sub

Private Sub FindPerson()
    Dim strSql As String, vRect As RECT, rsTmp As Recordset, strTmp As String, varPara As Variant, i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    '��������֣����סԺ�ţ����������
    On Error GoTo errH
    varPara = txtPerson.Text
    If IsNumeric(varPara) And InStr(varPara, ".") = 0 And InStr(varPara, "-") = 0 And InStr(varPara, "+") = 0 Then
        strTmp = " And E.סԺ��=[1]"
        varPara = CLng(txtPerson.Text)
    Else
        strTmp = " And E.���� like [1]"
        varPara = gstrLike & txtPerson.Text & "%"
    End If
    strSql = "Select Rownum As ID, a.·��id, a.����id, NVL(B.����,e.����) ����,NVL(B.�Ա�,e.�Ա�) �Ա�,NVL(B.����,e.����) ����, e.סԺ��, a.����ʱ��" & vbNewLine & _
            "From �����ٴ�·�� A, ������Ϣ E, ������ҳ B" & vbNewLine & _
            "Where a.����id = e.����id and E.����id=B.����id And E.��ҳID=B.��ҳID"
    strSql = strSql & strTmp
    vRect = zlControl.GetControlRect(txtPerson.Hwnd)
    
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, Me.Caption, _
            False, "", "", False, True, True, vRect.Left, vRect.Top, _
            txtPerson.Height, False, False, False, varPara)
            
    If rsTmp Is Nothing Then
        MsgBox "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
        Call txtPerson.SetFocus
        txtPerson.SelStart = 0
        txtPerson.SelLength = Len(txtPerson.Text)
        Exit Sub
    End If
    
    For i = 0 To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If Not .GroupRow Then
                If .Record(COL_ID).Value = Val("" & rsTmp!·��ID) Then Exit For
            End If
        End With
    Next

    If i <= rptPath.Rows.count - 1 Then
        rptPath.Tag = "1"
        Set rptPath.FocusedRow = rptPath.Rows(i)
        rptPath.Tag = ""
        If rptPath.Visible Then rptPath.SetFocus
    Else
        MsgBox "�Ҳ������������Ĳ��ˡ�", vbInformation, gstrSysName
    End If
    Call LoadPatiList(Val("" & rsTmp!·��ID), Val("" & rsTmp!����ID))
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsgInfo_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim strDateTmp As String
    Dim i As Long
    
    If Index = vsg_��Ŀ Then
        If Not vsgInfo(VSG_��ϸ).Visible Then Exit Sub
        If vsgInfo(vsg_��Ŀ).Rows = vsgInfo(vsg_��Ŀ).FixedRows And NewRow <> vsgInfo(vsg_��Ŀ).FixedRows - 1 Then Exit Sub
        vsgInfo(VSG_��ϸ).Rows = 1
        strDateTmp = cboTimeType.List(cboTimeType.ListIndex)
    
        If strDateTmp = "����ʱ��" Then strDateTmp = "A.����ʱ��"
        If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժʱ��"
        If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժʱ��"
        strSql = "Select d.����id, NVl(F.����,d.����) ����, d.סԺ��, c.�Ǽ���, e.���� As ԭ��, c.�Ǽ�ʱ��" & vbNewLine & _
                " From �����ٴ�·�� A, ����·��ִ�� C, �ٴ�·����Ŀ B, ������Ϣ D,������ҳ F, ���쳣��ԭ�� E" & vbNewLine & _
                " Where c.��Ŀid = b.Id And c.·����¼id = a.Id And d.����id = a.����id And a.��ҳid = d.��ҳid And F.����id = a.����id And F.��ҳid =a.��ҳid And e.���� = c.����ԭ�� And e.���� = 1 And" & vbNewLine & _
                "      c.��Ŀid Is Not Null And c.����ԭ�� Is Not Null"
        strSql = strSql & " And c.��Ŀid=[1]"
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS') Order by c.�Ǽ�ʱ�� desc"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(vsgInfo(vsg_��Ŀ).RowData(NewRow) & ""), _
                Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        With vsgInfo(VSG_��ϸ)
        For i = 1 To rsTmp.RecordCount
            
                .AddItem ""
                .RowData(i) = rsTmp!����ID & ""
                .TextMatrix(i, VCol_����) = rsTmp!���� & ""
                .TextMatrix(i, VCol_סԺ��) = rsTmp!סԺ�� & ""
                .TextMatrix(i, VCOL_ҽ��) = rsTmp!�Ǽ��� & ""
                .TextMatrix(i, VCol_δʹ��ԭ��) = rsTmp!ԭ�� & ""
                .TextMatrix(i, VCol_����ʱ��) = rsTmp!�Ǽ�ʱ�� & ""
                
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(VSG_��ϸ).Rows = 1 Then vsgInfo(VSG_��ϸ).Rows = 2
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetArea(ByVal lngRow As Long, ByVal lngCol As Long) As CONST_AREA
'���ܣ���ȡָ����������һ������
    With vsgInfo(vsg_��Ŀ)
        If lngRow = -1 Or lngCol = -1 Then
            GetArea = -1
        ElseIf lngRow <= .FixedRows - 1 Or lngCol <= .FixedCols - 1 Then
            GetArea = -1
        ElseIf lngCol >= .FixedCols And lngCol <= .FixedCols + .FrozenCols - 1 _
            And lngRow >= .FixedRows And lngRow <= .FixedRows + .FrozenRows - 1 Then
            GetArea = Area_Cross
        ElseIf lngCol >= .FixedCols And lngCol <= .FixedCols + .FrozenCols - 1 Then
            GetArea = Area_Category
        ElseIf lngRow >= .FixedRows And lngRow <= .FixedRows + .FrozenRows - 1 Then
            GetArea = Area_Step
        Else
            GetArea = Area_Item
        End If
    End With
End Function

Private Sub vsgInfo_Click(Index As Integer)
    Call SetFlagBySelectedTable(True, "VSGINFO_" & Index)
End Sub

Private Sub vsgInfo_DblClick(Index As Integer)
    Dim vArea As CONST_AREA
    Dim lngRow As Long, lngCol As Long
    Dim strSql As String, rsTmp As Recordset
    Dim strDateTmp As String
    Dim i As Long
    
    '˫����Ŀ���鿴ҽ��
    If Index = vsg_��Ŀ Then
        If Not vsgInfo(VSG_��ϸ).Visible Then Exit Sub
        With vsgInfo(vsg_��Ŀ)
            lngRow = .MouseRow
            lngCol = .MouseCol
            
            vArea = GetArea(lngRow, lngCol)
            If vArea <> Area_Cross And vArea <> -1 Then
                If Val(.RowData(lngRow)) <> 0 Then
                    Call frmPathItemEdit.ShowView(Me, Val(.RowData(lngRow)))
                End If
            End If
        End With
    End If
End Sub


Private Sub GetPathOutAdvice()
'���ܣ����·������Ŀ����Ӧ��ҽ����Ϣ
    Dim strSql As String, rsTmp As Recordset
    Dim lngPathID As Long
    Dim i As Long
    Dim strDateTmp As String

    
    If txtFindNum.Tag = "" Then Exit Sub
    
    strDateTmp = cboTimeType.List(cboTimeType.ListIndex)
    
    If strDateTmp = "����ʱ��" Then strDateTmp = "A.����ʱ��"
    If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժ����"
    If strDateTmp = "��Ժ����" Then strDateTmp = "D.��Ժ����"
    vsgInfo(vsg_��Ŀ).Rows = 1
    
    If rptPath.SelectedRows.count > 0 And Not rptPath.SelectedRows(0).GroupRow Then lngPathID = Val(rptPath.SelectedRows(0).Record(COL_ID).Value)
    'ҽ�����ݻ��ܱ�
    strSql = "select * from(Select c.������Ŀid, c.ҽ������, c.����, e.���� As �׶�����, Nvl(f.���, e.���) �׶����,ROW_NUMBER() over(PARTITION BY e.���� order by Nvl(f.���, e.���),c.���� desc) as Top" & vbNewLine & _
            " From (With Test As (Select g.Id, g.���id, h.���, h.���� As ������Ŀ����, h.��������, h.Id As ������Ŀid, g.ҽ������, c.�׶�id" & vbNewLine & _
            "                    From �����ٴ�·�� A, ����·��ҽ�� B, ����·��ִ�� C, ����ҽ����¼ G, ������ҳ D, ������ĿĿ¼ H" & vbNewLine & _
            "                    Where c.·����¼id = a.Id And b.·��ִ��id = c.Id And g.Id = b.����ҽ��id And d.����id = a.����id And a.��ҳid = d.��ҳid And" & vbNewLine & _
            "                          c.��Ŀid Is Null And h.Id = g.������Ŀid"
    strSql = strSql & " And a.·��id=[1]"
    'ʱ�䷶Χ
    strSql = strSql & " And " & strDateTmp & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
    strSql = strSql & ")" & vbNewLine & _
            "     --1����Ѫ��ͼ�����" & vbNewLine & _
            "       Select Test.�׶�id, Test.������Ŀid, Test.ҽ������ As ҽ������, Count(1) As ����" & vbNewLine & _
            "       From Test" & vbNewLine & _
            "       Where Test.���id Is Null And (Test.��� = 'K' Or (Test.��� = 'E' And Test.�������� = '6'))" & vbNewLine & _
            "       Group By Test.������Ŀid, Test.�׶�id, Test.ҽ������" & vbNewLine & _
            "       --2��һ����ҩ������ҩ;����ÿ��ҩ�ֿ���ʾ" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select Test.�׶�id, Test.������Ŀid, Test.������Ŀ���� As ҽ������, Count(1) As ����" & vbNewLine & _
            "       From Test" & vbNewLine & _
            "       Where Test.��� In ('4', '5', '6')" & vbNewLine & _
            "       Group By Test.������Ŀid, Test.�׶�id, Test.������Ŀ����"
    strSql = strSql & "--3����ҩ��ȡ�ۺϺ��������Ŀ����" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select �׶�id, Max(������Ŀid) As ������Ŀid, f_List2str(Cast(Collect(ҽ������) As t_Strlist)) ҽ������, Count(1) ����" & vbNewLine & _
            "       From (Select Test.�׶�id, Test.������Ŀid, Test.������Ŀ���� As ҽ������, Test.���id" & vbNewLine & _
            "              From Test" & vbNewLine & _
            "              Where Test.��� = '7'" & vbNewLine & _
            "              Order By ҽ������)" & vbNewLine & _
            "       Group By ���id, �׶�id" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       --4������" & vbNewLine & _
            "       Select Test.�׶�id, Test.������Ŀid, Test.������Ŀ���� As ҽ������, Count(1) As ����" & vbNewLine & _
            "       From Test" & vbNewLine & _
            "       Where Test.���id Is Null And (Test.��� <> 'E' Or (Test.��� = 'E' And Test.�������� Not In ('2', '4', '6'))) And Test.��� <> 'K'" & vbNewLine & _
            "       Group By Test.������Ŀid, Test.�׶�id, Test.������Ŀ����) C, �ٴ�·���׶� E, �ٴ�·���׶� F" & vbNewLine & _
            "Where e.Id = c.�׶�id And e.��id = f.Id(+)" & vbNewLine & _
            "  Order By �׶����, ���� Desc) where top<=[4]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"), IIf(Val(txtFindNum.Text) = 0, 5, Val(txtFindNum.Text)))
    
    With vsgInfo(vsg_��Ŀ)
    For i = 1 To rsTmp.RecordCount
    
            
            .AddItem ""
            .RowData(i) = rsTmp!������ĿID & ""
            .TextMatrix(i, VCol_�׶�) = rsTmp!�׶����� & ""
            .TextMatrix(i, VCol_����) = rsTmp!ҽ������ & ""
            .TextMatrix(i, VCol_��Ŀ����) = rsTmp!���� & ""
            
        rsTmp.MoveNext
    Next
    End With
    txtFindNum.Tag = ""
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SaveImage()
'���ܣ��򿪹����Ի���,����ͼƬ

    With dlgPublic
        .DialogTitle = "����ͼƬ�ļ�"
        .Filter = "Jpeg|*.jpg"
        .Flags = &H200000 + &H2000 + &H2 + &H800
        .InitDir = App.Path
        .FileName = Format(Now, "yyyymmddhhmmss")
        .ShowSave
        Call chtThis.SaveImageAsJpeg(.FileName, 100, False, False, False) '���ļ�����ͼƬ������0-100��:�Ƿ���ʾΪ�Ҷ�ͼ���Ƿ�ѹ�����Ƿ���ǿ��ʾ
    End With
End Sub

Private Sub SetPicContrastFace()
'����:����PicContrast������ʾЧ��
    '
    dtpThree.Visible = cboYorM.ListIndex = 1
    dtpFour.Visible = cboYorM.ListIndex = 1
    lblFromToOne.Visible = cboYorM.ListIndex = 1
    lblFromToTwo.Visible = cboYorM.ListIndex = 1
    
    '����λ��
    cboYorM.Left = 120
    If cboYorM.ListIndex = 1 Then '������
        lblFromToOne.Left = 1250
        dtpOne.Left = 1440
        dtpThree.Left = 3020
        chkContrast.Left = 4500
        lblFromToTwo.Left = 6300
        dtpTwo.Left = 6480
        dtpFour.Left = 8070
        cmdContrast.Left = dtpFour.Left + dtpFour.Width + 100
    Else
        dtpOne.Left = lblFromToOne.Left
        chkContrast.Left = dtpOne.Left + dtpOne.Width + 500
        dtpTwo.Left = chkContrast.Left + chkContrast.Width + 100
        cmdContrast.Left = dtpTwo.Left + dtpTwo.Width + 100
    End If

End Sub

Private Sub SetFlagBySelectedTable(Optional ByVal blnVisible As Boolean = True, Optional ByVal strFlag As String)
'���ܣ�ѡ�б��ʱ,����Ϸ���ʾ���ڱ�ʶѡ�е�ͼ�꣬�����û����
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngIndex As Long

    fraFlag.Visible = blnVisible And strFlag <> ""
    If strFlag <> mstrFlag Then
        mstrFlag = strFlag
    ElseIf strFlag = "" Then
        Exit Sub
    End If
    
    
    If mstrFlag <> "" Then
        If mstrFlag = "RPTPATH" Then
            Set fraFlag.Container = rptPath.Container
            lngLeft = rptPath.Left + rptPath.Width - 500
            lngTop = rptPath.Top
        ElseIf mstrFlag = "RPTPATI" Then
            Set fraFlag.Container = rptPati.Container
            lngLeft = rptPati.Left + rptPati.Width - 250
            lngTop = rptPati.Top
        ElseIf InStr(mstrFlag, "VSGINFO_") > 0 Then
            Set fraFlag.Container = picTable
            lngIndex = Val(Replace(mstrFlag, "VSGINFO_", ""))
            lngLeft = lblInfo(lngIndex).Left + lblInfo(lngIndex).Width + 120
            lngTop = lblInfo(lngIndex).Top
        End If

        fraFlag.Left = lngLeft
        fraFlag.Top = lngTop
    End If
End Sub
