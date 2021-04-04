VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmPathTrackOut 
   AutoRedraw      =   -1  'True
   Caption         =   "门诊路径跟踪"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   14370
   FillColor       =   &H80000008&
   Icon            =   "frmPathTrackOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   14370
   Begin XtremeReportControl.ReportControl rptPath 
      Height          =   3090
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   5450
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin MSComDlg.CommonDialog dlgPublic 
      Left            =   1320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "保存为图片"
      Filter          =   "Jpeg|*.jpg"
   End
   Begin VB.PictureBox picReason 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   4800
      ScaleHeight     =   2775
      ScaleWidth      =   8295
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   8295
      Begin VB.PictureBox picTable 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   3480
         ScaleHeight     =   2415
         ScaleWidth      =   4695
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
         Begin VB.TextBox txtFindNum 
            Height          =   270
            Left            =   4800
            TabIndex        =   35
            Text            =   "5"
            Top             =   240
            Width           =   350
         End
         Begin VB.Frame fraGroupUD 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   45
            Left            =   240
            MousePointer    =   7  'Size N S
            TabIndex        =   44
            Top             =   1080
            Width           =   2000
         End
         Begin VB.Frame fraGroupLR 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1000
            Left            =   2280
            MousePointer    =   9  'Size W E
            TabIndex        =   43
            Top             =   240
            Width           =   45
         End
         Begin VSFlex8Ctl.VSFlexGrid vsgInfo 
            Height          =   495
            Index           =   0
            Left            =   360
            TabIndex        =   40
            Top             =   360
            Width           =   1815
            _cx             =   3201
            _cy             =   873
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
            Height          =   495
            Index           =   1
            Left            =   2520
            TabIndex        =   41
            Top             =   360
            Width           =   1815
            _cx             =   3201
            _cy             =   873
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
            Height          =   735
            Index           =   2
            Left            =   480
            TabIndex        =   42
            Top             =   1440
            Width           =   2295
            _cx             =   4048
            _cy             =   1296
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
            Picture         =   "frmPathTrackOut.frx":058A
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   345
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "未生成项目明细表"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   47
            Top             =   1200
            Width           =   1440
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "未生成项目汇总表(双击查看对应医嘱)"
            Height          =   250
            Index           =   1
            Left            =   2520
            TabIndex        =   46
            Top             =   120
            Width           =   3060
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "未生成原因汇总表"
            Height          =   250
            Index           =   0
            Left            =   360
            TabIndex        =   45
            Top             =   120
            Width           =   1440
         End
      End
      Begin C1Chart2D8.Chart2D chtThis 
         Height          =   2295
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3135
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   5530
         _ExtentY        =   4048
         _StockProps     =   0
         ControlProperties=   "frmPathTrackOut.frx":6DDC
      End
      Begin VB.Label lblMSG 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提示信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.PictureBox picVariation 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   2040
      ScaleHeight     =   2655
      ScaleWidth      =   10935
      TabIndex        =   14
      Top             =   2280
      Width           =   10935
      Begin VB.PictureBox picContrast 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   12735
         TabIndex        =   60
         Top             =   480
         Width           =   12735
         Begin VB.CheckBox chkContrast 
            BackColor       =   &H80000005&
            Caption         =   "与指定期间比较"
            Height          =   255
            Left            =   4530
            TabIndex        =   65
            Top             =   53
            Width           =   1575
         End
         Begin VB.CommandButton cmdContrast 
            Caption         =   "对比(&C)"
            Height          =   345
            Left            =   9600
            TabIndex        =   62
            Top             =   8
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.ComboBox cboYorM 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   30
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpTwo 
            Height          =   300
            Left            =   6570
            TabIndex        =   63
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpOne 
            Height          =   300
            Left            =   1530
            TabIndex        =   64
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpFour 
            Height          =   300
            Left            =   8130
            TabIndex        =   74
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpThree 
            Height          =   300
            Left            =   3090
            TabIndex        =   77
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin VB.Label lblFromToTwo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从               至"
            Height          =   180
            Left            =   6330
            TabIndex        =   76
            Top             =   90
            Width           =   1710
         End
         Begin VB.Label lblFromToOne 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从               至"
            Height          =   180
            Left            =   1290
            TabIndex        =   75
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
         TabIndex        =   53
         Top             =   480
         Width           =   7695
         Begin VB.ComboBox cboTimeType 
            Height          =   300
            Left            =   0
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   56
            Text            =   "cboTimeType"
            Top             =   30
            Width           =   1095
         End
         Begin VB.ComboBox cboPathTime 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   30
            Width           =   1245
         End
         Begin VB.CommandButton cmdVariation 
            Caption         =   "统计(&T)"
            Height          =   345
            Left            =   5880
            TabIndex        =   54
            Top             =   0
            Visible         =   0   'False
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   4560
            TabIndex        =   57
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpStart 
            Height          =   300
            Left            =   2790
            TabIndex        =   58
            Top             =   30
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin VB.Label lblFromTo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从                 至"
            Height          =   180
            Left            =   2460
            TabIndex        =   59
            Top             =   90
            Width           =   1890
         End
      End
      Begin VB.ComboBox cboPathEdition 
         Height          =   300
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   60
         Visible         =   0   'False
         Width           =   1040
      End
      Begin VB.Frame fraGroupLine 
         Height          =   30
         Left            =   120
         TabIndex        =   29
         Top             =   405
         Width           =   5175
      End
      Begin VB.OptionButton optThisPath 
         BackColor       =   &H80000005&
         Caption         =   "当前路径"
         Height          =   180
         Left            =   1230
         TabIndex        =   27
         Top             =   120
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton optAllPath 
         BackColor       =   &H80000005&
         Caption         =   "全院路径"
         Height          =   180
         Left            =   2400
         TabIndex        =   28
         Top             =   120
         Width           =   1200
      End
      Begin XtremeSuiteControls.TabControl tbcVariation 
         Height          =   1095
         Left            =   600
         TabIndex        =   36
         Top             =   1200
         Width           =   3615
         _Version        =   589884
         _ExtentX        =   6376
         _ExtentY        =   1931
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
         TabIndex        =   66
         Top             =   480
         Width           =   11800
         Begin VB.OptionButton optOut 
            BackColor       =   &H80000005&
            Caption         =   "非路径病人"
            Height          =   180
            Left            =   9210
            TabIndex        =   73
            Top             =   90
            Width           =   1200
         End
         Begin VB.OptionButton optIn 
            BackColor       =   &H80000005&
            Caption         =   "路径病人"
            Height          =   180
            Left            =   8160
            TabIndex        =   72
            Top             =   90
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.ComboBox cboInterval 
            Height          =   300
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   30
            Width           =   1095
         End
         Begin VB.ComboBox cboTrendTime 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   30
            Width           =   1095
         End
         Begin VB.CommandButton cmdTrend 
            Caption         =   "查询(&Q)"
            Height          =   335
            Left            =   5520
            TabIndex        =   67
            Top             =   0
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtpTrendStart 
            Height          =   300
            Left            =   2160
            TabIndex        =   69
            Top             =   30
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin VB.Label lblTrend 
            BackColor       =   &H80000005&
            Caption         =   "开始时间                   之后"
            Height          =   255
            Left            =   1380
            TabIndex        =   70
            Top             =   85
            Width           =   3975
         End
      End
      Begin VB.Label lblPathEdition 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径版本"
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
         Left            =   3780
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblZY 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "注："
         Height          =   660
         Left            =   360
         TabIndex        =   37
         Top             =   6360
         Width           =   360
      End
      Begin VB.Label lblPathType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "路径种类"
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
         TabIndex        =   26
         Top             =   120
         Width           =   780
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
      _Version        =   589884
      _ExtentX        =   3413
      _ExtentY        =   3201
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
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   10200
      Width           =   14370
      _ExtentX        =   25347
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPathTrackOut.frx":743B
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11192
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11192
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
   Begin MSComctlLib.ImageList img16 
      Left            =   720
      Top             =   0
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
            Picture         =   "frmPathTrackOut.frx":7CCD
            Key             =   "Path"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrackOut.frx":8267
            Key             =   "PatiMan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrackOut.frx":8801
            Key             =   "PatiWoman"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrackOut.frx":8D9B
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrackOut.frx":9335
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPathTrackOut.frx":98CF
            Key             =   "单病种"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   2640
      ScaleHeight     =   4335
      ScaleWidth      =   10575
      TabIndex        =   5
      Top             =   5640
      Width           =   10575
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   720
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1440
         _Version        =   589884
         _ExtentX        =   2540
         _ExtentY        =   1270
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
         Height          =   1425
         Left            =   0
         ScaleHeight     =   1425
         ScaleWidth      =   10545
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   0
         Width           =   10545
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "变异退出"
            Height          =   180
            Index           =   3
            Left            =   4920
            TabIndex        =   50
            Top             =   495
            Width           =   1020
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "不符合"
            Height          =   180
            Index           =   0
            Left            =   1245
            TabIndex        =   13
            Top             =   495
            Width           =   840
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "执行中"
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
            Caption         =   "正常结束"
            Height          =   180
            Index           =   2
            Left            =   2880
            TabIndex        =   17
            Top             =   495
            Width           =   1020
         End
         Begin VB.OptionButton optState 
            BackColor       =   &H80000005&
            Caption         =   "变异结束"
            Height          =   180
            Index           =   4
            Left            =   3885
            TabIndex        =   19
            Top             =   495
            Width           =   1020
         End
         Begin VB.ComboBox cboForDate 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   120
            Locked          =   -1  'True
            Style           =   1  'Simple Combo
            TabIndex        =   23
            Text            =   "cboForDate"
            Top             =   915
            Width           =   1095
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   915
            Width           =   1245
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "查找(&F)"
            Height          =   350
            Left            =   9360
            TabIndex        =   34
            Top             =   885
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
            TabIndex        =   33
            Top             =   915
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   3000
            TabIndex        =   32
            Top             =   915
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   165019651
            CurrentDate     =   40256
         End
         Begin VB.Label lblState 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "按路径状态"
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
            TabIndex        =   11
            Top             =   495
            Width           =   975
         End
         Begin VB.Label lblFrom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "从                 至"
            Height          =   180
            Left            =   2670
            TabIndex        =   31
            Top             =   975
            Width           =   1890
         End
         Begin VB.Label lblPerson 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "按病人查询"
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
            TabIndex        =   9
            Top             =   135
            Width           =   975
         End
      End
      Begin VB.Frame fraUD 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   45
         Left            =   3000
         MousePointer    =   7  'Size N S
         TabIndex        =   22
         Top             =   2640
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
         ScaleWidth      =   5775
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2520
         Width           =   5775
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
            Caption         =   "未生成原因汇总表"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Visible         =   0   'False
            Width           =   5880
         End
         Begin VB.Label lblMerges 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "合并路径："
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
            TabIndex        =   51
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblDiagInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   1200
            TabIndex        =   21
            Top             =   120
            Width           =   90
         End
         Begin VB.Label lblDiag 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "主要诊断"
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
            TabIndex        =   20
            Top             =   120
            Width           =   780
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTemp 
         Height          =   660
         Left            =   1800
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1080
         _cx             =   1905
         _cy             =   1164
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   78
      Top             =   0
      Visible         =   0   'False
      Width           =   240
      Begin VB.Image imgFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Left            =   0
         OLEDropMode     =   1  'Manual
         Picture         =   "frmPathTrackOut.frx":10131
         Top             =   0
         Width           =   240
      End
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPathTrackOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mlngModul As Long
Private mblnShowStoped As Boolean
Private mblnFirstLoad As Boolean '判断是否是第一次加载
Private mrsTmp As Recordset
Private mstrFlag As String     '标记当前选中表格

Private Enum COL_PATH_LIST
    COL_ID = 0
    COL_图标 = 1
    COL_分类 = 2
    COL_编码 = 3
    COL_名称 = 4
    COL_适用性别 = 5
    COL_适用年龄 = 6
    COL_说明 = 7
    COL_通用 = 8
    COL_报表期间 = 9
    COL_诊断编码 = 10
    COL_诊断名称 = 11
    COL_疾病编码 = 12
    COL_疾病名称 = 13
End Enum

Private Enum COL_PATI_LIST
    col_打印 = 0
    COL_病人ID = 1
    COL_挂号ID = 2
    COL_科室 = 3
    COL_姓名 = 4
    COL_性别 = 5
    COL_年龄 = 6
    COL_门诊号 = 7
    COl_进度 = 8
    COl_标准治疗时间 = 9
    COL_标准费用 = 10
    COL_版本号 = 11
    COL_导入人 = 12
    COl_导入时间 = 13
    COL_结束时间 = 14
    COL_科室ID = 15
    COL_病人状态 = 16
    COL_数据转出 = 17
    COL_不符合原因 = 18
    COL_变异退出原因 = 19
    COL_打印人 = 20
    COL_打印时间 = 21
    COL_患者版打印 = 22
    col_病人路径ID = 23
End Enum

Private Enum CONST_AREA
    Area_Cross = 0
    Area_Category = 1
    Area_Step = 2
    Area_Item = 3
End Enum

Private Enum COL_OPER_LIST
    COL_记录ID = 0
    COL_手术名称 = 1
    COL_手术日期 = 2
    COL_主刀医师 = 3
    COL_麻醉医师 = 4
End Enum

Private Enum VSG_Info
    vsg_原因 = 0
    vsg_项目 = 1
    VSG_明细 = 2
End Enum

Private Enum COL_VSG_Info
    VCol_分类 = 0
    VCol_原因 = 1
    VCol_阶段 = 0
    VCOL_科室 = 0
    VCol_姓名 = 0
    VCol_原因例数 = 2
    VCol_名称 = 1
    VCOL_门诊号 = 1
    VCol_项目例数 = 2
    VCOL_医生 = 2
    VCol_未使用原因 = 3
    VCol_生成时间 = 4
    VCOL_医生姓名 = 1
    VCOL_病人数 = 2
    vcol_入径人数 = 3
    vcol_入径率 = 4
    vcol_变异退出数 = 5
    vcol_变异退出率 = 6
    vcol_变异完成数 = 7
    vcol_变异完成率 = 8
    VCOL_医嘱符合度 = 9
    VCOL_指标 = 0
    VCOL_同期一 = 1
    VCOL_同期二 = 2
    VCOL_差值 = 3
End Enum

Private Const conMenu_View_FindName = 7211                 '*按路径名称查找(&F)
Private Const conMenu_View_FindIll = 7212                 '*按疾病诊断查找(&F)
Private mlng病人ID As Long, mlng挂号ID As Long, mlng病人路径ID As Long
Private mlngVariation As Long, mlngSurvey As Long, mlngTrend As Long
Private mblnIsPathTo As Boolean
Private mblnIsEdition As Boolean
Private mlngOldPathID As Long      '上一次查询的路径id
Private mdateOldStart As Date      '上一次的开始时间
Private mdateOldEnd As Date       '上一次的结束时间
Private mstrDateType As String     '上一次的时间类型
Private mlng路径ID As Long   '上次选择的路径ID

Private Sub cboForDate_Click()
    If cboTime.ListIndex <> cboTime.ListCount - 1 Then
        Call rptPath_SelectionChanged
    End If
End Sub

Private Sub cboPathEdition_Click()
    mblnIsEdition = True
    If tbcSub.Selected.Tag <> "病人路径" Then
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
        Case 0 '今日
            dtpStart.Value = Format(curDate, "yyyy-MM-dd")
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
        Case 1 '最近一周
            dtpStart.Value = DateAdd("ww", -1, curDate)
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
        Case 2 '最近一月
            dtpStart.Value = DateAdd("m", -1, curDate)
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
        Case 3 '最近一季
            dtpStart.Value = DateAdd("q", -1, curDate)
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
        Case 4 '最近半年
            dtpStart.Value = DateAdd("m", -6, curDate)
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
        Case 5 '最近一年
            dtpStart.Value = DateAdd("yyyy", -1, curDate)
            dtpEnd.Value = Format(curDate, "yyyy-MM-dd")
        Case 6 '指  定
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
        Case 0 '今日
            dtpTime(0).Value = Format(curDate, "yyyy-MM-dd")
            dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
        Case 1 '最近一周
            dtpTime(0).Value = DateAdd("ww", -1, curDate)
            dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
        Case 2 '最近一月
            dtpTime(0).Value = DateAdd("m", -1, curDate)
            dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
        Case 3 '最近一季
            dtpTime(0).Value = DateAdd("q", -1, curDate)
            dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
        Case 4 '最近半年
            dtpTime(0).Value = DateAdd("m", -6, curDate)
            dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
        Case 5 '最近一年
            dtpTime(0).Value = DateAdd("yyyy", -1, curDate)
            dtpTime(1).Value = Format(curDate, "yyyy-MM-dd")
        Case 6 '指  定
            dtpTime(0).SetFocus
            cmdFind.Visible = True
    End Select
    
    If cboTime.ListIndex <> cboTime.ListCount - 1 Then
        Call rptPath_SelectionChanged
    End If
End Sub

Private Sub cboTimeType_Click()
    If tbcSub.Selected.Tag <> "病人路径" Then
        Call tbcVariation_SelectedChanged(tbcVariation.Selected)
    End If
End Sub

Private Sub cboTrendTime_Click()
    cboInterval.Clear
    If cboTrendTime.ListIndex = 0 Then
        '按月
        cboInterval.AddItem "一周"
        cboInterval.AddItem "一月"
        cboInterval.AddItem "两月"
        cboInterval.AddItem "一季度"
        dtpTrendStart.CustomFormat = "yyyy年MM月dd日"
    Else
        cboInterval.AddItem "半年"
        cboInterval.AddItem "一年"
        cboInterval.AddItem "两年"
        cboInterval.AddItem "三年"
        dtpTrendStart.CustomFormat = "yyyy年MM月"
    End If
    cboInterval.ListIndex = 1
End Sub

Private Sub cboYorM_Click()
    If cboYorM.ListIndex = 0 Then
        dtpOne.CustomFormat = "yyyy年MM月"
        dtpTwo.CustomFormat = "yyyy年MM月"
    ElseIf cboYorM.ListIndex = 1 Then
        dtpOne.CustomFormat = "yyyy年MM月"
        dtpTwo.CustomFormat = "yyyy年MM月"
        dtpThree.CustomFormat = "yyyy年MM月"
        dtpFour.CustomFormat = "yyyy年MM月"
    ElseIf cboYorM.ListIndex = 2 Then
        dtpOne.CustomFormat = "yyyy年"
        dtpTwo.CustomFormat = "yyyy年"
    End If
    If tbcSub.Selected.Tag <> "病人路径" Then
        Call tbcVariation_SelectedChanged(tbcVariation.Selected)
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim objRow As ReportRow, i As Long
    Dim lng路径ID As Long
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
        Case conMenu_View_Show '查看路径表
            Call FuncShowPath
        Case conMenu_Edit_OutLogView
            If rptPath.SelectedRows.count > 0 Then
                If Not rptPath.SelectedRows(0).GroupRow Then
                    lng路径ID = rptPath.SelectedRows(0).Record(COL_ID).Value
                End If
            End If
            Call frmPathOutLogOut.ShowMe(Me, mlng病人ID, mlng挂号ID, 1, Nothing, lng路径ID, mlng病人路径ID)
        Case conMenu_View_ShowStoped
            mblnShowStoped = Not mblnShowStoped
            Control.Checked = mblnShowStoped
            Call LoadPathList
        Case conMenu_View_Find '查找
            If Me.ActiveControl Is txtFind Then
                txtFind.SetFocus '有时需要定位一下
                If txtFind.Text <> "" Then
                    Call FuncFindPath
                End If
            Else
                txtFind.SetFocus
            End If
        Case conMenu_View_FindNext '查找下一个
            If txtFind.Text = "" Then
                txtFind.SetFocus
            Else
                Call FuncFindPath(True)
            End If
        Case conMenu_View_ToolBar_Button '工具栏
            For i = 2 To cbsMain.count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '按钮文字
            For i = 2 To cbsMain.count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '大图标
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '状态栏
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
        Case conMenu_View_Expend_CurCollapse '折叠当前组
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    rptPath.SelectedRows(0).Expanded = False
                ElseIf Not rptPath.SelectedRows(0).ParentRow Is Nothing Then
                    If rptPath.SelectedRows(0).ParentRow.GroupRow Then
                        rptPath.SelectedRows(0).ParentRow.Expanded = False
                    End If
                End If
            End If
            '因折叠定位到分组上,不会自动激活该事件
            Call rptPath_SelectionChanged
        Case conMenu_View_Expend_CurExpend '展开当前组
            If rptPath.SelectedRows.count > 0 Then
                rptPath.SelectedRows(0).Expanded = True
            End If
        Case conMenu_View_Expend_AllCollapse '折叠所有组
            For Each objRow In rptPath.Rows
                If objRow.GroupRow Then objRow.Expanded = False
            Next
            '因折叠定位到分组上,不会自动激活该事件
            Call rptPath_SelectionChanged
        Case conMenu_View_Expend_AllExpend '展开所有组
            For Each objRow In rptPath.Rows
                If objRow.GroupRow Then objRow.Expanded = True
            Next
        Case conMenu_View_Refresh '刷新
            Call LoadPathList
        Case conMenu_Help_Web_Home 'Web上的中联
            Call zlHomePage(Me.Hwnd)
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.Hwnd)
        Case conMenu_Help_Web_Mail '发送反馈
            Call zlMailTo(Me.Hwnd)
        Case conMenu_Help_About '关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
        Case conMenu_View_FindName '按路径名称查找
            Set objPopup = cbsMain.FindControl(, Control.Parent.BarID)
            objPopup.Caption = Control.Caption
        Case conMenu_View_FindIll '按疾病诊断查找
            Set objPopup = cbsMain.FindControl(, Control.Parent.BarID)
            objPopup.Caption = Control.Caption
        Case Else
            If Between(Control.ID, conMenu_ReportPopup * 100# + 1, conMenu_ReportPopup * 100# + 99) And Control.Parameter <> "" Then
                If rptPath.SelectedRows.count > 0 Then
                    If Not rptPath.SelectedRows(0).GroupRow Then
                        lng路径ID = rptPath.SelectedRows(0).Record(COL_ID).Value
                    End If
                End If
                
                '执行发布到当前模块的报表
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, _
                       "路径=" & lng路径ID, "病人ID=" & mlng病人ID, "挂号ID=" & mlng挂号ID)
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
'功能：根据权限设置菜单和工具栏的可见状态
    Dim blnVisible As Boolean, strItem As String

    '权限只需判断一次,已经判断过的命令不用再判断
    If Control.Category = "已判断" Then Exit Sub

    blnVisible = True
    Select Case Control.ID
        Case conMenu_Tool_Report
            If InStr(";" & mstrPrivs & ";", ";单病种统计表;") = 0 Then blnVisible = False
        Case conMenu_Edit_OutLogView
            blnVisible = CheckPathOutLog
    End Select
    
    Control.Visible = blnVisible
    Control.Category = "已判断"
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnEnabled As Boolean
        
    '根据权限设置按钮可见状态
    Call SetControlVisible(Control)
    If Not Control.Visible And Control.ID <> conMenu_File_SaveJpeg Then Exit Sub
        
    Select Case Control.ID
        Case conMenu_View_Show, conMenu_Edit_OutLogView '查看路径表,查看出径登记表
            blnEnabled = mlng病人ID > 0
            
            If Control.ID = conMenu_Edit_OutLogView And blnEnabled Then
                blnEnabled = optState(2).Value Or optState(3).Value Or optState(4).Value
            End If
            Control.Enabled = blnEnabled
        Case conMenu_File_SaveJpeg '保存图片
            Control.Enabled = chtThis.Visible
            Control.Visible = chtThis.Visible
        Case conMenu_View_ToolBar_Button '工具栏
            If cbsMain.count >= 2 Then
                Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text '图标文字
            If cbsMain.count >= 2 Then
                Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '大图标
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_FindNext '查找下一个
            Control.Visible = False
        Case conMenu_View_StatusBar '状态栏
            Control.Checked = Me.stbThis.Visible
        Case conMenu_View_Expend_CurExpend '展开当前组
            blnEnabled = False
            If rptPath.SelectedRows.count > 0 Then
                If rptPath.SelectedRows(0).GroupRow Then
                    blnEnabled = Not rptPath.SelectedRows(0).Expanded
                End If
            End If
            Control.Enabled = blnEnabled
        Case conMenu_View_Expend_CurCollapse '折叠当前组
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
        Case conMenu_View_Expend '折叠/展开组
            Control.Enabled = rptPath.GroupsOrder.count > 0 And rptPath.Rows.count > 0
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
        vsgInfo(vsg_原因).ColHidden(VCOL_同期二) = False
        vsgInfo(vsg_原因).ColHidden(VCOL_差值) = False
    Else
        vsgInfo(vsg_原因).ColHidden(VCOL_同期二) = True
        vsgInfo(vsg_原因).ColHidden(VCOL_差值) = True
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
    Call set总体情况比对(lngPathID)
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
    If tbcSub.Selected.Tag <> "病人路径" Then
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
    mlng病人ID = 0
    mlng挂号ID = 0
    mlng病人路径ID = 0
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
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "病人路径", picPati.Hwnd, 0).Tag = "病人路径"
        .InsertItem(1, "变异分析", picVariation.Hwnd, 0).Tag = "变异分析"
        .InsertItem(2, "概况分析", picVariation.Hwnd, 0).Tag = "概况分析"
        .InsertItem(3, "趋势分析", picVariation.Hwnd, 0).Tag = "趋势分析"
        .Item(0).Selected = True
    End With
    
     With tbcVariation
        With .PaintManager
            .Appearance = xtpTabAppearanceVisualStudio
            .Color = xtpTabColorOffice2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
        End With
        .InsertItem(0, "按医生统计", picReason.Hwnd, 0).Tag = "按医生统计"
        .InsertItem(1, "科室变异率排名", picReason.Hwnd, 0).Tag = "科室变异率排名"
        .InsertItem(2, "未导入原因", picReason.Hwnd, 0).Tag = "未导入原因"
        .InsertItem(3, "未生成原因", picReason.Hwnd, 0).Tag = "未生成原因"
        .InsertItem(4, "路径外项目", picReason.Hwnd, 0).Tag = "路径外项目"
        .InsertItem(5, "时间变异分析", picReason.Hwnd, 0).Tag = "时间变异分析"
        .InsertItem(6, "变异退出分析", picReason.Hwnd, 0).Tag = "变异退出分析"
        .InsertItem(7, "路径完成情况", picReason.Hwnd, 0).Tag = "路径完成情况"
        .InsertItem(8, "阶段平均费用", picReason.Hwnd, 0).Tag = "阶段平均费用"
        .InsertItem(9, "就诊日分布图", picReason.Hwnd, 0).Tag = "就诊日分布图"
        .InsertItem(10, "总体情况", picReason.Hwnd, 0).Tag = "总体情况"
        .InsertItem(11, "平均就诊费用", picReason.Hwnd, 0).Tag = "平均就诊费用"
        .InsertItem(12, "入径率", picReason.Hwnd, 0).Tag = "入径率"
        .InsertItem(13, "完成率", picReason.Hwnd, 0).Tag = "完成率"
        .InsertItem(14, "变异率", picReason.Hwnd, 0).Tag = "变异率"
    End With
    tbcVariation.Item(tbcVariation.ItemCount - 1).Selected = True
    tbcVariation.Item(0).Selected = True
    
    'vsFlexGrid
    '-----------------------------------------------------
    strHead = "病人姓名,1500,1;门诊号,1500,1;医生,1500,1;未使用原因,3200,1;生成时间,3000,1"
    Call InitTable(vsgInfo(VSG_明细), strHead)
    vsgInfo(VSG_明细).ExplorerBar = flexExSortShowAndMove
    
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
        '.UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    
    '---cboTime
    cboTime.AddItem "今    日"
    cboTime.AddItem "最近一周"
    cboTime.AddItem "最近一月"
    cboTime.AddItem "最近一季"
    cboTime.AddItem "最近半年"
    cboTime.AddItem "最近一年"
    cboTime.AddItem "[指  定]"
    cboTime.ListIndex = 2
    
    '---cboPathTime
    cboPathTime.AddItem "今    日"
    cboPathTime.AddItem "最近一周"
    cboPathTime.AddItem "最近一月"
    cboPathTime.AddItem "最近一季"
    cboPathTime.AddItem "最近半年"
    cboPathTime.AddItem "最近一年"
    cboPathTime.AddItem "[指  定]"
    cboPathTime.ListIndex = 2
    
    '---cboForDate
    cboForDate.AddItem "导入时间"
    cboForDate.ListIndex = 0
    
    '---cboTimeType
    cboTimeType.AddItem "导入时间"
    cboTimeType.ListIndex = 0
    
    '---cboYorM
    cboYorM.AddItem "按月"
    cboYorM.AddItem "按季度"
    cboYorM.AddItem "按年"
    cboYorM.ListIndex = 0
    dtpOne.Value = Format(zlDatabase.Currentdate, "yyyy-mm")
    dtpTwo.Value = Format(CDate(Format(dtpOne.Value, "yyyy-mm")) - 1, "yyyy-MM-01")
    
    '---cboTrendTime
    cboTrendTime.AddItem "按天"
    cboTrendTime.AddItem "按月"
    cboTrendTime.ListIndex = 0
    dtpTrendStart.Value = Format(CDate(Format(dtpOne.Value, "yyyy-mm")) - 1, "yyyy-MM-01")
    
    Call RestoreWinState(Me, App.ProductName)
    Call LoadPathList
End Sub

Private Sub InitPathReportColumn()
    Dim objCol As ReportColumn

    With rptPath
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(COL_ID, "", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_图标, "", 18, False)
            objCol.Sortable = False
            objCol.AllowDrag = False
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_分类, "分类", 80, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_编码, "编码", 35, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(COL_名称, "名称", 150, True)
            objCol.Groupable = False
        Set objCol = .Columns.Add(COL_适用性别, "适用性别", 55, True)
            objCol.Alignment = xtpAlignmentCenter
        Set objCol = .Columns.Add(COL_适用年龄, "适用年龄", 55, True)
        Set objCol = .Columns.Add(COL_说明, "", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_通用, "", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_报表期间, "报表期间", 55, True)
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
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的临床路径..."
            '.ShadeGroupHeadings = True
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
        .SetImageList Me.img16
        .GroupsOrder.Add .Columns(COL_分类)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(COL_分类)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_编码)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub InitPatiReportColumn()
    Dim objCol As ReportColumn

    With rptPati
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(col_打印, "打印", 50, True)
        objCol.Sortable = False
        objCol.AllowDrag = False
        objCol.Alignment = xtpAlignmentLeft
        objCol.Editable = True
        objCol.Icon = img16.ListImages("Check").Index - 1
        Set objCol = .Columns.Add(COL_病人ID, "病人ID", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_挂号ID, "挂号ID", 0, False)
            objCol.Visible = False
        
        Set objCol = .Columns.Add(COL_科室, "科室", 70, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_姓名, "姓名", 70, True)
        Set objCol = .Columns.Add(COL_性别, "性别", 45, True)
        Set objCol = .Columns.Add(COL_年龄, "年龄", 60, True)
        Set objCol = .Columns.Add(COL_门诊号, "门诊号", 62, True)
        Set objCol = .Columns.Add(COl_进度, "进度", 40, True)
        Set objCol = .Columns.Add(COl_标准治疗时间, "标准治疗时间", 70, True)
        Set objCol = .Columns.Add(COL_标准费用, "标准费用", 80, True)
        Set objCol = .Columns.Add(COL_版本号, "版本号", 45, True)
        Set objCol = .Columns.Add(COL_导入人, "导入人", 55, True)
        Set objCol = .Columns.Add(COl_导入时间, "导入时间", 106, True)
        Set objCol = .Columns.Add(COL_结束时间, "结束时间", 106, True)
        
        Set objCol = .Columns.Add(COL_科室ID, "科室ID", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_病人状态, "病人状态", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_数据转出, "数据转出", 0, False)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_不符合原因, "不符合原因", 200, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_变异退出原因, "变异退出原因", 200, True)
            objCol.Visible = False
        Set objCol = .Columns.Add(COL_打印人, "打印人", 55, True)
        Set objCol = .Columns.Add(COL_打印时间, "打印时间", 106, True)
        Set objCol = .Columns.Add(COL_患者版打印, "患者版打印", 70, True)
        Set objCol = .Columns.Add(col_病人路径ID, "病人路径ID", 0, False)
        For Each objCol In .Columns
            If objCol.Index <> col_打印 Then
                objCol.Editable = False
            End If
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有临床路径应用的病人数据..."
            '.ShadeGroupHeadings = True
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = True
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
        .SetImageList Me.img16
        .GroupsOrder.Add .Columns(COL_科室)
        .GroupsOrder(0).SortAscending = True '分组之后,如果分组列不显示,分组列的排序是不变的
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(COL_科室)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_门诊号)
        .SortOrder(1).SortAscending = True
    End With
End Sub

Private Sub InitOperReportColumn()
    Dim objCol As ReportColumn

    With rptOper
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)或ItemIndex查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(COL_记录ID, "记录ID", 0, False)
            objCol.Visible = False
            
        Set objCol = .Columns.Add(COL_手术名称, "手术名称", 300, True)
        Set objCol = .Columns.Add(COL_手术日期, "手术日期", 200, True)
        Set objCol = .Columns.Add(COL_主刀医师, "主刀医师", 80, True)
        Set objCol = .Columns.Add(COL_麻醉医师, "麻醉医师", 80, True)
        
        For Each objCol In .Columns
            objCol.Editable = False
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有病人手术信息..."
            '.ShadeGroupHeadings = True
        End With
        .AutoColumnSizing = False
        .AllowColumnRemove = False
        .ShowGroupBox = False
        .ShowItemsInGroups = False
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
        .SetImageList Me.img16
    End With
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim lngCount As Long
    
    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)…")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveJpeg, "保存为图片(&J)")
            objControl.IconId = 8104
        Set objControl = .Add(xtpControlButton, conMenu_File_BatPrint, "批量打印")
            objControl.IconId = 8128
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
            objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "工具栏(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "大图标(&B)", -1, False
        End With
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_Expend, "展开/折叠组(&X)"):
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllCollapse, "折叠所有组(&L)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_AllExpend, "展开所有组(&X)", -1, False)
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurCollapse, "折叠当前组(&C)", -1, False)
                objControl.BeginGroup = True
            Set objControl = .Add(xtpControlButton, conMenu_View_Expend_CurExpend, "展开当前组(&E)", -1, False)
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "查看路径表(&P)")
            objControl.IconId = 126401202
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "查看出径登记表(&O)")
            objControl.IconId = 3032
        Set objControl = .Add(xtpControlButton, conMenu_View_ShowStoped, "显示已停用的路径表(&D)")
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)")
            objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB上的" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "主页(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "论坛(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)…")
            objControl.BeginGroup = True
    End With

    '主菜单右侧的查找
    With cbsMain.ActiveMenuBar.Controls
        Set objPopup = .Add(xtpControlSplitButtonPopup, 0, "按名称查找")
        objPopup.ID = 0
        objPopup.Style = xtpButtonIconAndCaption
        objPopup.IconId = conMenu_View_Find
        objPopup.Flags = xtpFlagRightAlign
        With objPopup.CommandBar.Controls
            Set objControl = .Add(xtpControlButton, conMenu_View_FindName, "按名称查找")
            Set objControl = .Add(xtpControlButton, conMenu_View_FindIll, "按诊断查找")
        End With
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "")
        objCustom.Handle = txtFind.Hwnd
        objCustom.Flags = xtpFlagRightAlign
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Print, "打印")
        Set objControl = .Add(xtpControlButton, conMenu_File_Preview, "预览")
        Set objControl = .Add(xtpControlButton, conMenu_File_BatPrint, "批量打印")
            objControl.IconId = 3903
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_SaveJpeg, "保存为图片")
            objControl.IconId = 8104
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_Show, "路径表")
            objControl.IconId = 126401202
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Edit_OutLogView, "出径登记表")
            objControl.IconId = 3032
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    '设置一些公共的热键绑定
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyF, conMenu_View_Find '查找
        .Add 0, vbKeyF3, conMenu_View_FindNext '查找下一个
        .Add FCONTROL, vbKeyAdd, conMenu_View_Expend_AllExpend '展开所有组
        .Add FCONTROL, vbKeySubtract, conMenu_View_Expend_AllCollapse '折叠所有组
        .Add FCONTROL, vbKeyP, conMenu_File_Print '打印
        .Add FALT, vbKeyJ, conMenu_File_SaveJpeg   '保存为图片
        .Add 0, vbKeyF5, conMenu_View_Refresh '刷新
        .Add 0, vbKeyF1, conMenu_Help_Help '帮助
    End With

    '恢复及固定的一些菜单设置
    cbsMain.ActiveMenuBar.Title = "菜单"
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
    
    '读取发布到该模块的报表(不含虚拟模块的)
    '-----------------------------------------------------
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModul, mstrPrivs)
End Sub

Private Function LoadPathList(Optional ByVal str分类 As String, Optional ByVal str编码 As String) As Boolean
'功能：根据当前设置的条件读取门诊路径目录数据
'参数：用于定位
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strSQLTmp As String
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim objRow As ReportRow, i As Long
    Dim lngPreID As Long, lngPreIdx As Long
    
    Screen.MousePointer = 11
        
    On Error GoTo errH
    
    'SQL中不排序提高效率,ReportControl有排序处理
    strSQLTmp = "Select  " & vbNewLine & _
                " ID, 分类, 编码, 名称, 适用性别, 适用年龄, 说明, 通用, 期间, f_List2str(Cast(Collect(疾病编码) As t_Strlist)) As 疾病编码," & vbNewLine & _
                " f_List2str(Cast(Collect(疾病名称) As t_Strlist)) As 疾病名称, f_List2str(Cast(Collect(诊断编码) As t_Strlist)) As 诊断编码," & vbNewLine & _
                " f_List2str(Cast(Collect(诊断名称) As t_Strlist)) As 诊断名称" & vbNewLine & _
                "From ("

    strSql = "Select a.*, e.编码 As 疾病编码, e.名称 As 疾病名称, d.编码 As 诊断编码, d.名称 As 诊断名称" & vbNewLine & _
            "       From (Select a.Id, a.分类, a.编码, a.名称, a.适用性别, a.适用年龄, a.说明, a.通用, b.期间," & vbNewLine & _
            "                     Row_Number() Over(Partition By a.Id Order By b.期间 Desc) As Top" & vbNewLine & _
            "              From 门诊路径目录 A, 门诊路径报表文件 B" & vbNewLine & _
            "              Where a.Id = b.路径id(+) And b.报表id(+) = 1 And Exists" & vbNewLine & _
            "               (Select 路径id From 门诊路径版本 C Where a.Id = c.路径id And 审核人 Is Not Null" & _
            IIf(mblnShowStoped, "", " And 停用人 Is Null") & ")) A, 门诊路径病种 C, 疾病诊断目录 D," & vbNewLine & _
            "            疾病编码目录 E" & vbNewLine & _
            "       Where c.路径id = a.Id And d.Id(+) = c.诊断id And e.Id(+) = c.疾病id And Top = 1"

    If InStr(mstrPrivs, "全院路径") = 0 Then
        '没有权限时，只能对只应用于本科的路径进行处理
        strSql = strSql & _
            " And 通用=2 And Not Exists(" & _
                " Select 科室ID From 门诊路径科室 Where 路径ID=a.ID" & _
                " Minus Select 部门ID From 部门人员 Where 人员ID=[1])"
        optThisPath.Value = True
        optAllPath.Enabled = False
        optThisPath.Enabled = False
    End If
    strSql = strSQLTmp & strSql & ")" & vbNewLine & _
    " Group By ID, 分类, 编码, 名称, 适用性别, 适用年龄, 说明, 通用, 期间"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
        
    '记录现在选中的反馈
    If rptPath.SelectedRows.count > 0 Then
        If Not rptPath.SelectedRows(0).GroupRow Then
            lngPreIdx = rptPath.SelectedRows(0).Index '用于快速重新定位
            lngPreID = rptPath.SelectedRows(0).Record(COL_ID).Value
        End If
    End If
    
    rptPath.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPath.Records.Add()
        Set objItem = objRecord.AddItem(Val(rsTmp!ID))
        Set objItem = objRecord.AddItem("")
            objItem.Icon = img16.ListImages("Path").Index - 1
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!分类, "<未指定分类>")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!编码)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!名称)))
        Set objItem = objRecord.AddItem(CStr(Decode(NVL(rsTmp!适用性别, 0), 0, "", 1, "男", 2, "女")))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!适用年龄)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!说明)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!通用, 1)))
        Set objItem = objRecord.AddItem("" & rsTmp!期间)
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!疾病编码)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!疾病名称)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!诊断编码)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!诊断名称)))
        
        rsTmp.MoveNext
    Loop
    rptPath.Populate
    
    If rptPath.Rows.count = 0 Then
        Call ClearSubData
    Else
        If str分类 <> "" And str编码 <> "" Then
            For i = 0 To rptPath.Rows.count - 1
                If Not rptPath.Rows(i).GroupRow Then
                    If rptPath.Rows(i).Record(COL_分类).Value = str分类 _
                        And rptPath.Rows(i).Record(COL_编码).Value = str编码 Then
                        Set objRow = rptPath.Rows(i): Exit For
                    End If
                End If
            Next
        Else
            If lngPreID <> 0 Then
                '先快速定位
                If lngPreIdx <= rptPath.Rows.count - 1 Then
                    If Not rptPath.Rows(lngPreIdx).GroupRow Then
                        If rptPath.Rows(lngPreIdx).Record(COL_ID).Value = lngPreID Then
                            Set objRow = rptPath.Rows(lngPreIdx)
                        End If
                    End If
                End If
                '再进行查找
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
            '取第一个非分组行
            If objRow Is Nothing Then
                For i = 0 To rptPath.Rows.count - 1
                    If Not rptPath.Rows(i).GroupRow Then Set objRow = rptPath.Rows(i): Exit For
                Next
            End If
        End If
        
        Set rptPath.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
        Me.stbThis.Panels(2).Text = "共有 " & rptPath.Records.count & " 个临床路径"
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
End Sub

Private Sub fraFlag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If fraFlag.Visible Then
       Call zlCommFun.ShowTipInfo(fraFlag.Hwnd, "允许该表格内容预览、打印、输出到EXCEL", True)
    End If
End Sub

Private Sub fraGroupLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If vsgInfo(vsg_原因).Width + X < 2000 Or vsgInfo(vsg_项目).Width - X < 2000 Then Exit Sub
        fraGroupLR.Left = fraGroupLR.Left + X
        vsgInfo(vsg_原因).Width = vsgInfo(vsg_原因).Width + X
        vsgInfo(vsg_项目).Width = vsgInfo(vsg_项目).Width - X
        vsgInfo(vsg_项目).Left = vsgInfo(vsg_项目).Left + X
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
        If vsgInfo(vsg_原因).Height + Y < 1740 Or vsgInfo(vsg_原因).Height - Y < 1000 Then Exit Sub
        If vsgInfo(VSG_明细).Height + Y < 1000 Or vsgInfo(VSG_明细).Height - Y < 1740 Then Exit Sub

        fraGroupUD.Top = fraGroupUD.Top + Y
        fraGroupLR.Height = fraGroupLR.Height + Y
        vsgInfo(vsg_原因).Height = vsgInfo(vsg_原因).Height + Y
        vsgInfo(vsg_项目).Height = vsgInfo(vsg_项目).Height + Y
        vsgInfo(VSG_明细).Top = vsgInfo(VSG_明细).Top + Y
        vsgInfo(VSG_明细).Height = vsgInfo(VSG_明细).Height - Y
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
    zlCommFun.ShowTipInfo picFilter.Hwnd, "设定条件后，请执行刷新读取数据(F5)"
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
        lblMsg.Top = 50
        lblMsg.Left = 50
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
    With vsgInfo(vsg_原因)
        For i = 0 To .Cols - 1
            lngWidth = lngWidth + .ColWidth(i)
        Next
        lblInfo(0).Move 50, 50
        .Move 0, lblInfo(0).Top + lblInfo(0).Height, lngWidth + 100, IIf(vsgInfo(VSG_明细).Visible Or tbcVariation.Selected.Tag = "未生成原因", picTable.Height / 2, picTable.Height) - lblInfo(0).Top - lblInfo(0).Height
        fraGroupLR.Move .Width, 0, fraGroupLR.Width, .Height + lblInfo(0).Top + lblInfo(0).Height
        If vsgInfo(vsg_项目).Visible = False Then vsgInfo(vsg_原因).Width = picTable.Width
    End With
    
    With vsgInfo(vsg_项目)
        lblInfo(1).Move fraGroupLR.Left + fraGroupLR.Width + 50, 50
        txtFindNum.Move lblInfo(1).Left + lblInfo(1).Width - 950, lblInfo(1).Top - 30
        .Move vsgInfo(vsg_原因).Width + fraGroupLR.Width, vsgInfo(vsg_原因).Top, picTable.Width - vsgInfo(vsg_原因).Width - fraGroupLR.Width, vsgInfo(vsg_原因).Height
        If Not vsgInfo(VSG_明细).Visible Then Exit Sub
        fraGroupUD.Move 0, vsgInfo(vsg_原因).Height + vsgInfo(vsg_原因).Top, picTable.Width
    End With
    
    With vsgInfo(VSG_明细)
        lblInfo(2).Move 50, fraGroupUD.Top + fraGroupUD.Height + 50
        imgFrom.Move vsgInfo(vsg_项目).Left + vsgInfo(vsg_项目).Width / 2, lblInfo(2).Top - 50
        .Move 0, lblInfo(2).Top + lblInfo(2).Height, picTable.Width, picTable.Height - lblInfo(2).Top - lblInfo(2).Height
        .ColWidth(VCol_未使用原因) = .Width / 2.88
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

Private Sub rptPath_SelectionChanged()
    If rptPath.Tag = "1" Then Exit Sub
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
    If tbcSub.Selected.Tag <> "病人路径" Then
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
    mlng病人ID = 0
    mlng挂号ID = 0
    mlng病人路径ID = 0
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
    
    '如果点击表头的图片，就选中全部
    If Button = 1 Then
        If rptPati.HitTest(X, Y).ht = xtpHitTestHeader Then
            Set objColumn = rptPati.HitTest(X, Y).Column
            If Not objColumn Is Nothing Then
                If objColumn.Index = col_打印 Then
                    If rptPati.Columns(col_打印).Icon = img16.ListImages("Check").Index - 1 Then
                        rptPati.Columns(col_打印).Icon = img16.ListImages("UnCheck").Index - 1
                        For i = 0 To rptPati.Records.count - 1
                            rptPati.Records(i)(col_打印).Checked = True
                        Next
                    Else
                        rptPati.Columns(col_打印).Icon = img16.ListImages("Check").Index - 1
                        For i = 0 To rptPati.Records.count - 1
                            rptPati.Records(i)(col_打印).Checked = False
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
    
    mlng病人ID = 0
    mlng挂号ID = 0
    mlng病人路径ID = 0
    
    If Me.Visible Then
        Call SetFlagBySelectedTable(True, "RPTPATI")
    End If
    If rptPati.FocusedRow Is Nothing Then Exit Sub
    If rptPati.FocusedRow.GroupRow Then Exit Sub
    cbsMain_Resize
    
    mlng病人ID = Val(rptPati.FocusedRow.Record(COL_病人ID).Value)
    mlng挂号ID = Val(rptPati.FocusedRow.Record(COL_挂号ID).Value)
    mlng病人路径ID = Val(rptPati.FocusedRow.Record(col_病人路径ID).Value)
    
    Set rsTmp = Get病种ID(mlng病人ID, mlng挂号ID)
    If rsTmp.RecordCount > 0 Then
        lblDiagInfo.Caption = "" & rsTmp!诊断描述
    End If
    
    picInfo.Height = rptOper.Height + IIf(lblMerge.Visible, lblMerge.Height + lblMerge.Top, lblDiag.Height + lblDiag.Top) + 100
    Call LoadOperInfo(mlng病人ID, mlng挂号ID)
    picInfo.Visible = rptPati.Rows.count And Not rptPati.FocusedRow Is Nothing
 
    Call picPati_Resize
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim i As Long

    If Item.Tag = "病人路径" Then
        cboForDate.ListIndex = cboTimeType.ListIndex
        cboTime.ListIndex = cboPathTime.ListIndex
        dtpTime(0).Value = dtpStart.Value
        dtpTime(1).Value = dtpEnd.Value
        mblnIsPathTo = True
        Me.stbThis.Panels(3).Visible = True
    ElseIf Item.Tag = "变异分析" Then
        '判断上次是否是病人路径页面转过来的
        If mblnIsPathTo Then
            cboTimeType.ListIndex = cboForDate.ListIndex
            cboPathTime.ListIndex = cboTime.ListIndex
            dtpStart.Value = dtpTime(0).Value
            dtpEnd.Value = dtpTime(1).Value
        End If
        mblnIsPathTo = False
        '隐藏概况分析的选项卡，显示变异原因的选项卡
        For i = 0 To tbcVariation.ItemCount - 1
            If i <= 7 Then
                tbcVariation.Item(i).Visible = True
            Else
                tbcVariation.Item(i).Visible = False
            End If
        Next
        '跳到上次最后浏览的选项卡
        If mlngVariation <= 7 Then
            tbcVariation.Item(mlngVariation).Selected = True
        Else
            tbcVariation.Item(0).Selected = True
        End If
        Me.stbThis.Panels(3).Visible = False
    ElseIf Item.Tag = "概况分析" Then
        If mblnIsPathTo Then
            cboTimeType.ListIndex = cboForDate.ListIndex
            cboPathTime.ListIndex = cboTime.ListIndex
            dtpStart.Value = dtpTime(0).Value
            dtpEnd.Value = dtpTime(1).Value
        End If
        mblnIsPathTo = False
        '显示概况分析的选项卡，隐藏变异原因的选项卡
        For i = 0 To tbcVariation.ItemCount - 1
            If i <= 7 Or i > 11 Then
                tbcVariation.Item(i).Visible = False
            Else
                tbcVariation.Item(i).Visible = True
            End If
        Next
        '跳到上次最后浏览的选项卡
        If mlngSurvey <= 7 Then
            tbcVariation.Item(8).Selected = True
        Else
            tbcVariation.Item(mlngSurvey).Selected = True
        End If
        Me.stbThis.Panels(3).Visible = False
    ElseIf Item.Tag = "趋势分析" Then
        If mblnIsPathTo Then
            cboTimeType.ListIndex = cboForDate.ListIndex
            cboPathTime.ListIndex = cboTime.ListIndex
            dtpStart.Value = dtpTime(0).Value
            dtpEnd.Value = dtpTime(1).Value
        End If
        mblnIsPathTo = False
        '显示概况分析的选项卡，隐藏变异原因的选项卡
        For i = 0 To tbcVariation.ItemCount - 1
            If i <= 11 Then
                tbcVariation.Item(i).Visible = False
            Else
                tbcVariation.Item(i).Visible = True
            End If
        Next
        '跳到上次最后浏览的选项卡
        If mlngTrend <= 11 Then
            tbcVariation.Item(12).Selected = True
        Else
            tbcVariation.Item(mlngTrend).Selected = True
        End If
        Me.stbThis.Panels(3).Visible = False
    End If
    If Me.Visible And InStr(";病人路径;变异分析;", ";" & Item.Tag & ";") > 0 Then
        If rptPati.Records.count > 0 And Item.Tag = "病人路径" Then
        
            Call SetFlagBySelectedTable(True, "RPTPATI")
        ElseIf Item.Tag = "变异分析" Then
            Call SetFlagBySelectedTable(True, "VSGINFO_0")
        End If
    End If
End Sub

Private Sub Set未导入原因(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        lblMsg.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypePie
        .ChartArea.Border.Width = 4
        '未导入原因
        .Header.Text = "未导入原因分布图"
        .Header.Interior.ForegroundColor = vbBlack
        '.ChartArea.Pie.StartAngle = 90
        strSql = "Select b.编码, b.名称, Count(1) As 未导入数量, 100 * Round(Count(1) / Sum(Count(1)) Over(), 4) 比例" & vbNewLine & _
                "From 病人门诊路径 A, 门诊变异常见原因 B " & vbNewLine & _
                "Where a.未导入原因 = b.编码 And b.性质 = 0"
        strSql = strSql & IIf(optAllPath.Value, "", " And a.路径id=[1]")
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.编码, b.名称"
    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.NumSeries = rsTmp.RecordCount
        If .ChartGroups.Item(1).Data.NumSeries <> 0 Then .ChartGroups.Item(1).Data.NumPoints(1) = 1
        i = 1
        Do Until rsTmp.EOF
            .ChartGroups.Item(1).SeriesLabels.Add (rsTmp!名称 & Space(2) & "共" & rsTmp!未导入数量 & "人(" & Val(rsTmp!比例 & "") & "%)")
            .ChartGroups.Item(1).Data.Y(i, 1) = Val(rsTmp!比例 & "")
            rsTmp.MoveNext
            i = i + 1
        Loop
        '注意信息
        lblZY.Caption = "注：该图的计数规则是一个病人的一次就诊（每次就诊发生变异则为一次）。" & vbCrLf & _
                        "其中：没有使用过的未导入原因不显示出来。"
        If rsTmp.RecordCount = 0 Then
            lblMsg.Caption = "您所指定的时间范围内未发现未导入的病人。"
            lblMsg.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        '保存上次浏览的图
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Set变异退出分析(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        lblMsg.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        '变异退出分析
        .ChartGroups.Item(1).ChartType = oc2dTypePie
        .ChartArea.Border.Width = 4
        .Header.Text = "变异退出原因分布图"
        .Header.Interior.ForegroundColor = vbBlack
        '.ChartArea.Pie.StartAngle = 90
        strSql = "Select g.编码, g.名称, Count(1) As 变异退出数量, 100 * Round(Count(1) / Sum(Count(1)) Over(), 4) ""比例""" & vbNewLine & _
                "From 病人门诊路径 A, 病人门诊路径评估 B,病人门诊路径变异 C ," & IIf(strDateTmp = "A.导入时间", "", "病人挂号记录 D,") & " 门诊变异常见原因 G" & vbNewLine & _
                "Where " & IIf(strDateTmp = "A.导入时间", "", "a.病人id = d.病人id And a.挂号ID = d.挂号ID And ") & " b.路径记录id = a.Id And b.天数 = a.当前天数 And  " & vbNewLine & _
                " b.路径记录Id=C.路径记录ID(+) And b.阶段ID=C.阶段ID(+) and b.日期=c.日期(+) " & vbNewLine & _
                " And g.编码 = NVl(c.变异原因,b.变异原因) And a.状态 = 3 And G.性质=2"
                '表“病人门诊路径评估”与表 “病人路径变异”采用外连接是为了兼容查询以前数据（病人路径变异为 10.34.0新增）
        '是否当前路径统计
        strSql = strSql & IIf(optAllPath.Value, "", " And a.路径id=[1]")
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By g.编码, g.名称"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.NumSeries = rsTmp.RecordCount
        If .ChartGroups.Item(1).Data.NumSeries <> 0 Then .ChartGroups.Item(1).Data.NumPoints(1) = 1
        i = 1
        Do Until rsTmp.EOF
            .ChartGroups.Item(1).SeriesLabels.Add (rsTmp!名称 & Space(2) & "共" & rsTmp!变异退出数量 & "人(" & Val(rsTmp!比例 & "") & "%)")
            .ChartGroups.Item(1).Data.Y(i, 1) = Val(rsTmp!比例 & "")
            rsTmp.MoveNext
            i = i + 1
        Loop
        '注意信息
        lblZY.Caption = "注：该图的计数规则是一个病人的一次就诊（每次就诊发生变异则为一次）。" & vbCrLf & _
                        "其中：没有使用过的变异退出原因不显示出来。"
        If rsTmp.RecordCount = 0 Then
            lblMsg.Caption = "您所指定的时间范围内未发现变异退出的病人。"
            lblMsg.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        '保存上次浏览的图
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub set时间变异分析(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        lblMsg.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '设置图形大小
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 40
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 80
        .ChartArea.Border.Width = 4
        .Header.Text = "时间变异情况分布图"
        .Header.Interior.ForegroundColor = vbBlack
        '柱的填充颜色，数量
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartGroups.Item(1).Data.NumPoints(1) = 5
        .ChartArea.Bar.ClusterWidth = 35
        '坐标阴影
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3D效果
        .ChartArea.View3D.depth = 10
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 25
        .ChartGroups.Item(1).SeriesLabels.Add ("例数")
        '坐标属性
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 45
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        '坐标内容
        .ChartGroups.Item(1).PointLabels.Add ("正常")
        .ChartGroups.Item(1).PointLabels.Add ("阶段提前")
        .ChartGroups.Item(1).PointLabels.Add ("阶段延后")
        .ChartGroups.Item(1).PointLabels.Add ("低于标准治疗时间")
        .ChartGroups.Item(1).PointLabels.Add ("超过标准治疗时间")
        strSql = "Select 变异,例数, 100 * Round(例数 / Decode(Sum(例数) Over(), 0, 1,Sum(例数) Over()), 4) 比例 From (With Test As" & vbNewLine & _
                " (Select Distinct b.路径记录id, Decode(b.时间进度, 0, '正常', 1, '阶段提前',2,'阶段提前', -1, '阶段延后') As 变异" & vbNewLine & _
                "  From 病人门诊路径 A, 病人门诊路径评估 B, 病人挂号记录 D" & vbNewLine & _
                "  Where a.挂号ID = d.ID And b.时间进度 <> 0 And a.id=b.路径记录ID"
        '是否当前路径统计
        strSql = strSql & IIf(optAllPath.Value, "", " And a.路径id=[1]")
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & "Group By b.路径记录id, b.时间进度" & vbNewLine & _
                " Union All " & vbNewLine & _
                " Select 路径记录id, 变异" & vbNewLine & _
                " From (Select a.Id As 路径记录id," & vbNewLine & _
                "              Decode(Sign(a.当前天数 -" & vbNewLine & _
                "                           Nvl(Substr(c.标准治疗时间, 0, Instr(c.标准治疗时间, '-') - 1), Substr(c.标准治疗时间, Instr(c.标准治疗时间, '-') + 1))), 0," & vbNewLine & _
                "                      '正常', -1, '低于标准治疗时间', 1," & vbNewLine & _
                "                      Decode(Sign(a.当前天数 - Substr(c.标准治疗时间, Instr(c.标准治疗时间, '-') + 1)), 1, '超过标准治疗时间', '正常')) As 变异" & vbNewLine & _
                "       From 病人门诊路径 A, 门诊路径版本 C, 病人挂号记录 D" & vbNewLine & _
                "       Where a.路径id = c.路径id And a.版本号 = c.版本号 And a.挂号ID = d.ID And a.结束时间 Is Not Null And a.当前天数 is not null"
        '是否当前路径统计
        strSql = strSql & IIf(optAllPath.Value, "", " And a.路径id=[1]")
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & ") Where 变异 <> '正常')" & vbNewLine & _
                "Select '正常' As 变异, Count(1) As 例数" & vbNewLine & _
                "From 病人门诊路径 A, 病人挂号记录 D" & vbNewLine & _
                "Where a.挂号ID = d.ID And a.当前天数 is not null And Not Exists (Select 1 From Test Where a.Id = Test.路径记录id)"
        '是否当前路径统计
        strSql = strSql & IIf(optAllPath.Value, "", " And a.路径id=[1]")
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & "Union All" & vbNewLine & _
                "Select 变异, Count(1) As 例数 From Test Group By 变异) group by 变异,例数"
                
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.Y(1, 1) = 0
        .ChartGroups.Item(1).Data.Y(1, 2) = 0
        .ChartGroups.Item(1).Data.Y(1, 3) = 0
        .ChartGroups.Item(1).Data.Y(1, 4) = 0
        .ChartGroups.Item(1).Data.Y(1, 5) = 0
        If rsTmp.RecordCount = 1 And Val(rsTmp!例数 & "") = 0 Then
            lblMsg.Caption = "您所指定的时间范围内未发现时间变异的病人。"
            lblMsg.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        Do Until rsTmp.EOF
            Select Case rsTmp!变异 & ""
                Case "正常"
                    .ChartGroups.Item(1).Data.Y(1, 1) = Val(rsTmp!例数 & "")
                    i = 1
                Case "阶段提前"
                    .ChartGroups.Item(1).Data.Y(1, 2) = Val(rsTmp!例数 & "")
                    i = 2
                Case "阶段延后"
                    .ChartGroups.Item(1).Data.Y(1, 3) = Val(rsTmp!例数 & "")
                    i = 3
                Case "低于标准治疗时间"
                    .ChartGroups.Item(1).Data.Y(1, 4) = Val(rsTmp!例数 & "")
                    i = 4
                Case "超过标准治疗时间"
                    .ChartGroups.Item(1).Data.Y(1, 5) = Val(rsTmp!例数 & "")
                    i = 5
            End Select
            '设置每个分区的标签
            Set chtLabel = .ChartLabels.Add()
            chtLabel.Offset = 15
            chtLabel.Border.Type = oc2dBorderShadow
            chtLabel.Border.Width = 2
            chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
            chtLabel.AttachMethod = oc2dAttachDataIndex
            chtLabel.AttachDataIndex.Point = i
            chtLabel.IsConnected = True
            chtLabel.Anchor = oc2dAnchorNorthEast
            chtLabel.Name = rsTmp!变异 & ""
            chtLabel.Text = "共" & rsTmp!例数 & "人(" & Val(rsTmp!比例 & "") & "%)"
            chtLabel.Font.Size = 8
            rsTmp.MoveNext
        Loop
        For i = 1 To 5
            If .ChartGroups.Item(1).Data.Y(1, i) = 0 Then
                 '没有例数的将标签=0补上
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
                chtLabel.Text = "共0人(0%)"
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
        '注意信息
        lblZY.Caption = "在一个病人的一次路径过程中：" & vbCrLf & _
                        "正常：指未发生其他4种变异的情况。" & vbCrLf & _
                        "阶段提前\阶段延后：一个病人在路径过程中只要发生了就算且仅算一次。(两个都发生了分别算一次)" & vbCrLf & _
                        "低于标准治疗时间\超过标准治疗时间:一个病人在路径过程中已经结束的路径的天数低于或高于标准治疗时间算一次。"
        '保存上次浏览的图
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Set未生成原因(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    With chtThis
        chtThis.Visible = False
        picTable.Visible = True
        lblZY.Visible = True
        vsgInfo(vsg_项目).Visible = True
        strHead = "分类,1500,1;原因,2000,1;例数,800,7"
        Call InitTable(vsgInfo(vsg_原因), strHead)
        
        strHead = "阶段名称,1500,1;项目名称,5000,1;例数,800,7"
        Call InitTable(vsgInfo(vsg_项目), strHead)
        '相同合并单元格
        vsgInfo(vsg_项目).MergeCells = flexMergeRestrictColumns
        vsgInfo(vsg_项目).MergeCol(VCol_阶段) = True
        vsgInfo(vsg_原因).Rows = 1
        vsgInfo(vsg_项目).Rows = 1
        fraGroupLR.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        vsgInfo(vsg_项目).TextMatrix(0, VCol_名称) = "项目名称"
        lblInfo(0).Caption = "未生成原因汇总表"
        lblInfo(1).Caption = "未生成项目汇总表(双击查看对应医嘱)"
        txtFindNum.Visible = False
        '原因表
        strSql = "Select b.编码,e.名称 as 上级名称, b.名称, Count(1) As 例数" & vbNewLine & _
                " From 病人门诊路径 A, 门诊变异常见原因 B, 病人门诊路径执行 C, 门诊变异常见原因 E" & vbNewLine & _
                " Where c.变异原因 = b.编码 And c.路径记录id = a.Id and e.编码=b.上级 And b.性质 = 1 And c.项目id Is Not Null"
        strSql = strSql & " And a.路径id=[1]"
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.编码, b.名称,e.名称 order by 例数 desc"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        
        With vsgInfo(vsg_原因)
        For i = 1 To rsTmp.RecordCount
                .AddItem ""
                .RowData(i) = rsTmp!编码 & ""
                .TextMatrix(i, VCol_分类) = rsTmp!上级名称
                .TextMatrix(i, VCol_原因) = rsTmp!名称 & ""
                .TextMatrix(i, VCol_原因例数) = rsTmp!例数 & ""
                
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(vsg_原因).Rows = 1 Then vsgInfo(vsg_原因).Rows = 2
        '未生成路径项目表
        strSql = "Select b.Id, b.项目内容, b.阶段id, e.名称 As 阶段名称, Count(1) As 例数, Nvl(f.序号, e.序号) 序号" & vbNewLine & _
                " From 病人门诊路径执行 C, 门诊路径项目 B, 病人门诊路径 A, 门诊路径阶段 E,门诊路径阶段 F" & vbNewLine & _
                " Where c.项目id = b.Id And c.路径记录id = a.Id And e.Id = b.阶段id And e.父id = f.Id(+) And" & vbNewLine & _
                "      c.项目id Is Not Null And c.变异原因 Is Not Null "
        strSql = strSql & " And a.路径id=[1]"
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.Id, b.项目内容, b.阶段id, e.名称,Nvl(f.序号, e.序号) Order By 序号,例数 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        
        With vsgInfo(vsg_项目)
        For i = 1 To rsTmp.RecordCount
                .AddItem ""
                .RowData(i) = rsTmp!ID & ""
                .TextMatrix(i, VCol_阶段) = rsTmp!阶段名称 & ""
                .Cell(flexcpData, i, VCol_阶段) = Val(rsTmp!阶段ID & "")
                .TextMatrix(i, VCol_名称) = rsTmp!项目内容 & ""
                .TextMatrix(i, VCol_项目例数) = rsTmp!例数 & ""
                
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(vsg_项目).Rows = 1 Then vsgInfo(vsg_项目).Rows = 2
        '注意信息
        lblZY.Caption = "注：本页面是为了统计单个病种中，必须生成但是又没有生成路径项目的变异信息。"
        '保存上次浏览的图
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Set路径外项目(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    With chtThis
        chtThis.Visible = False
        picTable.Visible = True
        lblZY.Visible = True
        vsgInfo(vsg_项目).Visible = True
        strHead = "分类,1300,1;原因,1800,1;例数,800,7"
        Call InitTable(vsgInfo(vsg_原因), strHead)
        
        strHead = "阶段名称,1300,1;项目名称,3050,1;例数,800,7"
        Call InitTable(vsgInfo(vsg_项目), strHead)
        '相同合并单元格
        vsgInfo(vsg_项目).MergeCells = flexMergeRestrictColumns
        vsgInfo(vsg_项目).MergeCol(VCol_阶段) = True
        vsgInfo(VSG_明细).Visible = False
        fraGroupUD.Visible = False
        fraGroupLR.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        imgFrom.Visible = False
        txtFindNum.Visible = False
        vsgInfo(vsg_原因).Rows = 1
        vsgInfo(vsg_项目).TextMatrix(0, VCol_名称) = "医嘱内容"
        lblInfo(0).Caption = "路径外项目产生原因汇总表"
        lblInfo(1).Caption = "路径外项目对应医嘱汇总表   显示每个阶段前     种医嘱"
        txtFindNum.Visible = True
        txtFindNum.Tag = "OK"
        '原因表
        strSql = "Select b.编码, b.名称,e.名称 as 上级名称, Count(1) As 例数" & vbNewLine & _
                " From 病人门诊路径 A, 门诊变异常见原因 B, 病人门诊路径执行 C,门诊变异常见原因 E" & vbNewLine & _
                " Where c.变异原因 = b.编码 And c.路径记录id = a.Id And e.编码=b.上级 And b.性质 = 1 And c.项目id is Null"
        strSql = strSql & " And a.路径id=[1]"
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & " Group By b.编码, b.名称,e.名称 order by 例数 desc"
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        
        With vsgInfo(vsg_原因)
            
        For i = 1 To rsTmp.RecordCount
                .AddItem ""
                .RowData(i) = rsTmp!编码 & ""
                .TextMatrix(i, VCol_分类) = rsTmp!上级名称
                .TextMatrix(i, VCol_原因) = rsTmp!名称 & ""
                .TextMatrix(i, VCol_原因例数) = rsTmp!例数 & ""
                
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(vsg_原因).Rows = 1 Then vsgInfo(vsg_原因).Rows = 2
        '获得路径外项目对应的医嘱
        Call GetPathOutAdvice
        If vsgInfo(vsg_项目).Rows = 1 Then vsgInfo(vsg_项目).Rows = 2
        '注意信息
        lblZY.Caption = "注：本页面是为了统计单个病种中，各个阶段添加的路径外项目的变异信息。"
        '保存上次浏览的图
        mlngVariation = tbcVariation.Selected.Index
    End With
End Sub

Private Sub Set路径完成情况(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        '保存上次浏览的图
        mlngSurvey = tbcVariation.Selected.Index
        '路径完成情况
        lblMsg.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '设置图形大小
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 40
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 80
        .ChartArea.Border.Width = 4
        .Header.Text = "路径完成情况分布图"
        .Header.Interior.ForegroundColor = vbBlack
        '柱的填充颜色，数量
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartGroups.Item(1).Data.NumPoints(1) = 5
        .ChartArea.Bar.ClusterWidth = 30
        '坐标阴影
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3D效果
        .ChartArea.View3D.depth = 10
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 25
        .ChartGroups.Item(1).SeriesLabels.Add ("例数")
        '坐标属性
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 45
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        '坐标内容
        .ChartGroups.Item(1).PointLabels.Add ("未导入")
        .ChartGroups.Item(1).PointLabels.Add ("正在执行")
        .ChartGroups.Item(1).PointLabels.Add ("正常完成")
        .ChartGroups.Item(1).PointLabels.Add ("变异完成")
        .ChartGroups.Item(1).PointLabels.Add ("变异退出")

        strSql = "Select 病例数, 未导入例数, Round(未导入例数 / 病例数, 4) * 100 As 未导入比例, 正在执行例数, Round(正在执行例数 / 病例数, 4) * 100 As 正在执行比例, 正常完成例数," & vbNewLine & _
                "       Round(正常完成例数 / 病例数, 4) * 100 As 正常完成比例, 变异退出例数, Round(变异退出例数 / 病例数, 4) * 100 As 变异退出比例, 变异完成例数," & vbNewLine & _
                "       Round(变异完成例数 / 病例数, 4) * 100 As 变异完成比例" & vbNewLine & _
                "From (Select Count(1) As 病例数, Sum(Decode(a.状态, 0, 1, 0)) As 未导入例数, Sum(Decode(a.状态, 1, 1, 0)) As 正在执行例数," & vbNewLine & _
                "              Sum(Decode(a.状态, 2, 1, 0)) As 正常完成例数, Sum(Decode(a.状态, 3, 1, 0)) As 变异退出例数," & vbNewLine & _
                "              Sum(Decode(a.状态, 100, 1, 0)) As 变异完成例数" & vbNewLine & _
                "       From (Select a.Id, Decode(a.状态, 2, Decode(Sign(Sum(Decode(p.评估结果, -1, 1, 0))), 0, 2, 1, 100), a.状态) As 状态" & vbNewLine & _
                "              From 病人门诊路径 A, 病人门诊路径评估 P" & vbNewLine & _
                "              Where a.Id = p.路径记录id(+) "
        '是否当前路径统计
        strSql = strSql & IIf(optAllPath.Value, "", " And a.路径id=[1]")
        '时间范围
        strSql = strSql & " And " & strDateTmp & " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
        "              Group By a.Id, a.状态) A)"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        .ChartGroups.Item(1).Data.Y(1, 1) = 0
        .ChartGroups.Item(1).Data.Y(1, 2) = 0
        .ChartGroups.Item(1).Data.Y(1, 3) = 0
        .ChartGroups.Item(1).Data.Y(1, 4) = 0
        .ChartGroups.Item(1).Data.Y(1, 5) = 0
        If rsTmp.RecordCount = 1 And Val(rsTmp!病例数 & "") = 0 Then
            lblMsg.Caption = "您所指定的时间范围内未发现临床路径病人。"
            lblMsg.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        
        If Not rsTmp.EOF Then
            For i = 1 To 5
                '设置每个分区的标签
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
                    .ChartGroups.Item(1).Data.Y(1, 1) = Val(rsTmp!未导入例数 & "")
                    chtLabel.Name = "未导入例数"
                    chtLabel.Text = "共" & rsTmp!未导入例数 & "例(" & Val(rsTmp!未导入比例 & "") & "%)"
                ElseIf i = 2 Then
                    .ChartGroups.Item(1).Data.Y(1, 2) = Val(rsTmp!正在执行例数 & "")
                    chtLabel.Name = "正在执行例数"
                    chtLabel.Text = "共" & rsTmp!正在执行例数 & "例(" & Val(rsTmp!正在执行比例 & "") & "%)"
                ElseIf i = 3 Then
                    .ChartGroups.Item(1).Data.Y(1, 3) = Val(rsTmp!正常完成例数 & "")
                    chtLabel.Name = "正常完成例数"
                    chtLabel.Text = "共" & rsTmp!正常完成例数 & "例(" & Val(rsTmp!正常完成比例 & "") & "%)"
                ElseIf i = 4 Then
                    .ChartGroups.Item(1).Data.Y(1, 4) = Val(rsTmp!变异完成例数 & "")
                    chtLabel.Name = "变异完成例数"
                    chtLabel.Text = "共" & rsTmp!变异完成例数 & "例(" & Val(rsTmp!变异完成比例 & "") & "%)"
                Else
                    .ChartGroups.Item(1).Data.Y(1, 5) = Val(rsTmp!变异退出例数 & "")
                    chtLabel.Name = "变异退出例数"
                    chtLabel.Text = "共" & rsTmp!变异退出例数 & "例(" & Val(rsTmp!变异退出比例 & "") & "%)"
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
        '注意信息
        lblZY.Caption = "注：该图的计数规则是一个病人的一次路径（每次就诊导入路径则为一次）。" & vbCrLf & _
                        "其中：未导入--导入时不符合导入条件的病人       正在执行--正在路径中的病人      正常完成--正常走完路径的病人。" & vbCrLf & _
                        "      变异完成--变异后继续走完路径的病人         变异退出--发生变异而没有走完路径的病人。"
    End With
End Sub

Private Sub Set阶段平均费用(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim lngEdition As Long
    
    With chtThis
        '保存上次浏览的图
        mlngSurvey = tbcVariation.Selected.Index
        '路径完成情况
        lblMsg.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        lblPathEdition.Visible = True
        cboPathEdition.Visible = True
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '设置图形大小
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 60
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 100
        .ChartArea.Border.Width = 4
        .Header.Text = "阶段平均费用分布图"
        .Header.Interior.ForegroundColor = vbBlack
        '柱的填充颜色，数量
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartArea.Bar.ClusterWidth = 15
        '坐标阴影
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3D效果
        .ChartArea.View3D.depth = 5
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 15
        
        '坐标属性
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 45
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        If Not mblnIsEdition And (mlngOldPathID <> lngPathID Or mdateOldStart <> CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00")) Or _
                                    mdateOldEnd <> CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59")) Or mstrDateType <> cboTimeType.Text) Then
            strSql = "Select Distinct 版本号" & vbNewLine & _
                    " From 病人门诊路径 A " & vbNewLine & _
                    " Where a.路径id=[1] "
                    
            '时间范围
            strSql = strSql & " And " & strDateTmp & _
                    " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
            strSql = strSql & " Order By 版本号 Desc"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                        Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
            cboPathEdition.Clear
            Do Until rsTmp.EOF
            
                cboPathEdition.AddItem "第 " & rsTmp!版本号 & " 版"
                cboPathEdition.ItemData(cboPathEdition.NewIndex) = Val(rsTmp!版本号 & "")
                rsTmp.MoveNext
            Loop
            If cboPathEdition.ListCount > 0 Then Call Cbo.SetIndex(cboPathEdition.Hwnd, 0)
            
        End If
        mblnIsEdition = False

        strSql = "Select h.名称 阶段名称, a.版本号, Avg(a.费用) As 平均费用 ,Nvl(g.序号, h.序号) 序号" & vbNewLine & _
                "From (Select f.病人id, b.阶段id, a.版本号, Sum(f.实收金额) As 费用" & vbNewLine & _
                "       From 病人门诊路径执行 B, 病人门诊路径 A, 病人门诊路径医嘱 C, 门诊费用记录 F " & vbNewLine & _
                "       Where b.路径记录id = a.Id And c.路径执行id = b.Id And c.病人医嘱id = f.医嘱序号 And" & vbNewLine & _
                "             f.记录状态 <> 0 And a.状态=2 "
        '当前路径统计
        strSql = strSql & " And a.路径id=[1] And a.版本号=[4]"

        
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
                
        strSql = strSql & "Group By f.病人id, a.版本号, b.阶段id" & vbNewLine & _
                "       Having Sum(f.实收金额) <> 0) A, 门诊路径阶段 H ,门诊路径阶段 G" & vbNewLine & _
                "Where h.Id = a.阶段id and h.父id = g.Id(+)" & vbNewLine & _
                "Group By nvl(g.id,h.Id), h.名称, a.版本号,Nvl(g.序号, h.序号) Order By 序号"

        If cboPathEdition.ListIndex = -1 Or cboPathEdition.ListCount = 0 Then
            lngEdition = 0
        Else
            lngEdition = Val(cboPathEdition.ItemData(cboPathEdition.ListIndex))
        End If
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"), lngEdition)

        If rsTmp.RecordCount = 0 Then
            lblMsg.Caption = "您所指定的时间范围内未发现产生费用的路径病人。"
            lblMsg.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        Else
            .ChartGroups.Item(1).Data.NumPoints(1) = rsTmp.RecordCount
        End If
        i = 1
        Do While Not rsTmp.EOF
            '坐标内容
            .ChartGroups.Item(1).PointLabels.Add (Mid(rsTmp!阶段名称 & "", 1, 10) & IIf(Len(rsTmp!阶段名称 & "") > 10, "...", ""))
            .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!平均费用 & "")
                
            '设置每个分区的标签
            Set chtLabel = .ChartLabels.Add()
            chtLabel.Offset = 15
            chtLabel.Border.Type = oc2dBorderShadow
            chtLabel.Border.Width = 2
            chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
            chtLabel.AttachMethod = oc2dAttachDataIndex
            chtLabel.AttachDataIndex.Point = i
            chtLabel.IsConnected = True
            chtLabel.Anchor = oc2dAnchorNorthEast
            chtLabel.Name = rsTmp!阶段名称 & ""
            chtLabel.Text = Format(rsTmp!平均费用, "##.00") & "元"
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
        '注意信息
        lblZY.Caption = "注：该图是当前路径下所选择的路径版本对应的阶段人均费用图。" & vbCrLf & _
                        "其中：1、该图统计的对象是已经正常走完当前路径的病人。" & vbCrLf & _
                        "       2、默认显示最新版本的阶段人均费用，可选择查看更早版本费用信息。" & vbCrLf & _
                        "       3、可选择的版本为当前选择的时间区域所用过的路径版本。"
        mlngOldPathID = lngPathID
        mdateOldStart = CDate(Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"))
        mdateOldEnd = CDate(Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        mstrDateType = cboTimeType.Text
    End With
End Sub

Private Sub Set就诊日分布图(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    
    With chtThis
        '保存上次浏览的图
        mlngSurvey = tbcVariation.Selected.Index
        '路径完成情况
        lblMsg.Visible = False
        chtThis.Visible = True
        lblZY.Visible = True
        optThisPath.Enabled = False
        optAllPath.Enabled = False
        .ChartGroups.Item(1).ChartType = oc2dTypeBar
        '设置图形大小
        .ChartArea.PlotArea.Top = 20
        .ChartArea.PlotArea.Left = 60
        .ChartArea.PlotArea.Right = 20
        .ChartArea.PlotArea.Bottom = 50
        .ChartArea.Border.Width = 4
        .Header.Interior.ForegroundColor = vbBlack
        '柱的填充颜色，数量
        .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = &H8000000D
        .ChartGroups.Item(1).Data.NumSeries = 1
        .ChartArea.Bar.ClusterWidth = 15
        '坐标阴影
        .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
        '3D效果
        .ChartArea.View3D.depth = 5
        .ChartArea.View3D.Elevation = 10
        .ChartArea.View3D.Rotation = 15
        
        '坐标属性
        .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
        .ChartArea.Axes.Item(1).Font.Size = 10
        .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
        .ChartGroups.Item(1).SeriesLabels.Add ("人数")
        

        strSql = "Select 当前天数, 人数, 标准治疗时间, Round(人数 / Sum(人数) Over(), 4) * 100 As 比例" & vbNewLine & _
                "From (Select a.当前天数, c.标准治疗时间, Count(1) As 人数" & vbNewLine & _
                "       From 病人门诊路径 A, 门诊路径目录 B, 门诊路径版本 C" & vbNewLine & _
                "       Where b.Id = a.路径id And b.Id = c.路径id And b.最新版本 = c.版本号 And a.状态 = 2 And a.当前天数 Is Not Null"
   
        '当前路径统计
        strSql = strSql & " And a.路径id=[1]"
        '时间范围
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
                
        strSql = strSql & " Group By a.当前天数, c.标准治疗时间" & vbNewLine & _
                        " Order By a.当前天数)"

        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                    Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))

        If rsTmp.RecordCount = 0 Then
            lblMsg.Caption = "您所指定的时间范围内未发现完成路径的病人。"
            lblMsg.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        Else
            .ChartGroups.Item(1).Data.NumPoints(1) = rsTmp.RecordCount
            .Header.Text = "就诊日分布图 " & vbCrLf & "(标准治疗时间：" & IIf(InStr(rsTmp!标准治疗时间 & "", "-") > 0, "", "≤") & rsTmp!标准治疗时间 & "天)"
        End If
        i = 1
        Do While Not rsTmp.EOF
            '坐标内容
            .ChartGroups.Item(1).PointLabels.Add (rsTmp!当前天数 & "天")
            .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!人数 & "")
                
            '设置每个分区的标签
            Set chtLabel = .ChartLabels.Add()
            chtLabel.Offset = 5
            chtLabel.Border.Type = oc2dBorderShadow
            chtLabel.Border.Width = 2
            chtLabel.Interior.BackgroundColor = RGB(255, 235, 205)
            chtLabel.AttachMethod = oc2dAttachDataIndex
            chtLabel.AttachDataIndex.Point = i
            chtLabel.IsConnected = True
            chtLabel.Anchor = oc2dAnchorAuto
            chtLabel.Name = rsTmp!当前天数 & ""
            chtLabel.Text = "共" & rsTmp!人数 & "人(" & rsTmp!比例 & "%)"
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
        '注意信息
        lblZY.Caption = "注：该图是当前路径下对应的时间范围内所有病人的就诊天数分布图。" & vbCrLf & _
                        "其中：1、该图统计的对象是已经正常走完当前路径的病人。" & vbCrLf & _
                        "       2、统计的就诊日表示病人在路径中的就诊天数。"
    End With
End Sub

Private Sub Set按医生统计(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    chtThis.Visible = False
    lblZY.Visible = True
    picTable.Visible = True
    strHead = "科室,2500,1;医生,1400,1;病人数,1100,7;入径人数,1100,7;入径率,1100,7;变异退出数,1100,7;变异退出率,1100,7;变异完成数,1100,7;变异完成率,1100,7;医嘱符合度,1100,7"
    Call InitTable(vsgInfo(vsg_原因), strHead)
    vsgInfo(vsg_原因).Width = picTable.Width
    vsgInfo(vsg_项目).Visible = False
    vsgInfo(VSG_明细).Visible = False
    fraGroupLR.Visible = False
    fraGroupUD.Visible = False
    imgFrom.Visible = False
    txtFindNum.Visible = False
    lblInfo(1).Caption = ""
    lblInfo(0).Caption = "按医生统计路径基本信息"
    vsgInfo(vsg_原因).MergeCells = flexMergeRestrictColumns
    vsgInfo(vsg_原因).MergeCol(VCOL_科室) = True
    
    strSql = "Select a.科室id, b.名称 As 科室, a.导入人, Sum(a.病人数) As 病人数, Sum(a.入径人数) As 入径人数, Sum(a.入径率) || '%' As 入径率," & vbNewLine & _
        "       Sum(a.变异退出数) As 变异退出数, Sum(a.变异退出率) || '%' As 变异退出率, Sum(a.变异完成数) As 变异完成数, Sum(a.变异完成率) || '%' As 变异完成率," & vbNewLine & _
        "       Round(Decode(Nvl(Sum(a.医嘱数), 0), 0, '0', (Nvl(Sum(a.医嘱数), 0) - Nvl(Sum(a.路径外医嘱), 0)) / Nvl(Sum(a.医嘱数), 0)) * 100," & vbNewLine & _
        "              2) || '%' As 医嘱符合度" & vbNewLine & _
        "From (Select a.科室id, a.导入人, Count(1) As 病人数, Sum(Decode(a.状态, 0, 0, 1)) As 入径人数," & vbNewLine & _
        "              Round(Sum(Decode(a.状态, 0, 0, 1)) / Count(1) * 100, 2) As 入径率, Sum(Decode(a.状态, 3, 1, 0)) As 变异退出数," & vbNewLine & _
        "              Decode(Sum(Decode(a.状态, 0, 0, 1)), 0, '0'," & vbNewLine & _
        "                      Round(Sum(Decode(a.状态, 3, 1, 0)) / Sum(Decode(a.状态, 0, 0, 1)) * 100, 2)) As 变异退出率," & vbNewLine & _
        "              Sum(Decode(a.状态, 100, 1, 0)) As 变异完成数," & vbNewLine & _
        "              Decode(Sum(Decode(a.状态, 0, 0, 1)), 0, '0'," & vbNewLine & _
        "                      Round(Sum(Decode(a.状态, 100, 1, 0)) / Sum(Decode(a.状态, 0, 0, 1)) * 100, 2)) As 变异完成率, Null As 医嘱数," & vbNewLine & _
        "              Null As 路径外医嘱" & vbNewLine & _
        "       From (Select a.Id, a.科室id, a.导入人," & vbNewLine & _
        "                     Decode(a.状态, 2, Decode(Sign(Sum(Decode(p.评估结果, -1, 1, 0))), 0, 2, 1, 100), a.状态) As 状态" & vbNewLine & _
        "              From 病人门诊路径 A, 病人门诊路径评估 P" & vbNewLine & _
        "              Where a.Id = p.路径记录id(+) And a.状态 <> 1 " & IIf(optAllPath.Value, "", " And a.路径id=[1]") & vbNewLine & _
        " And " & strDateTmp & " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
        "              Group By a.Id, a.科室id, a.导入人, a.状态) A" & vbNewLine & _
        "       Group By a.科室id, a.导入人 "
    strSql = strSql & vbNewLine & _
        "   Union All   " & vbNewLine & _
        "       Select 科室id, 导入人, Null, Null, Null, Null, Null, Null, Null, Count(1) As 医嘱数, Sum(路径外医嘱) As 路径外医嘱" & vbNewLine & _
        "       From (Select Distinct a.科室id, a.导入人, c.Id, Decode(b.路径执行id, Null, 1, 0) As 路径外医嘱" & vbNewLine & _
        "              From 病人门诊路径 A,病人门诊路径医嘱 B, 病人医嘱记录 C,病人门诊路径记录 D, 病人挂号记录 E " & vbNewLine & _
        "              Where a.ID = d.路径记录ID And d.挂号ID = E.ID And E.NO=C.挂号单 And c.Id = b.病人医嘱id(+) And" & vbNewLine & _
        "                    c.相关id Is Null And c.前提id Is Null And c.开始执行时间 Between a.开始时间 And" & vbNewLine & _
        "                    Nvl(a.结束时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.状态 = 2 " & IIf(optAllPath.Value, "", " And a.路径id=[1]") & vbNewLine & _
                    " And " & strDateTmp & " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS'))" & vbNewLine & _
        "       Group By 科室id, 导入人) A, 部门表 B" & vbNewLine & _
        "Where a.科室id = b.Id" & vbNewLine & _
        "Group By a.科室id, a.导入人, b.名称" & vbNewLine & _
        "Order By b.名称, a.科室id, Sum(a.变异退出率) Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
                    
    With vsgInfo(vsg_原因)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, VCOL_科室) = rsTmp!科室 & ""
                .TextMatrix(.Rows - 1, VCOL_医生姓名) = rsTmp!导入人 & ""
                .TextMatrix(.Rows - 1, VCOL_病人数) = rsTmp!病人数 & ""
                .TextMatrix(.Rows - 1, vcol_入径人数) = rsTmp!入径人数 & ""
                .TextMatrix(.Rows - 1, vcol_入径率) = rsTmp!入径率 & ""
                .TextMatrix(.Rows - 1, vcol_变异退出数) = rsTmp!变异退出数 & ""
                .TextMatrix(.Rows - 1, vcol_变异退出率) = rsTmp!变异退出率 & ""
                .TextMatrix(.Rows - 1, vcol_变异完成数) = rsTmp!变异完成数 & ""
                .TextMatrix(.Rows - 1, vcol_变异完成率) = rsTmp!变异完成率 & ""
                .TextMatrix(.Rows - 1, VCOL_医嘱符合度) = rsTmp!医嘱符合度 & ""
                rsTmp.MoveNext
            Loop
        Else
            .Rows = 2
        End If
    End With
    '注意信息
    lblZY.Caption = _
                    "说明：1、医嘱符合度=由路径模版产生的医嘱数/医生所有完成路径的病人路径期间的医嘱数。" & vbCrLf & _
                    "       2、医嘱符合度中的医嘱不包括医技科室下达的医嘱和路径期间以外（导入前和完成后的）的医嘱。" & vbCrLf & _
                    "       3、医生是指路径的导入人。"
    '保存上次浏览的图
    mlngVariation = tbcVariation.Selected.Index
End Sub

Private Sub set科室变异率排名(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String
    
    chtThis.Visible = False
    lblZY.Visible = True
    picTable.Visible = True
    fraGroupUD.Visible = False
    fraGroupLR.Visible = True
    vsgInfo(vsg_项目).Visible = True
    imgFrom.Visible = False
    txtFindNum.Visible = False
    vsgInfo(VSG_明细).Visible = False
    lblInfo(1).Caption = "科室变异率最低十名"
    lblInfo(0).Caption = "科室变异率最高十名"
    
    strHead = "科室,3000,1;变异退出率,1500,7;变异完成率,1500,7"
    Call Grid.Init(vsgInfo(vsg_原因), strHead)
    
    strHead = "科室,3000,1;变异退出率,1500,7;变异完成率,1500,7"
    Call Grid.Init(vsgInfo(vsg_项目), strHead)
    
    vsgInfo(vsg_项目).MergeCells = flexMergeRestrictColumns
    vsgInfo(vsg_项目).MergeCol(vsgInfo(vsg_项目).ColIndex("变异退出率")) = False
    vsgInfo(vsg_项目).MergeCol(vsgInfo(vsg_项目).ColIndex("变异完成率")) = False
    vsgInfo(vsg_原因).MergeCol(vsgInfo(vsg_原因).ColIndex("变异退出率")) = False
    vsgInfo(vsg_原因).MergeCol(vsgInfo(vsg_原因).ColIndex("变异完成率")) = False
            
    strSql = "Select a.科室id, a.名称 As 科室, Count(1), Round(Sum(Decode(a.状态, 3, 1, 0)) / Count(1) * 100, 2) As 变异退出率," & vbNewLine & _
            "       Round(Sum(Decode(a.状态, 100, 1, 0)) / Count(1) * 100, 2) As 变异完成率" & vbNewLine & _
            "From (Select a.Id, a.科室id, b.名称, Decode(a.状态, 2, Decode(Sign(Sum(Decode(p.评估结果, -1, 1, 0))), 0, 2, 1, 100), a.状态) As 状态" & vbNewLine & _
            "       From 病人门诊路径 A, 病人挂号记录 D, 部门表 B, 病人门诊路径评估 P" & vbNewLine & _
            "       Where a.挂号ID = d.ID And a.科室id = b.Id And a.Id = p.路径记录id  And a.状态 In (2, 3) " & vbNewLine
    '当前路径统计
    strSql = strSql & IIf(optAllPath.Value, "", " And a.路径id=[1]")
    '时间范围
    strSql = strSql & " And " & strDateTmp & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
    strSql = strSql & _
            "Group By a.Id, a.科室id, b.名称, a.状态) A" & vbNewLine & _
            "Group By a.科室id, a.名称" & vbNewLine & _
            "Order By 变异退出率 Desc"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
        Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
    
    With vsgInfo(vsg_原因)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("科室")) = rsTmp!科室 & ""
                .TextMatrix(.Rows - 1, .ColIndex("变异退出率")) = rsTmp!变异退出率 & "%"
                .TextMatrix(.Rows - 1, .ColIndex("变异完成率")) = rsTmp!变异完成率 & "%"
                If .Rows = 11 Then Exit Do
                rsTmp.MoveNext
            Loop
        Else
            .Rows = 2
        End If
    End With
    
    With vsgInfo(vsg_项目)
        .Rows = 1
        If rsTmp.RecordCount > 0 Then
            rsTmp.Sort = "变异退出率"
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                .AddItem ""
                .TextMatrix(.Rows - 1, .ColIndex("科室")) = rsTmp!科室 & ""
                .TextMatrix(.Rows - 1, .ColIndex("变异退出率")) = rsTmp!变异退出率 & "%"
                 .TextMatrix(.Rows - 1, .ColIndex("变异完成率")) = rsTmp!变异完成率 & "%"
                If .Rows = 11 Then Exit Do
                rsTmp.MoveNext
            Loop
        Else
            .Rows = 2
        End If
            
    End With
    
    '注意信息
    lblZY.Caption = "说明：变异率仅包含变异退出的。"
    '保存上次浏览的图
    mlngVariation = tbcVariation.Selected.Index
End Sub

Private Function Get总体情况路径情况(ByVal varTime As Variant, ByVal lngPathID As Long) As Recordset
'功能：获得总体情况的记录，根据不同的时间
    Dim strSql As String
    Dim lngTmp As Long
      
    strSql = "Select Sum(a.病人数) As 病人数, Sum(a.入径人数) As 入径人数, Nvl(Round(Sum(a.入径人数) / Sum(a.病人数) * 100, 2), 0) || '%' As 入径率," & vbNewLine & _
        "       Sum(a.变异退出数) As 变异退出数," & vbNewLine & _
        "       Decode(Sum(a.入径人数), 0, '0', Nvl(Round(Sum(a.变异退出数) / Sum(a.入径人数) * 100, 2), 0)) || '%' As 变异退出率," & vbNewLine & _
        "       Sum(a.变异完成数) As 变异完成数," & vbNewLine & _
        "       Decode(Sum(a.入径人数), 0, '0', Nvl(Round(Sum(a.变异完成数) / Sum(a.入径人数) * 100, 2), 0)) || '%' As 变异完成率," & vbNewLine & _
        "       Nvl(Round(Decode(Nvl(Sum(a.医嘱数), 0), 0, '0', (Nvl(Sum(a.医嘱数), 0) - Nvl(Sum(a.路径外医嘱), 0)) / Nvl(Sum(a.医嘱数), 0)) * 100," & vbNewLine & _
        "                  2), 0) || '%' As 医嘱符合度, Round(Sum(就诊天数) / Sum(a.病人数), 2) As 平均就诊日," & vbNewLine & _
        "       Round(Sum(实收金额) / Sum(a.病人数), 2) As 平均就诊费用," & vbNewLine & _
        "       Nvl(Decode(Sum(a.入径人数), 0, '0', 100 - Round(Sum(a.变异退出数) / Sum(a.入径人数) * 100, 2)), 0) || '%' As 完成率"
    strSql = strSql & vbNewLine & _
        " From (Select a.科室id, a.路径id, Count(1) As 病人数, Sum(入径人数) As 入径人数, Sum(变异退出数) As 变异退出数, Sum(变异完成数) As 变异完成数, Sum(a.就诊天数) As 就诊天数," & vbNewLine & _
                "       Sum(a.实收金额) As 实收金额, Null As 医嘱数, Null As 路径外医嘱" & vbNewLine & _
                "From (Select a.科室id, a.路径id, a.病人id, Decode(a.状态, 0, 0, 1) As 入径人数, Decode(a.状态, 3, 1, 0) As 变异退出数," & vbNewLine & _
                "              Decode(a.状态, 2, Decode(Sign(Sum(Decode(p.评估结果, -1, 1, 0))), 0, 0, 1, 1), 0) As 变异完成数, a.就诊天数, a.实收金额" & vbNewLine & _
                "       From (Select a.Id, a.科室id, a.路径id, a.病人id, A.当前天数 as 就诊天数, a.状态, Sum(b.实收金额) As 实收金额" & vbNewLine & _
                "              From 病人门诊路径 A, 门诊费用记录 B, 病人门诊路径执行 C, 病人门诊路径医嘱 D" & vbNewLine & _
                "              Where A.ID =C.路径记录ID And C.ID = d.路径执行ID And d.病人医嘱ID = b.医嘱序号(+) And a.状态 <> 1 And" & vbNewLine & _
                "                    b.记录状态(+) <> 0 " & vbNewLine & _
                IIf(optAllPath.Value, "", " And a.路径id=[1]") & _
                " And A.导入时间 Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
                "              Group By a.Id, a.科室id, a.路径id, a.病人id, A.当前天数, a.状态) A, 病人门诊路径评估 P" & vbNewLine & _
                "       Where a.Id = p.路径记录id(+)" & vbNewLine & _
                "       Group By a.科室id, a.路径id, a.病人id, a.就诊天数, a.实收金额, a.状态) A" & vbNewLine & _
                "Group By a.科室id, a.路径id"
    strSql = strSql & vbNewLine & _
            "   Union All   " & vbNewLine & _
            "Select 科室id, 路径id, Null, Null, Null, Null, Null, Null, Count(1) As 医嘱数, Sum(路径外医嘱) As 路径外医嘱" & vbNewLine & _
            "From (Select Distinct a.科室id, a.路径id, c.Id, Decode(b.路径执行id, Null, 1, 0) As 路径外医嘱" & vbNewLine & _
            "       From 病人门诊路径 A,病人门诊路径记录 E, 病人医嘱记录 C, 病人门诊路径医嘱 B, 病人挂号记录 D" & vbNewLine & _
            "       Where a.ID = E.路径记录ID And E.挂号ID =D.ID And d.No = c.挂号单 And c.Id = b.病人医嘱id(+) And" & vbNewLine & _
            "             c.相关id Is Null And c.前提id Is Null And c.开始执行时间 Between a.开始时间 And" & vbNewLine & _
            "             Nvl(a.结束时间, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.状态 = 2 " & IIf(optAllPath.Value, "", " And a.路径id=[1]") & vbNewLine & _
             " And A.导入时间 Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS'))" & vbNewLine & _
            "Group By 科室id, 路径id) A"

    lngTmp = cboYorM.ListIndex
        
    Set Get总体情况路径情况 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
        IIf(lngTmp = 0 Or lngTmp = 1, Format(varTime, "yyyy-MM-01 00:00:00"), Format(varTime, "yyyy-01-01 00:00:00")), _
        IIf(lngTmp = 0, Format(DateAdd("M", 1, CDate(varTime)), "yyyy-MM-01 00:00:00"), IIf(lngTmp = 1, Format(DateAdd("M", 3, CDate(varTime)), "yyyy-MM-01 00:00:00"), Format(Format(varTime, "yyyy") + 1 & "-01-01", "yyyy-MM-dd 00:00:00"))))
End Function

Private Function Get总体情况科室病种数(ByVal varTime As Variant, ByVal lngPathID As Long) As Recordset
'功能：获得总体情况的科室病种数，根据不同的时间
    Dim strSql As String
    
    strSql = "Select Sum(科室数) As 科室数, Sum(病种数) As 病种数" & vbNewLine & _
        "From (Select 1 科室数, Null As 病种数" & vbNewLine & _
        "       From 病人门诊路径 A" & vbNewLine & _
        "       Where a.状态 <> 1" & vbNewLine & _
        IIf(optAllPath.Value, "", " And a.路径id=[1]") & _
        " And A.导入时间" & _
        " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
        "       Group By a.科室id" & vbNewLine & _
        "       Union All" & vbNewLine & _
        "       Select Null, 1" & vbNewLine & _
        "       From 病人门诊路径 A " & vbNewLine & _
        "       Where a.状态 <> 1" & vbNewLine & _
        IIf(optAllPath.Value, "", " And a.路径id=[1]") & _
        " And A.导入时间" & _
        " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
        "       Group By a.路径id)"

    Set Get总体情况科室病种数 = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
        IIf(cboYorM.ListIndex = 0, Format(varTime, "yyyy-MM-01 00:00:00"), Format(varTime, "yyyy-01-01 00:00:00")), IIf(cboYorM.ListIndex = 0, Format(CDate(Format(varTime, "yyyy-mm")) + 32, "yyyy-MM-01 00:00:00"), Format(Format(varTime, "yyyy") + 1 & "-01-01", "yyyy-MM-dd 00:00:00")))
End Function

Private Sub set总体情况比对(ByVal lngPathID As Long)
    Dim rsTmp As Recordset
    
    Set rsTmp = Get总体情况路径情况(dtpTwo.Value, lngPathID)
    
    With vsgInfo(vsg_原因)
        .TextMatrix(3, VCOL_同期二) = Val(rsTmp!病人数 & "")
        .TextMatrix(4, VCOL_同期二) = Val(rsTmp!入径人数 & "")
        .TextMatrix(5, VCOL_同期二) = Val(rsTmp!入径人数 & "") - Val(rsTmp!变异退出数 & "")
        .TextMatrix(6, VCOL_同期二) = rsTmp!入径率 & ""
        .TextMatrix(7, VCOL_同期二) = rsTmp!完成率 & ""
        .TextMatrix(8, VCOL_同期二) = rsTmp!变异完成率 & ""
        .TextMatrix(9, VCOL_同期二) = rsTmp!变异退出率 & ""
        .TextMatrix(10, VCOL_同期二) = rsTmp!医嘱符合度 & ""
        .TextMatrix(11, VCOL_同期二) = Val(rsTmp!平均住院日 & "")
        .TextMatrix(12, VCOL_同期二) = Val(rsTmp!平均住院费用 & "")
        
        .TextMatrix(1, VCOL_差值) = Val(.TextMatrix(1, VCOL_同期一)) - Val(.TextMatrix(1, VCOL_同期二))
        .TextMatrix(2, VCOL_差值) = Val(.TextMatrix(2, VCOL_同期一)) - Val(.TextMatrix(2, VCOL_同期二))
        .TextMatrix(3, VCOL_差值) = Val(.TextMatrix(3, VCOL_同期一)) - Val(.TextMatrix(3, VCOL_同期二))
        .TextMatrix(4, VCOL_差值) = Val(.TextMatrix(4, VCOL_同期一)) - Val(.TextMatrix(4, VCOL_同期二))
        .TextMatrix(5, VCOL_差值) = Val(.TextMatrix(5, VCOL_同期一)) - Val(.TextMatrix(5, VCOL_同期二))
        
        'val(89.3%) -此类小数点和百分号同时出现在val函数中报实时错误 需特殊处理
        '------------------------------------------
        .TextMatrix(6, VCOL_差值) = Val(Replace(.TextMatrix(6, VCOL_同期一), "%", "")) - Val(Replace(.TextMatrix(6, VCOL_同期二), "%", "")) & "%"
        .TextMatrix(7, VCOL_差值) = Val(Replace(.TextMatrix(7, VCOL_同期一), "%", "")) - Val(Replace(.TextMatrix(7, VCOL_同期二), "%", "")) & "%"
        .TextMatrix(8, VCOL_差值) = Val(Replace(.TextMatrix(8, VCOL_同期一), "%", "")) - Val(Replace(.TextMatrix(8, VCOL_同期二), "%", "")) & "%"
        .TextMatrix(9, VCOL_差值) = Val(Replace(.TextMatrix(9, VCOL_同期一), "%", "")) - Val(Replace(.TextMatrix(9, VCOL_同期二), "%", "")) & "%"
        .TextMatrix(10, VCOL_差值) = Val(Replace(.TextMatrix(10, VCOL_同期一), "%", "")) - Val(Replace(.TextMatrix(10, VCOL_同期二), "%", "")) & "%"
        '------------------------------------------
        .TextMatrix(11, VCOL_差值) = Val(.TextMatrix(11, VCOL_同期一)) - Val(.TextMatrix(11, VCOL_同期二))
        .TextMatrix(12, VCOL_差值) = Val(.TextMatrix(12, VCOL_同期一)) - Val(.TextMatrix(12, VCOL_同期二))
        
        If optAllPath.Value Then
            Set rsTmp = Get总体情况科室病种数(dtpTwo.Value, lngPathID)
            .RowHidden(1) = False
            .RowHidden(2) = False
            .TextMatrix(1, VCOL_同期二) = Val(rsTmp!科室数 & "")
            .TextMatrix(2, VCOL_同期二) = Val(rsTmp!病种数 & "")
        Else
            .RowHidden(1) = True
            .RowHidden(2) = True
        End If
        If cboYorM.ListIndex = 1 Then
            .TextMatrix(0, VCOL_同期二) = Format(dtpTwo.Value, dtpTwo.CustomFormat) & "-" & Format(dtpFour.Value, dtpFour.CustomFormat)
        Else
            .TextMatrix(0, VCOL_同期二) = Format(dtpTwo.Value, dtpTwo.CustomFormat)
        End If
    End With
End Sub

Private Sub set总体情况(ByVal strDateTmp As String, ByVal lngPathID As Long)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strHead As String

    chtThis.Visible = False
    lblZY.Visible = True
    
    picContrast.Visible = True
    Call SetPicContrastFace  '界面调整
    
    picFind.Visible = False
    picTable.Visible = True
    strHead = "指标,3000,1;2012年10月,1500,7;2012年11月,1500,7;差值,1500,7"
    Call InitTable(vsgInfo(vsg_原因), strHead)
    vsgInfo(vsg_原因).Width = picTable.Width
    vsgInfo(vsg_项目).Visible = False
    vsgInfo(VSG_明细).Visible = False
    fraGroupLR.Visible = False
    fraGroupUD.Visible = False
    imgFrom.Visible = False
    txtFindNum.Visible = False
    lblInfo(1).Caption = ""
    lblInfo(0).Caption = "统计医院临床路径总体情况"
    
    vsgInfo(vsg_原因).MergeCells = flexMergeRestrictColumns
    vsgInfo(vsg_原因).MergeCol(VCOL_科室) = False
    vsgInfo(vsg_原因).Rows = 11
    chkContrast_Click
     With vsgInfo(vsg_原因)
        .Rows = 13
        
        Set rsTmp = Get总体情况路径情况(dtpOne.Value, lngPathID)
        .TextMatrix(1, VCOL_指标) = "科室数"
        .TextMatrix(2, VCOL_指标) = "病种数"
        .TextMatrix(3, VCOL_指标) = "病例总人数"
        .TextMatrix(4, VCOL_指标) = "入径总人数"
        .TextMatrix(5, VCOL_指标) = "完成总人数"
        .TextMatrix(6, VCOL_指标) = "入径率"
        .TextMatrix(7, VCOL_指标) = "完成率"
        .TextMatrix(8, VCOL_指标) = "变异完成率"
        .TextMatrix(9, VCOL_指标) = "变异退出率"
        .TextMatrix(10, VCOL_指标) = "医嘱符合度"
        .TextMatrix(11, VCOL_指标) = "平均就诊天数"
        .TextMatrix(12, VCOL_指标) = "平均就诊总费用"
        
        .TextMatrix(3, VCOL_同期一) = Val(rsTmp!病人数 & "")
        .TextMatrix(4, VCOL_同期一) = Val(rsTmp!入径人数 & "")
        .TextMatrix(5, VCOL_同期一) = Val(rsTmp!入径人数 & "") - Val(rsTmp!变异退出数 & "")
        .TextMatrix(6, VCOL_同期一) = rsTmp!入径率 & ""
        .TextMatrix(7, VCOL_同期一) = rsTmp!完成率 & ""
        .TextMatrix(8, VCOL_同期一) = rsTmp!变异完成率 & ""
        .TextMatrix(9, VCOL_同期一) = rsTmp!变异退出率 & ""
        .TextMatrix(10, VCOL_同期一) = rsTmp!医嘱符合度 & ""
        .TextMatrix(11, VCOL_同期一) = Val(rsTmp!平均就诊日 & "")
        .TextMatrix(12, VCOL_同期一) = Val(rsTmp!平均就诊费用 & "")
        
        If optAllPath.Value Then
            Set rsTmp = Get总体情况科室病种数(dtpOne.Value, lngPathID)
            .RowHidden(1) = False
            .RowHidden(2) = False
            .TextMatrix(1, VCOL_同期一) = Val(rsTmp!科室数 & "")
            .TextMatrix(2, VCOL_同期一) = Val(rsTmp!病种数 & "")
        Else
            .RowHidden(1) = True
            .RowHidden(2) = True
        End If
        
        If cboYorM.ListIndex = 1 Then
            .TextMatrix(0, VCOL_同期一) = Format(dtpOne.Value, dtpOne.CustomFormat) & "-" & Format(dtpThree.Value, dtpThree.CustomFormat)
            .TextMatrix(0, VCOL_同期二) = Format(dtpTwo.Value, dtpTwo.CustomFormat) & "-" & Format(dtpFour.Value, dtpFour.CustomFormat)
            Call .AutoSize(VCOL_同期一, VCOL_同期二)
        Else
            .TextMatrix(0, VCOL_同期一) = Format(dtpOne.Value, dtpOne.CustomFormat)
            .TextMatrix(0, VCOL_同期二) = Format(dtpTwo.Value, dtpTwo.CustomFormat)
        End If
    End With
    
    '注意信息
    lblZY.Caption = _
    "说明：1、该表只统计就诊病人。" & vbCrLf & _
    "      2、医嘱符合度=由路径模版产生的医嘱数/医生所有完成路径的病人路径期间的医嘱数。" & vbCrLf & _
    "      3、按全院路径统计时，可统计使用门诊临床路径的科室数和病种数。"
    '保存上次浏览的图
    mlngSurvey = tbcVariation.Selected.Index
End Sub

Private Function GetXNum() As Long
'功能：获得趋势图X坐标的点数
    Dim lngXNum As Long
    
    If cboTrendTime.ListIndex = 0 Then
        '按天
        If cboInterval.List(cboInterval.ListIndex) = "一周" Then
            lngXNum = 7
        ElseIf cboInterval.List(cboInterval.ListIndex) = "一月" Then
            lngXNum = DateAdd("M", 1, Format(dtpTrendStart.Value, "yyyy-MM-dd")) - CDate(Format(dtpTrendStart.Value, "yyyy-MM-dd"))
        ElseIf cboInterval.List(cboInterval.ListIndex) = "两月" Then
            lngXNum = DateAdd("M", 2, Format(dtpTrendStart.Value, "yyyy-MM-dd")) - CDate(Format(dtpTrendStart.Value, "yyyy-MM-dd"))
        Else
            lngXNum = DateAdd("M", 3, Format(dtpTrendStart.Value, "yyyy-MM-dd")) - CDate(Format(dtpTrendStart.Value, "yyyy-MM-dd"))
        End If
    Else
        If cboInterval.List(cboInterval.ListIndex) = "半年" Then
            lngXNum = 6
        ElseIf cboInterval.List(cboInterval.ListIndex) = "一年" Then
            lngXNum = 12
        ElseIf cboInterval.List(cboInterval.ListIndex) = "两年" Then
            lngXNum = 24
        Else
            lngXNum = 36
        End If
    End If
    GetXNum = lngXNum
End Function

Private Sub set平均就诊费用(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '横向坐标数
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
     lblMsg.Visible = False
     chtThis.Visible = True
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = False
     optAllPath.Enabled = False
     optIn.Visible = True
     optOut.Visible = True
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '设置图形大小
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '柱的填充颜色，数量
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '坐标阴影
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '设置为3D效果
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '坐标属性
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("费用(元)")
         .ChartGroups.Item(1).SeriesLabels.Add ("标准值")
         '横向坐标标签
         
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
         If optIn.Value Then
            strSql = "Select 平均就诊费用, 结束时间, Sum(平均就诊费用) Over() As 总数 From " & _
            "(select sum(实收金额) as 实收金额,结束时间,round(sum(实收金额)/sum(人数),2) as 平均就诊费用 from " & _
            "(Select sum(d.实收金额) as 实收金额,1 as 人数,trunc(A.结束时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as 结束时间" & vbNewLine & _
            "       From 病人门诊路径 A ,病人门诊路径执行 B, 病人门诊路径医嘱 C, 门诊费用记录 D" & vbNewLine & _
            "       Where A.ID = B.路径记录ID And b.ID=C.路径执行ID And C.病人医嘱ID=D.医嘱序号 And a.状态 =2 And D.记录状态 <> 0 " & vbNewLine & _
            "        And a.路径ID=[1] And A.结束时间 Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(A.结束时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ,a.ID ) group by 结束时间)"
         Else
            strSql = "Select 平均就诊费用, 结束时间, Sum(平均就诊费用) Over() As 总数" & vbNewLine & _
                "From (Select round(sum(实收金额)/Count(1),2) As 平均就诊费用, Trunc(A.结束时间, " & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") As 结束时间" & vbNewLine & _
                "       From (Select A.ID, sum(c.实收金额) As 实收金额, Max(A.完成时间) As 结束时间" & vbNewLine & _
                "              From 病人挂号记录 A,病人医嘱记录 B,门诊费用记录 C " & vbNewLine & _
                "              Where  A.NO = B.挂号单 And B.ID = C.医嘱序号 And c.记录状态 <> 0 And" & vbNewLine & _
                "                    Not Exists (Select 1 From 病人门诊路径记录 C,病人门诊路径 D Where D.ID=C.路径记录ID And b.ID = c.挂号id And D.状态 =2 And D.路径id = [1]) And " & vbNewLine & _
                "                    A.完成时间 Between To_Date([2], 'YYYY-MM-DD HH24:MI:SS') And To_Date([3], 'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
                "              Group By A.ID) A" & vbNewLine & _
                "       Group By Trunc(A.结束时间, " & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & "))"

         End If
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!总数 & "")
         For i = 1 To lngXNum
             '最多显示19个标签
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM月"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "结束时间=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!平均就诊费用 & "")
                 If lngMax < Val(rsTmp!平均就诊费用 & "") Then lngMax = Val(rsTmp!平均就诊费用 & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
             
         Next
         .ChartArea.Axes(2).Max = lngMax + lngMax / 5
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "平均就诊费用趋势图"
         '注意信息
        lblZY.Caption = "说明：1、该图只统计门诊病人。" & vbCrLf & _
                        "       2、就诊费用只包括病人的医嘱收费。" & vbCrLf & _
                        "       3、标准值是指统计时间期间的平均值。"
         '保存上次浏览的图
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub set入径率(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '横向坐标数
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
     lblMsg.Visible = False
     chtThis.Visible = True
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = True
     optAllPath.Enabled = True
     optIn.Visible = False
     optOut.Visible = False
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '设置图形大小
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '柱的填充颜色，数量
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '坐标阴影
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '设置为3D效果
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '坐标属性
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("入径率(%)")
         .ChartGroups.Item(1).SeriesLabels.Add ("标准值")
         '横向坐标标签
         
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
        strSql = "Select 入径率, 出院日期, Sum(入径率) Over() As 总数 From " & _
            "(Select round(sum(decode(a.状态,0,0,1))/count(1) *100,2) as 入径率,trunc(A.导入时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as 出院日期" & vbNewLine & _
            "       From 病人门诊路径 A Where  " & vbNewLine & _
            IIf(optAllPath.Value, "", "a.路径id=[1] And ") & _
            " A.导入时间" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(A.导入时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ) "
         
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!总数 & "")
         For i = 1 To lngXNum
             '最多显示19个标签
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM月"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "出院日期=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!入径率 & "")
                 If lngMax < Val(rsTmp!入径率 & "") Then lngMax = Val(rsTmp!入径率 & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             '标准线
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
         Next
         .ChartArea.Axes(2).Max = IIf(lngMax + lngMax / 5 > 100, 100, lngMax + lngMax / 5)
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "路径入径率"
         '注意信息
        lblZY.Caption = "说明：1、该图只统计完成就诊的病人。" & vbCrLf & _
                        "      2、标准值是指统计时间期间的平均值。"
         '保存上次浏览的图
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub set完成率(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '横向坐标数
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long

     lblMsg.Visible = False
     chtThis.Visible = True
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = True
     optAllPath.Enabled = True
     optIn.Visible = False
     optOut.Visible = False
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '设置图形大小
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '柱的填充颜色，数量
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '坐标阴影
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '设置为3D效果
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '坐标属性
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("完成率(%)")
         .ChartGroups.Item(1).SeriesLabels.Add ("标准值")
         '横向坐标标签
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
        strSql = "Select 完成率, 出院日期, Sum(完成率) Over() As 总数 From " & _
            "(Select round(sum(decode(a.状态,2,1,0))/count(1) *100,2) as 完成率,trunc(A.导入时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as 出院日期" & vbNewLine & _
            "       From 病人门诊路径 A " & vbNewLine & _
            "       Where a.状态 in(2,3) " & vbNewLine & _
            IIf(optAllPath.Value, "", " And a.路径id=[1] ") & _
            " And A.导入时间" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(A.导入时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ) "
         
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!总数 & "")
         For i = 1 To lngXNum
             '最多显示19个标签
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM月"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "出院日期=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!完成率 & "")
                 If lngMax < Val(rsTmp!完成率 & "") Then lngMax = Val(rsTmp!完成率 & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             '标准线
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
             
         Next
         .ChartArea.Axes(2).Max = IIf(lngMax + lngMax / 5 > 100, 100, lngMax + lngMax / 5)
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "路径完成率"
         '注意信息
        lblZY.Caption = "说明：1、该图只统计完成路径的病人。" & vbCrLf & _
                        "      2、标准值是指统计时间期间的平均值。"
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub set变异率(ByVal lngPathID As Long)
     Dim strSql As String, rsTmp As Recordset
     Dim i As Long
     Dim chtLabel As ChartLabel
     Dim lngXNum As Long '横向坐标数
     Dim lngMax As Long, lngMin As Long
     Dim lngavg As Long
    
     lblMsg.Visible = False
     chtThis.Visible = True
     picTrend.Visible = True
     picFind.Visible = False
     lblZY.Visible = True
     optThisPath.Enabled = True
     optAllPath.Enabled = True
     optIn.Visible = False
     optOut.Visible = False
     With chtThis
         .ChartGroups.Item(1).ChartType = oc2dTypePlot
         '设置图形大小
         .ChartArea.PlotArea.Top = 20
         .ChartArea.PlotArea.Left = 60
         .ChartArea.PlotArea.Right = 20
         .ChartArea.PlotArea.Bottom = 50
         .ChartArea.Border.Width = 4
         .Header.Interior.ForegroundColor = vbBlack
         '柱的填充颜色，数量
         .ChartGroups.Item(1).Data.NumSeries = 2
         .ChartGroups.Item(1).Styles.Item(1).Fill.Interior.ForegroundColor = RGB(255, 128, 0)
         .ChartGroups.Item(1).Styles.Item(2).Fill.Interior.ForegroundColor = RGB(151, 64, 38)
         '坐标阴影
         .ChartArea.PlotArea.Interior.BackgroundColor = &HF0F8FF
         .ChartArea.Axes(2).MajorGrid.Spacing.IsDefault = True
        
         '设置为3D效果
         .ChartArea.View3D.depth = 0
         .ChartArea.View3D.Elevation = 0
         .ChartGroups.Item(1).Styles.Item(1).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Symbol.Shape = oc2dShapeNone
         .ChartGroups.Item(1).Styles.Item(2).Line.Width = 3
         .ChartGroups.Item(1).Styles.Item(1).Line.Width = 2
         '坐标属性
         .ChartArea.Axes.Item(1).AnnotationRotationAngle = 0
         .ChartArea.Axes.Item(1).Font.Size = 10
         .ChartArea.Axes.Item(1).AnnotationMethod = oc2dAnnotatePointLabels
         .ChartGroups.Item(1).SeriesLabels.Add ("变异率(%)")
         .ChartGroups.Item(1).SeriesLabels.Add ("标准值")
         '横向坐标标签
         lngXNum = GetXNum
         .ChartGroups.Item(1).Data.NumPoints(1) = lngXNum
         
        strSql = "Select 变异率, 完成时间, Sum(变异率) Over() As 总数 From " & _
            "(Select round(sum(decode(a.状态,3,1,0))/count(1) *100,2) as 变异率,trunc(A.导入时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") as 完成时间" & vbNewLine & _
            "       From 病人门诊路径 A " & vbNewLine & _
            "       Where a.状态 in(2,3) " & vbNewLine & _
            IIf(optAllPath.Value, "", " And a.路径id=[1] ") & _
            " And A.导入时间" & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')" & _
            "       group by trunc(A.导入时间," & IIf(cboTrendTime.ListIndex = 0, "'dd'", "'MM'") & ") ) "
         
         Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
             IIf(cboTrendTime.ListIndex = 0, Format(dtpTrendStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpTrendStart.Value, "yyyy-MM-01 00:00:00")), IIf(cboTrendTime.ListIndex = 0, Format(DateAdd("D", lngXNum, dtpTrendStart.Value), "yyyy-MM-dd 00:00:00"), Format(DateAdd("M", lngXNum, dtpTrendStart.Value), "yyyy-MM-01 00:00:00")))
        
         If rsTmp.RecordCount > 0 Then lngavg = Val(rsTmp!总数 & "")
         For i = 1 To lngXNum
             '最多显示19个标签
             If i Mod IIf(lngXNum < 10, 1, lngXNum \ 10) = 0 Then
                 .ChartGroups.Item(1).PointLabels.Add Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "MM.dd", "MM月"))
             Else
                  .ChartGroups.Item(1).PointLabels.Add ""
             End If
             rsTmp.Filter = "完成时间=" & Format(DateAdd(IIf(cboTrendTime.ListIndex = 0, "D", "M"), i - 1, dtpTrendStart.Value), IIf(cboTrendTime.ListIndex = 0, "yyyy-MM-dd", "yyyy-MM-01"))
             If rsTmp.RecordCount > 0 Then
                 .ChartGroups.Item(1).Data.Y(1, i) = Val(rsTmp!变异率 & "")
                 If lngMax < Val(rsTmp!变异率 & "") Then lngMax = Val(rsTmp!变异率 & "")
             Else
                 .ChartGroups.Item(1).Data.Y(1, i) = 0
                 lngMin = 0
             End If
             '标准线
             .ChartGroups.Item(1).Data.Y(2, i) = lngavg / lngXNum
         Next
         .ChartArea.Axes(2).Max = IIf(lngMax + lngMax / 5 > 100, 100, lngMax + lngMax / 5)
         .ChartArea.Axes(2).Min = lngMin - lngMin / 5
         .ChartArea.Axes(2).MajorGrid.Spacing.Value = .ChartArea.Axes(2).TickSpacing
         
         .Header.Text = "路径变异率"
         
         '注意信息
        lblZY.Caption = "说明：1、该图只统计完成路径的病人。" & vbCrLf & _
                        "       2、标准值是指统计时间期间的平均值。"
         '保存上次浏览的图
         mlngTrend = tbcVariation.Selected.Index
    End With
End Sub

Private Sub tbcVariation_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim strSql As String, rsTmp As Recordset
    Dim lngPathID As Long
    Dim i As Long
    Dim chtLabel As ChartLabel
    Dim strDateTmp As String
    Dim strHead As String
    Dim dblTmp As Double
    
    If mblnFirstLoad And Item.Tag <> "按医生统计" Then Exit Sub
    
    strDateTmp = cboTimeType.List(cboTimeType.ListIndex)
    If strDateTmp = "导入时间" Then strDateTmp = "A.导入时间"
    If strDateTmp = "结束时间" Then strDateTmp = "A.结束时间"
    On Error GoTo errH
    With chtThis
        '初始图形格式
        '屏幕禁止刷新，加载完成后再设为false
        .IsBatched = True
        picTable.Visible = False
        picTrend.Visible = False
        vsgInfo(VSG_明细).Visible = True
        fraGroupUD.Visible = True
        imgFrom.Visible = True
        lblInfo(2).Visible = True
        If InStr(mstrPrivs, "全院路径") <> 0 Then
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
        '右边的标签设置
        .Legend.Border = oc2dBorder3DOut
        .Legend.Border.Width = 4
        '图形表头
        .Header.IsShowing = True
        .Header.Font.Size = 18
        .Header.Font.Name = "楷体"
        .Header.Font.Bold = True
        '设置为3D效果
        .ChartArea.View3D.depth = 20
        .ChartArea.View3D.Elevation = 20
        '设置图形大小
        .ChartArea.PlotArea.Top = 60
        .ChartArea.PlotArea.Left = 55
        .ChartArea.PlotArea.Right = 60
        .ChartArea.PlotArea.Bottom = 35
        
        If rptPath.SelectedRows.count > 0 Or optAllPath.Value Then
            If Not rptPath.SelectedRows(0).GroupRow Or optAllPath.Value Then
                If rptPath.SelectedRows.count > 0 And Not rptPath.SelectedRows(0).GroupRow Then lngPathID = Val(rptPath.SelectedRows(0).Record(COL_ID).Value)
                Select Case Item.Tag
                    
                    Case "未导入原因"
                        Call Set未导入原因(strDateTmp, lngPathID)
                    Case "变异退出分析"
                        Call Set变异退出分析(strDateTmp, lngPathID)
                    Case "时间变异分析"
                        Call set时间变异分析(strDateTmp, lngPathID)
                    Case "未生成原因"
                        Call Set未生成原因(strDateTmp, lngPathID)
                    Case "路径外项目"
                        Call Set路径外项目(strDateTmp, lngPathID)
                    Case "路径完成情况"
                        Call Set路径完成情况(strDateTmp, lngPathID)
                    Case "阶段平均费用"
                        Call Set阶段平均费用(strDateTmp, lngPathID)
                    Case "就诊日分布图"
                        Call Set就诊日分布图(strDateTmp, lngPathID)
                    Case "按医生统计"
                        Call Set按医生统计(strDateTmp, lngPathID)
                    Case "科室变异率排名"
                        Call set科室变异率排名(strDateTmp, lngPathID)
                    Case "总体情况"
                        Call set总体情况(strDateTmp, lngPathID)
                    Case "平均就诊费用"
                        Call set平均就诊费用(lngPathID)
                    Case "入径率"
                        Call set入径率(lngPathID)
                    Case "完成率"
                        Call set完成率(lngPathID)
                    Case "变异率"
                        Call set变异率(lngPathID)
                End Select
            Else
                lblMsg.Caption = "按当前路径统计需要选中一个路径。"
                lblMsg.Visible = True
                .Visible = False
                lblZY.Visible = False
                .ChartArea.Border.Width = 0
            End If
        Else
            lblMsg.Caption = "按当前路径统计需要选中一个路径。"
            lblMsg.Visible = True
            .Visible = False
            lblZY.Visible = False
            .ChartArea.Border.Width = 0
        End If
        .IsBatched = False
        .Refresh
        Call picTable_Resize
        If Me.Visible And InStr(";按医生统计;科室变异率排名;未生成原因;路径外项目;总体情况;", ";" & Item.Tag & ";") > 0 Then
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
'参数：blnNext=是否查找下一个
    Static blnReStart As Boolean
    Dim blnHave As Boolean, i As Long, strFindTmp As String
    
    Call zlControl.TxtSelAll(txtFind)
            
    '开始查找行
    If rptPath.SelectedRows.count > 0 Then blnHave = True
    If Not blnNext Or blnReStart Or Not blnHave Then
        i = 0 'ReportControl的索引从是0开始
    Else
        i = rptPath.SelectedRows(0).Index + 1
    End If
    
    '查找路径
    strFindTmp = txtFind.Text
    For i = i To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If Not .GroupRow Then
                If cbsMain.FindControl(, 0).Caption = "路径查找" Then
                    If .Record(COL_名称).Value Like "*" & strFindTmp & "*" Then Exit For
                Else
                    If .Record(COL_诊断编码).Value = Trim(strFindTmp) Or _
                       .Record(COL_疾病编码).Value = Trim(strFindTmp) Or _
                       .Record(COL_诊断名称).Value Like "*" & strFindTmp & "*" Or _
                       .Record(COL_疾病名称).Value Like "*" & strFindTmp & "*" _
                       Then Exit For
                End If
            End If
        End With
    Next

    If i <= rptPath.Rows.count - 1 Then
        blnReStart = False
        '该行选中且显示在可见区域,并引发SelectionChanged事件
        Set rptPath.FocusedRow = rptPath.Rows(i)
        
        If rptPath.Visible Then rptPath.SetFocus
    Else
        blnReStart = True
        MsgBox IIf(blnNext, "后面已", "") & "找不到符合条件的临床路径。", vbInformation, gstrSysName
    End If
End Sub

Private Sub txtFind_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "请输入临床路径名称、诊断或者疾病" & vbCrLf & "查找(Ctrl+F)" & vbCrLf & "查找下一个(F3)"
    zlCommFun.ShowTipInfo txtFind.Hwnd, strTip, True
End Sub

Private Function LoadPatiList(ByVal lng路径ID As Long, Optional ByVal lngPersonID As Long) As Boolean
'功能：读取路径应用的病人清单
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intState As Integer
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strDateTmp As String
    Dim strIsVariation As String
    Dim strBranch As String
    Dim strBranchName As String
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    mlng病人ID = 0
    mlng挂号ID = 0
    mlng病人路径ID = 0
    
    If optState(0).Value Then intState = 0
    If optState(1).Value Then intState = 1
    If optState(2).Value Or optState(4).Value Then
        strIsVariation = " And " & IIf(optState(2).Value, "Not", "") & " Exists (Select 1 From 病人门诊路径评估 Where 路径记录id = a.Id And 评估结果 = -1) "
        intState = 2
    End If
    If optState(3).Value Then intState = 3

    strDateTmp = cboForDate.List(cboForDate.ListIndex)
    
    If strDateTmp = "导入时间" Then strDateTmp = "A.导入时间"
    If strDateTmp = "结束时间" Then strDateTmp = "A.结束时间"
    
    '不符合和变异退出显示原因
    rptPati.Columns(COL_不符合原因).Visible = optState(0)
    rptPati.Columns(COL_变异退出原因).Visible = optState(3)
    
    strSql = "Select Distinct A.ID,a.病人id, a.挂号ID,D.NO,D.执行状态 as 病人状态,a.科室ID, f.名称 As 科室,NVL(D.姓名, e.姓名) 姓名,NVL(D.性别, e.性别) 性别 ,NVL(D.年龄, e.年龄) 年龄 , d.门诊号, a.状态, a.当前天数, a.版本号," & vbNewLine & _
    "       b.最新版本, c.标准治疗时间, c.标准费用, a.导入人, a.导入时间, a.结束时间, E.就诊状态 As 病人状态,NULL As 数据转出,j.打印人,j.打印时间," & vbNewLine & _
    "       i.名称 As 不符合原因, " & IIf(intState = 2, "''", "decode(a.状态,3,g.名称,'')") & " As 变异退出原因,Decode(Q.Id,Null,0,1) as 患者版打印" & vbNewLine & _
    " From 病人门诊路径 A, 门诊路径目录 B, 门诊路径版本 C," & IIf(intState = 2, "", " 病人门诊路径评估 H, 门诊变异常见原因 G,") & _
    " 病人挂号记录 D, 病人信息 E, 部门表 F, 门诊变异常见原因 I,电子病历打印 J, 电子病历打印 Q " & vbNewLine & _
    " Where a.路径id = b.Id And a.路径id = c.路径id And a.版本号 = c.版本号 And a.挂号ID = d.ID And a.病人id = e.病人id And" & vbNewLine & _
    "      a.科室id = f.Id And j.文件id(+) = a.Id And j.种类(+) = 11 And (j.Id = (Select Max(ID) From 电子病历打印 Where 文件id(+) = a.Id And 种类 = 11) Or j.Id Is Null)" & _
    "And Q.文件id(+) = a.Id And Q.种类(+) = 12 And (Q.Id = (Select Max(ID) From 电子病历打印 Where 文件id(+) = a.Id And 种类 = 12) Or Q.Id Is Null )" & vbNewLine & _
    IIf(intState = 2, "", " And h.路径记录id(+) = a.Id And h.天数(+) = a.当前天数 And g.编码(+) = h.变异原因 ") & _
    " And i.编码(+) = a.未导入原因 And A.路径ID=[1] "
 
    If lngPersonID = 0 Then
        strSql = strSql & " And A.状态=[2]" & _
        " And " & strDateTmp & _
        " Between To_Date([3],'YYYY-MM-DD HH24:MI:SS') And To_Date([4],'YYYY-MM-DD HH24:MI:SS')"
        strSql = strSql & strIsVariation
        strSql = strSql & strBranch
        
        If intState = 3 Then
            strSql = strSql & " And g.性质=2"
        ElseIf intState = 0 Then
            strSql = strSql & " And i.性质=0"
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径ID, intState, _
        Format(dtpTime(0).Value, "yyyy-MM-dd 00:00:00"), Format(dtpTime(1).Value, "yyyy-MM-dd 23:59:59"), strBranchName)
    Else
        '查找病人，屏蔽时间等信息
        strSql = strSql & " And e.病人id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng路径ID, lngPersonID)
    End If
    
    '记录下刷新后的病人记录集，供打印使用
    '打印需要用到NO,病人状态，病人ID,挂号ID等字段
    Set mrsTmp = rsTmp
    
    rptPati.Records.DeleteAll
    Do While Not rsTmp.EOF
        Set objRecord = Me.rptPati.Records.Add()
        
        Set objItem = objRecord.AddItem("")
        objItem.HasCheckbox = True
            If rptPati.Columns(col_打印).Icon = img16.ListImages("UnCheck").Index - 1 Then
                objItem.Checked = True
            Else
                objItem.Checked = False
            End If
        Set objItem = objRecord.AddItem(Val(rsTmp!病人ID))
        Set objItem = objRecord.AddItem(Val(rsTmp!挂号ID))
        Set objItem = objRecord.AddItem(CStr(rsTmp!科室))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!姓名)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!性别)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!年龄)))
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!门诊号)))
        
        If NVL(rsTmp!状态, 0) = 1 And Not IsNull(rsTmp!当前天数) Then
            If InStr(rsTmp!标准治疗时间, "-") > 0 Then
                Set objItem = objRecord.AddItem(CInt(Val(rsTmp!当前天数) / Val(Split(rsTmp!标准治疗时间, "-")(1)) * 100) & "%")
            Else
                Set objItem = objRecord.AddItem(CInt(Val(rsTmp!当前天数) / Val(rsTmp!标准治疗时间) * 100) & "%")
            End If
        Else
            Set objItem = objRecord.AddItem("")
        End If
        
        Set objItem = objRecord.AddItem(NVL(rsTmp!标准治疗时间) & IIf(Not IsNull(rsTmp!标准治疗时间), "天", ""))
        Set objItem = objRecord.AddItem(NVL(rsTmp!标准费用) & IIf(Not IsNull(rsTmp!标准费用), "元", ""))
        Set objItem = objRecord.AddItem(rsTmp!版本号 & "/" & rsTmp!最新版本)
        Set objItem = objRecord.AddItem(CStr(NVL(rsTmp!导入人)))
        Set objItem = objRecord.AddItem(Format(rsTmp!导入时间, "yyyy-MM-dd HH:mm"))
        Set objItem = objRecord.AddItem(Format(rsTmp!结束时间, "yyyy-MM-dd HH:mm"))

        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!科室ID, 0)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!病人状态, 0)))
        Set objItem = objRecord.AddItem(Val(NVL(rsTmp!数据转出, 0)))
        Set objItem = objRecord.AddItem(NVL(rsTmp!不符合原因))
        Set objItem = objRecord.AddItem(NVL(rsTmp!变异退出原因))
        Set objItem = objRecord.AddItem(NVL(rsTmp!打印人))
        Set objItem = objRecord.AddItem(NVL(Format(rsTmp!打印时间, "yyyy-MM-dd HH:mm")))
        Set objItem = objRecord.AddItem(IIf(rsTmp!患者版打印 = 0, "", " √"))
        Set objItem = objRecord.AddItem(rsTmp!ID & "")
        rsTmp.MoveNext
    Loop
    rptPati.Populate
    
    If rptPati.Rows.count = 0 Then
        Me.stbThis.Panels(3).Text = ""
    Else
        Me.stbThis.Panels(3).Text = "当前路径共有 " & rptPati.Records.count & " 个应用病人"
    End If
    '设置窗体尺寸
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

Private Function LoadOperInfo(ByVal lng病人ID As Long, ByVal lng挂号ID As Long) As Boolean
    '功能：读取路径应用的病人清单
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim intSource As Integer
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    On Error GoTo errH
    Screen.MousePointer = 11
    intSource = -1
    strSql = "Select Id,记录来源,手术日期,已行手术 As 手术名称,主刀医师,麻醉医师 From 病人手麻记录 Where 病人ID=[1] And 主页ID=[2] Order By 记录来源"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病人ID, lng挂号ID)
    
    rptOper.Records.DeleteAll
    Do While Not rsTmp.EOF
        If intSource = -1 Then intSource = Val("" & rsTmp!记录来源)
        If intSource = Val("" & rsTmp!记录来源) Then
            Set objRecord = Me.rptOper.Records.Add()
            
            Set objItem = objRecord.AddItem("" & rsTmp!ID)
            Set objItem = objRecord.AddItem("" & rsTmp!手术名称)
            Set objItem = objRecord.AddItem("" & rsTmp!手术日期)
            Set objItem = objRecord.AddItem("" & rsTmp!主刀医师)
            Set objItem = objRecord.AddItem("" & rsTmp!麻醉医师)
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
'功能:记录表打印
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objReport As ReportControl
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    Dim strSubhead As String
    Dim objVSF As VSFlexGrid
    Dim blnIsRPT As Boolean   'True-是ReportControl对象需要转换成VSF对象
    Dim blnPath As Boolean    'True-输出 临床路径清单
    Dim strTmp As String
    Dim objTable As Object

    If rptPath.SelectedRows.count = 1 Then
        Select Case tbcSub.Selected.Caption
        Case "病人路径"
            If rptPati.Records.count > 0 And mstrFlag = "RPTPATI" Then
                Set objTable = rptPati
                strSubhead = rptPath.SelectedRows(0).Record(COL_名称).Value & "应用病人清单"
            Else
                blnPath = True  '临床路径清单
            End If
            blnIsRPT = True
        Case "变异分析", "概况分析"
             '“按医生统计”、"科室变异率排名"、“未生成原因”、“路径外项目”、“总体情况”
            If mstrFlag = "RPTPATH" Then
                blnPath = True: blnIsRPT = True
            Else
                If optAllPath.Value And optAllPath.Enabled Then
                    strTmp = "全院路径"
                Else
                    If Not rptPath.SelectedRows(0).GroupRow Then
                        strTmp = rptPath.SelectedRows(0).Record(COL_名称).Value
                    End If
                End If
                Select Case tbcVariation.Selected.Caption
                Case "按医生统计"
                   If mstrFlag <> "" And mstrFlag = "VSGINFO_0" Then
                       Set objTable = vsgInfo(vsg_原因)
                       strSubhead = strTmp & "_按医生统计路径基本信息"
                   End If
                Case "科室变异率排名"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_原因)
                        strSubhead = strTmp & "_科室变异率最高十名"
                    ElseIf mstrFlag = "VSGINFO_1" Then
                        Set objTable = vsgInfo(vsg_项目)
                        strSubhead = strTmp & "_科室变异率最低十名"
                    End If
                Case "未生成原因"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_原因)
                        strSubhead = strTmp & "_未生成原因汇总表"
                    ElseIf mstrFlag = "VSGINFO_1" Then
                        Set objTable = vsgInfo(vsg_项目)
                        strSubhead = strTmp & "_未生成项目汇总表"
                    ElseIf mstrFlag = "VSGINFO_2" Then
                        Set objTable = vsgInfo(VSG_明细)
                        strSubhead = strTmp & "_未生成项目明细表"
                    End If
                Case "路径外项目"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_原因)
                        strSubhead = strTmp & "_路径外项目产生原因汇总表"
                    ElseIf mstrFlag = "VSGINFO_1" Then
                        Set objTable = vsgInfo(vsg_项目)
                        strSubhead = strTmp & "_路径外项目对应医嘱汇总表"
                    End If
                   
                Case "总体情况"
                    If mstrFlag = "VSGINFO_0" Then
                        Set objTable = vsgInfo(vsg_原因)
                        strSubhead = "医院临床路径总体情况"
                    End If
                Case Else
                    blnPath = True: blnIsRPT = True '临床路径清单
                End Select
            End If
        Case Else
            blnPath = True: blnIsRPT = True '临床路径清单
        End Select
    End If
    
    If blnPath Then
        Set objTable = rptPath
        strSubhead = "门诊临床路径清单"  '输出 临床路径清单
    End If
    '-------------------------------------------------
    '复制数据表格
    If blnIsRPT Then
        Set objReport = objTable
        If objReport.Records.count = 0 Then Exit Sub
        If zlControl.RPTCopyToVSF(objReport, vsTemp) Is Nothing Then Exit Sub
    Else
        Set objVSF = objTable
        If Grid.CopyTo(objVSF, vsTemp) Is Nothing Then Exit Sub
    End If

    '调用打印部件处理
    Set objPrint.Body = Me.vsTemp
    objPrint.Title.Text = strSubhead
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("打印人:" & UserInfo.姓名)
    Call objAppRow.Add("打印时间:" & Format(Now, "yyyy-MM-dd HH:mm"))
    Call objPrint.BelowAppRows.Add(objAppRow)

    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub zlRptBatPrint()
'功能：批量打印路径表
    Dim i As Long
    
    With rptPati
        If rptPati.Rows.count < 1 Then MsgBox "当前病人列表没有可打印的路径表。", vbInformation, Me.Caption: Exit Sub
        If optState(0).Value Then MsgBox "当前选择的病人为[不符合]的路径病人，没有可用的路径表。", vbInformation, Me.Caption: Exit Sub
        If tbcSub.Selected.Tag <> "病人路径" Then MsgBox "请选择[病人路径]卡片，再进行打印操作。", vbInformation, Me.Caption: Exit Sub
        mrsTmp.Filter = 0
        For i = 1 To .Rows.count - 1
            If Not .Rows(i).GroupRow Then
                If .Rows(i).Record(col_打印).Checked Then
                    '过滤需要打印的病人
                    mrsTmp.Filter = IIf(mrsTmp.Filter = 0, "", mrsTmp.Filter) & IIf(mrsTmp.Filter = 0, "", " or ") & " (病人ID =" & .Rows(i).Record(COL_病人ID).Value & " And 挂号ID=" & .Rows(i).Record(COL_挂号ID).Value & ") "
                End If
            End If
        Next
        frmBatPrint.ShowMe Me, mrsTmp, 1
    End With
End Sub

Private Sub FuncShowPath()
    Dim vPati As TYPE_Pati
    
    With rptPati.SelectedRows(0)
        vPati.病人ID = .Record(COL_病人ID).Value
        vPati.挂号ID = .Record(COL_挂号ID).Value
        vPati.科室ID = .Record(COL_科室ID).Value
        vPati.病人状态 = .Record(COL_病人状态).Value
        
        frmPathTrackView.ShowMe Me, vPati, .Record(COL_数据转出).Value = 1, 1
    End With
End Sub

Private Sub FuncShowReport()
    Dim lng路径ID As Long, str名称 As String
    Dim rsTmp As ADODB.Recordset, strSql As String
 
    If rptPath.SelectedRows.count <= 0 Then Exit Sub
    
    lng路径ID = rptPath.SelectedRows(0).Record(COL_ID).Value
    If lng路径ID <> 0 Then
        str名称 = rptPath.SelectedRows(0).Record(COL_名称).Value
        Call frmReport1.ShowMe(gfrmMain, lng路径ID, str名称)
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
    If KeyCode = vbKeyReturn Then Call GetPathOutAdvice: vsgInfo(vsg_项目).SetFocus
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
        txtPerson.Tag = "不清空"
        Call FindPerson
        txtPerson.Tag = ""
    End If
End Sub

Private Sub txtPerson_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    strTip = "请输入病人姓名或门诊号。"
    zlCommFun.ShowTipInfo txtPerson.Hwnd, strTip, True
End Sub

Private Sub FindPerson()
    Dim strSql As String, vRect As RECT, rsTmp As Recordset, strTmp As String, varPara As Variant, i As Long
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    
    '如果是数字，则查门诊号，否则查姓名
    On Error GoTo errH
    varPara = txtPerson.Text
    If IsNumeric(varPara) And InStr(varPara, ".") = 0 And InStr(varPara, "-") = 0 And InStr(varPara, "+") = 0 Then
        strTmp = " And E.门诊号=[1]"
        varPara = CLng(txtPerson.Text)
    Else
        strTmp = " And E.姓名 like [1]"
        varPara = gstrLike & txtPerson.Text & "%"
    End If
    strSql = "Select a.路径id, a.病人id As ID, e.姓名,e.性别,e.年龄,e.门诊号,a.导入时间" & vbNewLine & _
            "From 病人门诊路径 A, 病人信息 E" & vbNewLine & _
            "Where a.病人id = e.病人id "
    strSql = strSql & strTmp
    vRect = zlControl.GetControlRect(txtPerson.Hwnd)
    
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, Me.Caption, _
            False, "", "", False, True, True, vRect.Left, vRect.Top, _
            txtPerson.Height, False, False, False, varPara)
            
    If rsTmp Is Nothing Then
        MsgBox "找不到符合条件的病人。", vbInformation, gstrSysName
        Call txtPerson.SetFocus
        txtPerson.SelStart = 0
        txtPerson.SelLength = Len(txtPerson.Text)
        Exit Sub
    End If
    
    For i = 0 To rptPath.Rows.count - 1
        With rptPath.Rows(i)
            If Not .GroupRow Then
                If .Record(COL_ID).Value = Val("" & rsTmp!路径ID) Then Exit For
            End If
        End With
    Next

    If i <= rptPath.Rows.count - 1 Then
        rptPath.Tag = "1"
        Set rptPath.FocusedRow = rptPath.Rows(i)
        rptPath.Tag = ""
        If rptPath.Visible Then rptPath.SetFocus
    Else
        MsgBox "找不到符合条件的病人。", vbInformation, gstrSysName
    End If
    Call LoadPatiList(Val("" & rsTmp!路径ID), Val("" & rsTmp!ID))
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
    
    If Index = vsg_项目 Then
        If Not vsgInfo(VSG_明细).Visible Then Exit Sub
        If vsgInfo(vsg_项目).Rows = vsgInfo(vsg_项目).FixedRows And NewRow <> vsgInfo(vsg_项目).FixedRows - 1 Then Exit Sub
        vsgInfo(VSG_明细).Rows = 1
        strDateTmp = cboTimeType.List(cboTimeType.ListIndex)
    
        If strDateTmp = "导入时间" Then strDateTmp = "A.导入时间"
        strSql = "Select A.病人id, F.姓名, F.门诊号, c.登记人, e.名称 As 原因, c.登记时间" & vbNewLine & _
                " From 病人门诊路径 A, 病人门诊路径执行 C,病人挂号记录 F, 门诊变异常见原因 E" & vbNewLine & _
                " Where c.路径记录id = a.Id And F.ID =a.挂号ID And e.编码 = c.变异原因 And e.性质 = 1 And" & vbNewLine & _
                "      c.项目id Is Not Null And c.变异原因 Is Not Null"
        strSql = strSql & " And c.项目id=[1]"
        strSql = strSql & " And " & strDateTmp & _
                " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS') Order by c.登记时间 desc"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(vsgInfo(vsg_项目).RowData(NewRow) & ""), _
                Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"))
        With vsgInfo(VSG_明细)
        For i = 1 To rsTmp.RecordCount
                .AddItem ""
                .RowData(i) = rsTmp!病人ID & ""
                .TextMatrix(i, VCol_姓名) = rsTmp!姓名 & ""
                .TextMatrix(i, VCOL_门诊号) = rsTmp!门诊号 & ""
                .TextMatrix(i, VCOL_医生) = rsTmp!登记人 & ""
                .TextMatrix(i, VCol_未使用原因) = rsTmp!原因 & ""
                .TextMatrix(i, VCol_生成时间) = rsTmp!登记时间 & ""
            rsTmp.MoveNext
        Next
        End With
        If vsgInfo(VSG_明细).Rows = 1 Then vsgInfo(VSG_明细).Rows = 2
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetArea(ByVal lngRow As Long, ByVal lngCol As Long) As CONST_AREA
'功能：获取指定行列在哪一块区域
    With vsgInfo(vsg_项目)
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
    
    '双击项目，查看医嘱
    If Index = vsg_项目 Then
        If Not vsgInfo(VSG_明细).Visible Then Exit Sub
        With vsgInfo(vsg_项目)
            lngRow = .MouseRow
            lngCol = .MouseCol
            
            vArea = GetArea(lngRow, lngCol)
            If vArea <> Area_Cross And vArea <> -1 Then
                If Val(.RowData(lngRow)) <> 0 Then
                    Call frmPathItemEditOut.ShowView(Me, Val(.RowData(lngRow)))
                End If
            End If
        End With
    End If
End Sub

Private Sub GetPathOutAdvice()
'功能：获得路径外项目所对应的医嘱信息
    Dim strSql As String, rsTmp As Recordset
    Dim lngPathID As Long
    Dim i As Long
    Dim strDateTmp As String

    If txtFindNum.Tag = "" Then Exit Sub
    
    strDateTmp = cboTimeType.List(cboTimeType.ListIndex)
    
    If strDateTmp = "导入时间" Then strDateTmp = "A.导入时间"
    vsgInfo(vsg_项目).Rows = 1
    
    If rptPath.SelectedRows.count > 0 And Not rptPath.SelectedRows(0).GroupRow Then lngPathID = Val(rptPath.SelectedRows(0).Record(COL_ID).Value)
    '医嘱内容汇总表
    strSql = "select * from(Select c.诊疗项目id, c.医嘱名称, c.例数, e.名称 As 阶段名称, Nvl(f.序号, e.序号) 阶段序号,ROW_NUMBER() over(PARTITION BY e.名称 order by Nvl(f.序号, e.序号),c.例数 desc) as Top" & vbNewLine & _
            " From (With Test As (Select g.Id, g.相关id, h.类别, h.名称 As 诊疗项目名称, h.操作类型, h.Id As 诊疗项目id, g.医嘱内容, c.阶段id" & vbNewLine & _
            "                    From 病人门诊路径 A, 病人门诊路径医嘱 B, 病人门诊路径执行 C, 病人医嘱记录 G, 病人挂号记录 D, 诊疗项目目录 H" & vbNewLine & _
            "                    Where c.路径记录id = a.Id And b.路径执行id = c.Id And g.Id = b.病人医嘱id And a.挂号ID = d.ID And" & vbNewLine & _
            "                          c.项目id Is Null And h.Id = g.诊疗项目id"
    strSql = strSql & " And a.路径id=[1]"
    '时间范围
    strSql = strSql & " And " & strDateTmp & _
            " Between To_Date([2],'YYYY-MM-DD HH24:MI:SS') And To_Date([3],'YYYY-MM-DD HH24:MI:SS')"
    strSql = strSql & ")" & vbNewLine & _
            "     --1、输血类和检验类" & vbNewLine & _
            "       Select Test.阶段id, Test.诊疗项目id, Test.医嘱内容 As 医嘱名称, Count(1) As 例数" & vbNewLine & _
            "       From Test" & vbNewLine & _
            "       Where Test.相关id Is Null And (Test.类别 = 'K' Or (Test.类别 = 'E' And Test.操作类型 = '6'))" & vbNewLine & _
            "       Group By Test.诊疗项目id, Test.阶段id, Test.医嘱内容" & vbNewLine & _
            "       --2、一并给药，除给药途径外每种药分开显示" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select Test.阶段id, Test.诊疗项目id, Test.诊疗项目名称 As 医嘱名称, Count(1) As 例数" & vbNewLine & _
            "       From Test" & vbNewLine & _
            "       Where Test.类别 In ('4', '5', '6')" & vbNewLine & _
            "       Group By Test.诊疗项目id, Test.阶段id, Test.诊疗项目名称"
    strSql = strSql & "--3、中药，取聚合后的诊疗项目名称" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select 阶段id, Max(诊疗项目id) As 诊疗项目id, f_List2str(Cast(Collect(医嘱名称) As t_Strlist)) 医嘱名称, Count(1) 例数" & vbNewLine & _
            "       From (Select Test.阶段id, Test.诊疗项目id, Test.诊疗项目名称 As 医嘱名称, Test.相关id" & vbNewLine & _
            "              From Test" & vbNewLine & _
            "              Where Test.类别 = '7'" & vbNewLine & _
            "              Order By 医嘱名称)" & vbNewLine & _
            "       Group By 相关id, 阶段id" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       --4、其他" & vbNewLine & _
            "       Select Test.阶段id, Test.诊疗项目id, Test.诊疗项目名称 As 医嘱名称, Count(1) As 例数" & vbNewLine & _
            "       From Test" & vbNewLine & _
            "       Where Test.相关id Is Null And (Test.类别 <> 'E' Or (Test.类别 = 'E' And Test.操作类型 Not In ('2', '4', '6'))) And Test.类别 <> 'K'" & vbNewLine & _
            "       Group By Test.诊疗项目id, Test.阶段id, Test.诊疗项目名称) C, 门诊路径阶段 E, 门诊路径阶段 F" & vbNewLine & _
            "Where e.Id = c.阶段id And e.父id = f.Id(+)" & vbNewLine & _
            "  Order By 阶段序号, 例数 Desc) where top<=[4]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPathID, _
                Format(dtpStart.Value, "yyyy-MM-dd 00:00:00"), Format(dtpEnd.Value, "yyyy-MM-dd 23:59:59"), IIf(Val(txtFindNum.Text) = 0, 5, Val(txtFindNum.Text)))
    
    With vsgInfo(vsg_项目)
    For i = 1 To rsTmp.RecordCount
            .AddItem ""
            .RowData(i) = rsTmp!诊疗项目ID & ""
            .TextMatrix(i, VCol_阶段) = rsTmp!阶段名称 & ""
            .TextMatrix(i, VCol_名称) = rsTmp!医嘱名称 & ""
            .TextMatrix(i, VCol_项目例数) = rsTmp!例数 & ""
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
'功能：打开公共对话框,保存图片
    With dlgPublic
        .DialogTitle = "保存图片文件"
        .Filter = "Jpeg|*.jpg"
        .Flags = &H200000 + &H2000 + &H2 + &H800
        .InitDir = App.Path
        .FileName = Format(Now, "yyyymmddhhmmss")
        .ShowSave
        Call chtThis.SaveImageAsJpeg(.FileName, 100, False, False, False) '长文件名：图片质量（0-100）:是否显示为灰度图像：是否压缩：是否增强显示
    End With
End Sub

Private Sub SetPicContrastFace()
'功能:设置PicContrast界面显示效果
    dtpThree.Visible = cboYorM.ListIndex = 1
    dtpFour.Visible = cboYorM.ListIndex = 1
    lblFromToOne.Visible = cboYorM.ListIndex = 1
    lblFromToTwo.Visible = cboYorM.ListIndex = 1
    
    '调整位置
    cboYorM.Left = 120
    If cboYorM.ListIndex = 1 Then '按季度
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
'功能：选中表格时,表格上方显示用于标识选中的图标，便于用户察觉
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
