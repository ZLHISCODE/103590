VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmPatientReprotFind 
   AutoRedraw      =   -1  'True
   Caption         =   "检验报告查询"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15435
   Icon            =   "frmPatientReprotFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   105
      ScaleHeight     =   930
      ScaleWidth      =   15045
      TabIndex        =   0
      Top             =   810
      Width           =   15045
      Begin VB.Frame fraTop 
         Height          =   945
         Left            =   45
         TabIndex        =   6
         Top             =   -60
         Width           =   14955
         Begin VB.PictureBox picDiseases 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   2940
            ScaleHeight     =   405
            ScaleWidth      =   2250
            TabIndex        =   83
            Top             =   480
            Visible         =   0   'False
            Width           =   2250
            Begin VB.ComboBox cboDiseases 
               Height          =   300
               ItemData        =   "frmPatientReprotFind.frx":6852
               Left            =   870
               List            =   "frmPatientReprotFind.frx":685F
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   90
               Width           =   1245
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "传 染 病"
               Height          =   180
               Left            =   45
               TabIndex        =   85
               Top             =   135
               Width           =   720
            End
         End
         Begin VB.TextBox txtDoctor 
            Height          =   300
            Left            =   1005
            TabIndex        =   57
            Top             =   570
            Width           =   1800
         End
         Begin VB.ComboBox cboPrint 
            Height          =   300
            ItemData        =   "frmPatientReprotFind.frx":687B
            Left            =   3810
            List            =   "frmPatientReprotFind.frx":687D
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   210
            Width           =   1245
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   1005
            TabIndex        =   46
            Top             =   210
            Width           =   1800
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Left            =   6270
            TabIndex        =   45
            Top             =   210
            Width           =   1530
         End
         Begin VB.TextBox txtPatiNo 
            Height          =   300
            Left            =   8730
            TabIndex        =   7
            Top             =   210
            Width           =   2190
         End
         Begin MSComCtl2.DTPicker dtpS 
            Height          =   300
            Left            =   6270
            TabIndex        =   8
            Top             =   570
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   202047491
            CurrentDate     =   40954
         End
         Begin MSComCtl2.DTPicker dtpE 
            Height          =   300
            Left            =   7980
            TabIndex        =   9
            Top             =   570
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   202047491
            CurrentDate     =   40954
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "开 单 人"
            Height          =   180
            Left            =   180
            TabIndex        =   58
            Top             =   615
            Width           =   720
         End
         Begin VB.Label Label5 
            Caption         =   "打印状态"
            Height          =   195
            Left            =   3030
            TabIndex        =   49
            Top             =   255
            Width           =   735
         End
         Begin VB.Label lblDept 
            Caption         =   "申请科室"
            Height          =   225
            Left            =   180
            TabIndex        =   47
            Top             =   255
            Width           =   930
         End
         Begin VB.Label lblName 
            Caption         =   "姓    名"
            Height          =   225
            Left            =   5400
            TabIndex        =   44
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "报告时间"
            Height          =   225
            Left            =   5400
            TabIndex        =   11
            Top             =   615
            Width           =   720
         End
         Begin VB.Label lblNo 
            Caption         =   "条码号↓"
            Height          =   225
            Left            =   7980
            TabIndex        =   10
            Top             =   255
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox PicPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   12300
      ScaleHeight     =   1515
      ScaleWidth      =   1035
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1035
      Begin C1Chart2D8.Chart2D chtPic 
         Height          =   705
         Index           =   0
         Left            =   180
         TabIndex        =   34
         Top             =   150
         Width           =   615
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   1085
         _ExtentY        =   1244
         _StockProps     =   0
         ControlProperties=   "frmPatientReprotFind.frx":687F
      End
   End
   Begin VB.PictureBox picComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   11670
      ScaleHeight     =   4620
      ScaleWidth      =   3495
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3540
      Width           =   3495
      Begin VB.TextBox txtSignificance 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   88
         Top             =   3330
         Width           =   3255
      End
      Begin VB.TextBox txtDiagnose 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1860
         Width           =   3255
      End
      Begin VB.TextBox txtComment 
         Appearance      =   0  'Flat
         Height          =   1275
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   300
         Width           =   3255
      End
      Begin VB.Label lblSignificance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "临床意义:"
         Height          =   180
         Left            =   30
         TabIndex        =   89
         Top             =   3090
         Width           =   810
      End
      Begin VB.Label lblDiagnose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊断:"
         Height          =   180
         Left            =   30
         TabIndex        =   31
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注:"
         Height          =   180
         Left            =   30
         TabIndex        =   29
         Top             =   60
         Width           =   450
      End
   End
   Begin VB.PictureBox PICContrast 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   11520
      ScaleHeight     =   3015
      ScaleWidth      =   3795
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1020
      Width           =   3795
      Begin VB.PictureBox PicContrast_Top 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   30
         ScaleHeight     =   975
         ScaleWidth      =   3150
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   30
         Width           =   3150
         Begin VB.TextBox txtMaxDay 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            TabIndex        =   24
            Text            =   "30"
            Top             =   60
            Width           =   705
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFContrast 
            Height          =   1335
            Left            =   60
            TabIndex        =   25
            Top             =   420
            Width           =   2265
            _cx             =   3995
            _cy             =   2355
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
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
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483635
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   2
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   350
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
            Editable        =   2
            ShowComboButton =   0
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
         Begin VB.Label lblMaxDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "最大跟踪天数:"
            Height          =   180
            Left            =   90
            TabIndex        =   27
            Top             =   90
            Width           =   1170
         End
         Begin VB.Label lblContrast 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "刷新"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F56C58&
            Height          =   180
            Left            =   2040
            TabIndex        =   26
            Top             =   90
            Width           =   360
         End
      End
      Begin VB.PictureBox PicContrast_Bottom 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCDBD8&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1635
         Left            =   0
         ScaleHeight     =   1635
         ScaleWidth      =   5280
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1110
         Width           =   5280
         Begin VB.OptionButton optContrast 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCDBD8&
            Caption         =   "变异率(&1)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   20
            Top             =   8
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.OptionButton optContrast 
            Appearance      =   0  'Flat
            BackColor       =   &H00FCDBD8&
            Caption         =   "结果值(&2)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2565
            TabIndex        =   19
            Top             =   8
            Width           =   1500
         End
         Begin C1Chart2D8.Chart2D chtContrast 
            Height          =   975
            Left            =   60
            TabIndex        =   32
            Top             =   300
            Width           =   1005
            _Version        =   524288
            _Revision       =   7
            _ExtentX        =   1773
            _ExtentY        =   1720
            _StockProps     =   0
            ControlProperties=   "frmPatientReprotFind.frx":6E14
         End
         Begin VB.Label lblCht 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "图形内容"
            Height          =   180
            Left            =   30
            TabIndex        =   22
            Top             =   60
            Width           =   720
         End
      End
   End
   Begin VB.Frame FraCR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4260
      Left            =   10890
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   2490
      Width           =   45
   End
   Begin VB.Frame FraLC 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Left            =   5745
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   1635
      Width           =   45
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   9660
      ScaleHeight     =   4335
      ScaleWidth      =   3795
      TabIndex        =   3
      Top             =   960
      Width           =   3795
      Begin XtremeSuiteControls.TabControl TabPage 
         Height          =   3495
         Left            =   300
         TabIndex        =   16
         Top             =   120
         Width           =   3135
         _Version        =   589884
         _ExtentX        =   5530
         _ExtentY        =   6165
         _StockProps     =   64
         Enabled         =   -1  'True
      End
      Begin VB.Frame fraRight 
         Height          =   2595
         Left            =   1050
         TabIndex        =   14
         Top             =   900
         Width           =   1815
      End
   End
   Begin VB.PictureBox picCenter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7965
      Left            =   1020
      ScaleHeight     =   7965
      ScaleWidth      =   11235
      TabIndex        =   2
      Top             =   -150
      Width           =   11235
      Begin VB.PictureBox picRpt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1140
         Left            =   240
         ScaleHeight     =   1140
         ScaleWidth      =   1395
         TabIndex        =   86
         Top             =   0
         Visible         =   0   'False
         Width           =   1395
         Begin SHDocVwCtl.WebBrowser webSub 
            Height          =   690
            Left            =   120
            TabIndex        =   87
            Top             =   330
            Width           =   810
            ExtentX         =   1429
            ExtentY         =   1217
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
      End
      Begin VB.PictureBox PicNegative 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   7665
         Left            =   -720
         ScaleHeight     =   7635
         ScaleWidth      =   7005
         TabIndex        =   59
         Top             =   -390
         Visible         =   0   'False
         Width           =   7035
         Begin VB.Frame frmChe 
            Caption         =   "结果选择"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   1650
            TabIndex        =   76
            Top             =   1080
            Width           =   5250
            Begin VB.CheckBox chkMicroscope 
               BackColor       =   &H80000004&
               Caption         =   "镜检结果"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   3810
               TabIndex        =   81
               Top             =   300
               Width           =   1305
            End
            Begin VB.CheckBox chkNoGerm 
               BackColor       =   &H80000004&
               Caption         =   "无细菌生长"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1935
               TabIndex        =   80
               Top             =   300
               Width           =   1815
            End
            Begin VB.CheckBox chkPathopoiesiaGerm 
               BackColor       =   &H80000004&
               Caption         =   "无致病菌生长"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   60
               TabIndex        =   79
               Top             =   300
               Width           =   1815
            End
            Begin VB.OptionButton optReport 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "阳性"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   0
               Left            =   1935
               TabIndex        =   78
               Top             =   600
               Width           =   885
            End
            Begin VB.OptionButton optReport 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Caption         =   "阴性"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Index           =   1
               Left            =   60
               TabIndex        =   77
               Top             =   600
               Value           =   -1  'True
               Width           =   885
            End
         End
         Begin VB.Frame frmNom 
            Height          =   2655
            Left            =   120
            TabIndex        =   69
            Top             =   120
            Width           =   5250
            Begin VB.TextBox txtNormalMicrobes 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   1050
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   72
               Top             =   1800
               Width           =   4065
            End
            Begin VB.TextBox txtNoFindMicrobe 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   750
               Left            =   1050
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   71
               Top             =   975
               Width           =   4065
            End
            Begin VB.TextBox txtNormalMicrobe 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   690
               Left            =   1050
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   70
               Top             =   210
               Width           =   4065
            End
            Begin VB.Label lblNormalMicrobes 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "补充描述"
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
               Left            =   60
               TabIndex        =   75
               Top             =   1800
               Width           =   960
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "未 检 出"
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
               Left            =   60
               TabIndex        =   74
               Top             =   930
               Width           =   960
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "常规结果"
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
               Left            =   60
               TabIndex        =   73
               Top             =   210
               Width           =   960
            End
         End
         Begin VB.Frame fraOne 
            Caption         =   "镜检结果"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   1740
            TabIndex        =   62
            Top             =   2160
            Width           =   5250
            Begin VB.TextBox txtMicroscopeFinded 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   1110
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   65
               Top             =   690
               Width           =   3915
            End
            Begin VB.TextBox txtMicroscopeNOFind 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   1110
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   64
               Top             =   1350
               Width           =   3915
            End
            Begin VB.TextBox txtMicroscope 
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1110
               Locked          =   -1  'True
               TabIndex        =   63
               Text            =   "显微镜检查"
               Top             =   270
               Width           =   3915
            End
            Begin VB.Label lblMicroscopeFinded 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "镜检检出"
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
               Left            =   90
               TabIndex        =   68
               Top             =   660
               Width           =   960
            End
            Begin VB.Label lblMicroscopeNOFind 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "未 检 出"
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
               Left            =   60
               TabIndex        =   67
               Top             =   1290
               Width           =   960
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "通过设备"
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
               Left            =   90
               TabIndex        =   66
               Top             =   300
               Width           =   960
            End
         End
         Begin VB.Frame fraTwo 
            Caption         =   "评语"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   1230
            TabIndex        =   60
            Top             =   4260
            Width           =   5340
            Begin VB.TextBox txtGermComment 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1665
               Left            =   90
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   61
               Top             =   270
               Width           =   5115
            End
         End
      End
      Begin VB.Frame fraCenter 
         Height          =   4305
         Left            =   2400
         TabIndex        =   13
         Top             =   3600
         Width           =   8085
         Begin VB.PictureBox picMicrobePositive 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4095
            Left            =   4080
            ScaleHeight     =   4095
            ScaleWidth      =   2865
            TabIndex        =   37
            Top             =   150
            Visible         =   0   'False
            Width           =   2865
            Begin VB.TextBox txtMicrobePositiveComment 
               Appearance      =   0  'Flat
               Height          =   855
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   43
               Top             =   3030
               Width           =   2715
            End
            Begin VSFlex8Ctl.VSFlexGrid VsfAntibiotic 
               Height          =   975
               Left            =   30
               TabIndex        =   38
               Top             =   1770
               Width           =   2745
               _cx             =   4842
               _cy             =   1720
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   16706793
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483635
               FloodColor      =   192
               SheetBorder     =   16777215
               FocusRect       =   0
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   250
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
               ShowComboButton =   0
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
            Begin VSFlex8Ctl.VSFlexGrid VsfMicrobe 
               Height          =   1275
               Left            =   30
               TabIndex        =   41
               Top             =   240
               Width           =   2745
               _cx             =   4842
               _cy             =   2249
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   16706793
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483635
               FloodColor      =   192
               SheetBorder     =   16777215
               FocusRect       =   0
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   250
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
               ShowComboButton =   0
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
            Begin VB.Label lblMicrobePositiveComment 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "评语:"
               Height          =   180
               Left            =   30
               TabIndex        =   42
               Top             =   2820
               Width           =   450
            End
            Begin VB.Label lblAntibiotic 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抗生素:"
               Height          =   180
               Left            =   30
               TabIndex        =   40
               Top             =   1560
               Width           =   630
            End
            Begin VB.Label lblMicrobe 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "细菌:"
               Height          =   180
               Left            =   30
               TabIndex        =   39
               Top             =   30
               Width           =   450
            End
         End
         Begin VB.PictureBox picGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1665
            Left            =   780
            ScaleHeight     =   1665
            ScaleWidth      =   1905
            TabIndex        =   35
            Top             =   450
            Visible         =   0   'False
            Width           =   1905
            Begin VB.CheckBox chkGroup 
               Caption         =   "显示组合项目"
               Height          =   255
               Left            =   60
               TabIndex        =   82
               Top             =   0
               Width           =   1665
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfCenter 
               Height          =   1035
               Left            =   180
               TabIndex        =   36
               Top             =   450
               Width           =   1605
               _cx             =   2831
               _cy             =   1826
               Appearance      =   0
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   16706793
               ForeColorSel    =   0
               BackColorBkg    =   16777215
               BackColorAlternate=   16777215
               GridColor       =   -2147483636
               GridColorFixed  =   -2147483636
               TreeColor       =   -2147483635
               FloodColor      =   192
               SheetBorder     =   16777215
               FocusRect       =   0
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   3
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   250
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
               ShowComboButton =   0
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
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   45
      ScaleHeight     =   4275
      ScaleWidth      =   6540
      TabIndex        =   1
      Top             =   2475
      Width           =   6540
      Begin VB.Frame fraLeft 
         Height          =   4335
         Left            =   15
         TabIndex        =   12
         Top             =   60
         Width           =   6420
         Begin VB.CheckBox chkAudit 
            Appearance      =   0  'Flat
            Caption         =   "未出"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   4290
            TabIndex        =   56
            Top             =   120
            Width           =   690
         End
         Begin VB.CheckBox chkSource 
            Appearance      =   0  'Flat
            Caption         =   "未知"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   54
            Top             =   120
            Value           =   1  'Checked
            Width           =   690
         End
         Begin VB.CheckBox chkSource 
            Appearance      =   0  'Flat
            Caption         =   "门诊"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   53
            Top             =   120
            Value           =   1  'Checked
            Width           =   690
         End
         Begin VB.CheckBox chkSource 
            Appearance      =   0  'Flat
            Caption         =   "住院"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   1530
            TabIndex        =   52
            Top             =   120
            Value           =   1  'Checked
            Width           =   690
         End
         Begin VB.CheckBox chkSource 
            Appearance      =   0  'Flat
            Caption         =   "院外"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   3
            Left            =   2220
            TabIndex        =   51
            Top             =   120
            Value           =   1  'Checked
            Width           =   690
         End
         Begin VB.CheckBox chkSource 
            Appearance      =   0  'Flat
            Caption         =   "体检"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   4
            Left            =   2910
            TabIndex        =   50
            Top             =   120
            Value           =   1  'Checked
            Width           =   690
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfLeft 
            Height          =   4035
            Left            =   210
            TabIndex        =   15
            Top             =   360
            Width           =   6045
            _cx             =   10663
            _cy             =   7117
            Appearance      =   0
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16706793
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   16777215
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483635
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   250
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
            ExplorerBar     =   5
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   0
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
         Begin VB.CheckBox chkAudit 
            Appearance      =   0  'Flat
            Caption         =   "已出"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   3600
            TabIndex        =   55
            Top             =   120
            Value           =   1  'Checked
            Width           =   690
         End
      End
   End
   Begin MSComctlLib.ImageList imgVsf 
      Left            =   1110
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":73A9
            Key             =   "选择"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":DC0B
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":1446D
            Key             =   "老版"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":1ACCF
            Key             =   "新版"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":21531
            Key             =   "序号"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgPrint 
      Left            =   2160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   0
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatientReprotFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const const_PicRectBackColour As Long = &HE0E0E0
Private mblnShow As Boolean                                         '窗体是否显示
Private mlngKey As Long                                                 '当前选择的标本ID
Private mlngPatientID As Long                                              '病人ID
Private mdteReportDate As Date                                             '报告时间
Private mlngGetPatientID As Long                                        '上级传来的病人ID
Private mSampleShowColour As SampleValShowColour                    '结果显示颜色
Private mstrPrivs As String                                         '传入的上级的权限
Private mblnLoad  As Boolean                                        '窗体是否第一次显示
Private mlngValueC As Long                                              '微生物结果次数
Private mintVer As Integer                                              '版本，25-新版 10-老板
Private mlngSelRow As Long                                              '之前选中行
Private mintIn As Integer                                               '实验室查看
Private mlngPicLeftWidth As Long         '左侧布局宽度
Private mlngPicCenterWidth As Long       '中间布宽度


Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Private mstrTag As String           '三方报告打印

Private Type SampleValShowColour                                    '结果颜色显示

    正常 As Double
    偏低 As Double
    偏高 As Double
    异常 As Double
    警示偏高 As Double
    警示偏低 As Double
    复查偏高 As Double
    复查偏低 As Double
End Type
Private mobjFSO As New Scripting.FileSystemObject    'FSO对象
Private mObjImg As Object

Private mObjIco As IcoObject                                            '图标对象

'定义图标对象
'如果在frmWorkBaseReprot,frmWorkBaseReprotFind,frmWorkBaseAuditingSample
'三个窗体中的任意一个窗体的图标控件中添加了图片,则这三个窗体都需要同步跟新
Private Enum mIcoIndex
    选择 = 1
    打印
    新版
    老版
    序号
End Enum

Private Type IcoObject
    Obj序号  As Object
    Obj打印 As Object
    Obj选择 As Object
    Obj新版 As Object
    Obj老版 As Object
End Type

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Sub setIcoObj(mobjImgList As ImageList, Optional strErr As String)
    '初始化图标类型
    If mobjImgList Is Nothing Or Not mObjIco.Obj打印 Is Nothing Then
        Exit Sub
    End If
    Set mObjIco.Obj序号 = mobjImgList.ListImages(mIcoIndex.序号).ExtractIcon
    Set mObjIco.Obj打印 = mobjImgList.ListImages(mIcoIndex.打印).ExtractIcon
    Set mObjIco.Obj选择 = mobjImgList.ListImages(mIcoIndex.选择).ExtractIcon
    Set mObjIco.Obj新版 = mobjImgList.ListImages(mIcoIndex.新版).ExtractIcon
    Set mObjIco.Obj老版 = mobjImgList.ListImages(mIcoIndex.老版).ExtractIcon
End Sub

Private Sub setIcoFree()
    '释放资源
    Set mObjIco.Obj序号 = Nothing
    Set mObjIco.Obj打印 = Nothing
    Set mObjIco.Obj选择 = Nothing
    Set mObjIco.Obj新版 = Nothing
    Set mObjIco.Obj老版 = Nothing
End Sub

Private Sub cboDept_GotFocus()
    Call selAllText(cboDept)
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With Me.cboDept
            Call GetDept(0, .Text)
        End With
    End If
End Sub

Private Sub cboPrint_Click()
    If mblnLoad = True Then
        '窗体显示后选择打印状态刷新病人列表
       Call ReadPatientList(1)
    Else
        mblnLoad = True
    End If
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_SelAll
            VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("选择"), 1
        Case ConMenu_Browse_ClsAll
            VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("选择"), 2
        Case conFun_Sample_Auditing     '复核
            Call AuditingSample(1)
        Case conFun_Sample_unAuditing     '取消复核
            Call AuditingSample(2)
        Case ConMenu_Browse_Find
            Call ReadPatientList(1)
        Case ConMenu_Browse_Print
            If mstrTag <> "" Then
                beginPrint
            Else
                BatchPrint
            End If
        Case ConMenu_Browse_PrintSet
            If mstrTag <> "" Then
                cdgPrint.ShowPrinter
            Else
                If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("版本"))) = 25 Then
                    PrintReport Me, mlngKey, 3
                Else
                    PtintOldReport Me, mlngKey, , 3
                End If
            End If
        Case ConMenu_Browse_PrintView   '预览
            If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("版本"))) = 25 Then
                PrintReport Me, mlngKey, 1
            Else
                PtintOldReport Me, mlngKey, , 1
            End If
        Case ConMenu_Browse_Exit
            Unload Me
        Case ConMenu_pop_SampleCode
            lblNo.Caption = "条码号↓"
        Case ConMenu_pop_Out
            lblNo.Caption = "门诊号↓"
        Case ConMenu_pop_In
            lblNo.Caption = "住院号↓"
        Case ConMenu_pop_bed
            lblNo.Caption = "  床号↓"
        Case ConMenu_pop_PatiCard
            lblNo.Caption = "就诊卡↓"
        Case ConMenu_Browse_unPrint     '重置打印
            Call ResetPrintType
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
            Call ExePlugIn(Control.Parameter, mlngKey)
    End Select
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/6/13
'功    能:重置自助机报告打印次数
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub ResetPrintType()
          Dim strSQL As String

1         On Error GoTo ResetPrintType_Error

2         If mlngKey <= 0 Then Exit Sub
3         strSQL = "Zl_检验报告打印_Edit(2," & mlngKey & ",2)"
4         Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
5         SaveDBLog 18, 6, mlngKey, "打印", "重置自助机打印次数", 2500, "临床实验室管理"

6         MsgBox "自助机打印状态已重置", vbInformation, Me.Caption


7         Exit Sub
ResetPrintType_Error:
8         Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(ResetPrintType)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
9         Err.Clear

End Sub

Private Sub AuditingSample(ByVal intType As Integer)
          '复核/取消复核
          'intType    1=复核,2=取消复核

          Dim strSQL As String
          Dim lngSampleKey As String  '标本id

1         On Error GoTo AuditingSample_Error

2         With Me.vsfLeft
3             If .Row > 0 Then
4                 lngSampleKey = Val(.TextMatrix(.Row, .ColIndex("ID")))
5             Else
6                 MsgBox "请选中传染病记录", vbInformation, Me.Caption
7                 Exit Sub
8             End If
9         End With

10        strSQL = "Zl_检验传染病复核_Edit(" & intType & "," & lngSampleKey & ",'" & UserInfo.Name & "')"
11        Call ComExecuteProc(Sel_Lis_DB, strSQL, "传染病报告复核")

12        If vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("版本")) = 25 Then
13            SaveDBLog 18, 6, Val(lngSampleKey), IIf(intType = 1, "复核", "取消复核"), IIf(intType = 1, "复核", "取消复核"), 2500, "临床实验室管理"
14        End If

          '刷新列表
15        mlngPatientID = 0
16        Call ReadPatientList(1)
17        Call vsfLeft_Click


18        Exit Sub
AuditingSample_Error:
19        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(AuditingSample)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
20        Err.Clear
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
      Select Case Control.ID
        Case ConMenu_Browse_PrintView   '预览
            If mstrTag = "" Then
                Control.Enabled = True
            Else
                Control.Enabled = False
            End If
    End Select
End Sub

Private Sub chkAudit_Click(Index As Integer)
   Call ReadPatientList(1)
End Sub

Private Sub chkGroup_Click()
    Dim i As Integer

    With Me.vsfCenter
        If .ColIndex("序号") < 0 Then Exit Sub
        For i = 1 To .Rows - 1
            If .Cell(flexcpFontBold, i, .ColIndex("序号")) = True Then
                .RowHidden(i) = IIf(Me.chkGroup.value = 1, False, True)
            End If
        Next
    End With
End Sub

Private Sub chkMicroscope_Click()
    PicNegative_Resize
End Sub

Private Sub chkSource_Click(Index As Integer)
   Call ReadPatientList(1)
End Sub

Private Sub Command1_Click()
    Call frmPaitReport.ShowMe(Me, 705845, "ICU病人;按科室或病区展示病人;本科病人;病案审查提交;参数设置;打印首页;湖南省病案首页;湖南省中医病案首页;护理监护;会诊病人;基本;抗菌药物越级使用汇总表;抗菌药物越级使用明细表;临床自管药;全院病人;审查反馈处理;首页基本信息;首页整理;四川省西医首页;四川省中医首页;危急值处理;修改手术等级;修改医疗付款方式;药占比查询;预约挂号;预约挂号单;云南省西医首页;云南省中医首页;中医病案首页;住院一览", 57, 57, 2, 1, , True, , True)
End Sub

Private Sub Form_Activate()
    If mblnShow = False Then
        InitFace
        If mlngPatientID = 0 Then Call ReadPatientList(-1, True)
        mblnShow = True
    End If
End Sub

Private Sub Form_Load()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim strPicWidth As String
    
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_SelAll, "全选")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_ClsAll, "全清")

        Set cbrControl = .Add(xtpControlButton, conFun_Sample_Auditing, "复核")
        cbrControl.Visible = False
        cbrControl.Enabled = False: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conFun_Sample_unAuditing, "取消复核")
        cbrControl.Visible = False
        cbrControl.Enabled = False

        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "查找"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "打印")
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "打印设置  ")
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_unPrint, "重置打印  ")
            cbrControl.Visible = InStr(mstrPrivs, "重置自助机报告打印次数") > 0
        End With
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "预览")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "退出"): cbrControl.BeginGroup = True
    End With

    '创建插件按钮
    Call CreatePlugInButton(cbrToolBar)

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next



    '列表
    With Me.TabPage
        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True

        '备注
        .InsertItem 1, "备注", picComment.hWnd, ConTab_Sample_Comment
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        '列表
        .InsertItem 2, "历次", PICContrast.hWnd, ConTab_Sample_History
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        .InsertItem 3, "图像", PicPic.hWnd, ConTab_Sample_Comment
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        .Item(0).Selected = True
    End With

    With cboPrint
        .AddItem "所有"
        .AddItem "已打印"
        .AddItem "未打印"
        .ListIndex = 0
    End With

    dtpS = Now - 7
    dtpE = Now
    picLeft.Width = GetSetWith(1)

    Me.chkGroup.value = Val(ComGetPara(Sel_Lis_DB, "是否显示组合项目", gSysInfo.SysNo, gSysInfo.ModlNo, 1))
    strPicWidth = GetSetting("ZLSOFT", "私有模块\" & UserInfo.DBUser & "\" & App.EXEName & "\" & Me.Name, "PICWIDTH", "")
    If strPicWidth <> "" Then
        mlngPicLeftWidth = Val(Split(strPicWidth, ";")(0))
        mlngPicCenterWidth = Val(Split(strPicWidth, ";")(1))
    End If
    
    ReadSampleBacteriology 0
    ReadSampleBacteriology 0
    ReadSampleVal 0

End Sub

Private Function GetSetWith(ByVal intType As Integer) As Long
    '读取/设置窗体左边部分的宽度
    '1-读取,2-设置
    If intType = 1 Then
        GetSetWith = ComGetPara(Sel_Lis_DB, "检验报告信息宽", 2500, 2500, "5000")
    ElseIf intType = 2 Then
        Call ComSetPara(Sel_Lis_DB, "检验报告信息宽", picLeft.Width, 2500, 2500)
    End If
End Function

Public Sub PicDrowBorder(Picobj As PictureBox, Optional lngLineColour As Long = -1)
    '功能       画图片边框
    On Error Resume Next
    With Picobj
        .AutoRedraw = True
        .Cls
        .DrawWidth = 2

        If lngLineColour = -1 Then
            .ForeColor = const_PicRectBackColour
        Else
            .ForeColor = lngLineColour
        End If
        Picobj.Line (25, 25)-(.Width - 50, .Height - 50), , B
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With Me.picFilter
        .Top = 430
        .Left = 15
        .Width = Me.ScaleWidth - .Left - 15
    End With

    With picLeft
        .Top = picFilter.Top + picFilter.Height
        .Left = 15
        .Height = Me.ScaleHeight - picFilter.Top - picFilter.Height
        If mlngPicLeftWidth > 0 Then
            .Width = mlngPicLeftWidth
        End If
    End With

    With FraLC
        .Top = picLeft.Top
        .Height = picLeft.Height
        .Left = picLeft.Left + picLeft.Width
    End With

    With Me.picCenter
        .Top = picLeft.Top
        .Height = picLeft.Height
        .Left = FraLC.Left + FraLC.Width
        If mlngPicCenterWidth > 0 Then
            .Width = mlngPicCenterWidth
        End If
    End With

    With FraCR
        .Top = picLeft.Top
        .Height = picLeft.Height
        .Left = picCenter.Left + picCenter.Width
    End With

    With picRight
        .Top = picLeft.Top
        .Height = picLeft.Height
        .Left = FraCR.Left + FraCR.Width
        .Width = Me.ScaleWidth - .Left - 30
    End With

    Call PicDrowBorder(picFilter)
    Call PicDrowBorder(picLeft)
    Call PicDrowBorder(picCenter)
    Call PicDrowBorder(picRight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnShow = False
    mlngKey = 0
    mlngValueC = 0
    mintVer = 0
    Set mObjImg = Nothing
    mblnLoad = False
    mintIn = 0
    mlngPatientID = 0
    mlngGetPatientID = 0
    Set mobjFSO = Nothing
    Call GetSetWith(2)

    Call ComSetPara(Sel_Lis_DB, "是否显示组合项目", Me.chkGroup.value, gSysInfo.SysNo, gSysInfo.ModlNo)
    Call setIcoFree
    '保存界面宽度
    Call SaveSetting("ZLSOFT", "私有模块\" & UserInfo.DBUser & "\" & App.EXEName & "\" & Me.Name, "PICWIDTH", picLeft.Width & ";" & picCenter.Width)
End Sub

Private Sub FraCR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeftColl As New Collection, Rightcoll As New Collection
    If Button = vbLeftButton Then
        LeftColl.Add Me.picCenter
        Rightcoll.Add Me.picRight
        Call SplitWE(LeftColl, Me.FraCR, Rightcoll, X, 1000)
        Set LeftColl = Nothing
        Set Rightcoll = Nothing
    End If
End Sub

Private Sub FraLC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeftColl As New Collection, Rightcoll As New Collection
    If Button = vbLeftButton Then
        LeftColl.Add Me.picLeft
        Rightcoll.Add Me.picCenter
        Call SplitWE(LeftColl, Me.FraLC, Rightcoll, X, 1000)
        Set LeftColl = Nothing
        Set Rightcoll = Nothing
    End If
End Sub

Private Sub lblContrast_Click()
    Call ReadContrastToVsf
End Sub



Private Sub lblNo_Click()
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    Dim vPoint As POINTAPI
    On Error Resume Next

    Set objPopup = Me.cbrthis.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_SampleCode, "条码号")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Out, "门诊号")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_In, "住院号")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_bed, "床号")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_PatiCard, "就诊卡")
    End With
    vPoint.X = lblNo.Left / Screen.TwipsPerPixelX
    vPoint.Y = (lblNo.Top + lblNo.Height + 30) / Screen.TwipsPerPixelY
    ClientToScreen picFilter.hWnd, vPoint
    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
End Sub



Private Sub optContrast_Click(Index As Integer)
    Call VSFContrast_SelChange
End Sub

Private Sub picCenter_Resize()
    On Error Resume Next
    With fraCenter
        .Top = -40
        .Left = 0
        .Width = picCenter.ScaleWidth
        .Height = picCenter.ScaleHeight
    End With
    With picGeneral
        .Top = 120
        .Left = 30
        .Width = picCenter.Width - 70
        .Height = picCenter.Height - 150
    End With


    With picMicrobePositive
        .Top = 120
        .Left = 30
        .Width = picCenter.Width - 70
        .Height = picCenter.Height - 150
    End With
    With PicNegative
        .Top = 120
        .Left = 30
        .Width = picCenter.Width - 70
        .Height = picCenter.Height - 200
    End With

    With picRpt
        .Top = 120
        .Left = 30
        .Width = picCenter.Width - 70
        .Height = picCenter.Height - 200
    End With
End Sub

Private Sub picComment_Resize()
    On Error Resume Next
    lblComment.Move 20, 50
    txtComment.Move lblComment.Left, lblComment.Top + lblComment.Height + 50, picComment.ScaleWidth - 40, picComment.ScaleHeight / 3 - 280
    lblDiagnose.Move lblComment.Left, txtComment.Top + txtComment.Height + 50
    txtDiagnose.Move lblComment.Left, lblDiagnose.Top + lblDiagnose.Height + 50, txtComment.Width, picComment.ScaleHeight / 3 - 280
    lblSignificance.Move lblComment.Left, txtDiagnose.Top + txtDiagnose.Height + 50
    txtSignificance.Move lblComment.Left, lblSignificance.Top + lblSignificance.Height + 50, txtComment.Width, picComment.ScaleHeight / 3 - 280
End Sub

Private Sub PicContrast_Bottom_Resize()

    On Error Resume Next
    With Me.chtContrast
        .Top = lblCht.Top + lblCht.Height + 75
        .Left = 0
        .Width = Me.PicContrast_Bottom.ScaleWidth
        .Height = Me.PicContrast_Bottom.ScaleHeight
    End With
End Sub

Private Sub PICContrast_Resize()
    On Error Resume Next
    With Me.PicContrast_Top
        .Top = 0
        .Left = 0
        .Width = Me.PICContrast.ScaleWidth
        .Height = Me.PICContrast.ScaleHeight / 2
    End With
    With Me.PicContrast_Bottom
        .Top = PicContrast_Top.Top + PicContrast_Top.Height + 25
        .Left = 0
        .Width = Me.PicContrast_Top.Width
        .Height = Me.PICContrast.ScaleHeight - .Top
    End With
End Sub

Private Sub PicContrast_Top_Resize()
    On Error Resume Next
    With Me.VSFContrast
        .Top = Me.lblMaxDay.Top + lblMaxDay.Height + 80
        .Left = 0
        .Width = PicContrast_Top.ScaleWidth
        .Height = PicContrast_Top.ScaleHeight - .Top
    End With
End Sub

Private Sub picFilter_Resize()
    With fraTop
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth - 50
        .Height = Me.picFilter.ScaleHeight + 10
    End With
End Sub



Private Sub picGeneral_Resize()
    On Error Resume Next
    With chkGroup
        .Left = 0
        .Top = 0
        .Width = Me.picGeneral.Width
    End With
    With vsfCenter
        .Top = Me.chkGroup.Top + Me.chkGroup.Height
        .Left = 0
        .Width = picGeneral.ScaleWidth
        .Height = picGeneral.ScaleHeight
    End With
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With fraLeft
        .Top = -40
        .Left = 0
        .Width = picLeft.ScaleWidth
        .Height = picLeft.ScaleHeight
    End With
    With vsfLeft
        .Top = 120 + chkSource(0).Top + chkSource(0).Height
        .Left = 30
        .Width = fraLeft.Width - 70
        .Height = fraLeft.Height - 150 - chkSource(0).Top - chkSource(0).Height
    End With
End Sub

Private Sub picMicrobePositive_Resize()
    On Error Resume Next
    With lblMicrobe
        .Top = 50
        .Left = 30
    End With

    With VsfMicrobe
        .Top = lblMicrobe.Top + lblMicrobe.Height + 30
        .Left = 20
        .Width = picMicrobePositive.ScaleWidth - 40
        .Height = (picMicrobePositive.Height - ((lblMicrobe.Height + 50) * 3)) / 3
    End With

    With lblAntibiotic
        .Top = VsfMicrobe.Top + VsfMicrobe.Height + 50
        .Left = 20
        .Width = VsfMicrobe.Width
    End With

    With VsfAntibiotic
        .Top = lblAntibiotic.Top + lblAntibiotic.Height + 30
        .Left = 20
        .Width = picMicrobePositive.ScaleWidth - 40
        .Height = (picMicrobePositive.Height - ((lblMicrobe.Height + 50) * 2)) / 2
    End With

    With Me.lblMicrobePositiveComment
        .Top = VsfAntibiotic.Top + VsfAntibiotic.Height + 50
        .Left = 20
        .Width = VsfMicrobe.Width
    End With

    With Me.txtMicrobePositiveComment
        .Top = lblMicrobePositiveComment.Top + lblMicrobePositiveComment.Height + 30
        .Left = 20
        .Width = picMicrobePositive.ScaleWidth - 40
        .Height = picMicrobePositive.Height - .Top - 20
    End With
End Sub

Private Sub PicNegative_Resize()
    On Error Resume Next
    With frmNom
        .Top = 20
        .Left = 60
        .Width = PicNegative.ScaleWidth - 60
    End With
    txtNormalMicrobe.Width = frmNom.Width - Label21.Width - 300
    txtNoFindMicrobe.Width = txtNormalMicrobe.Width
    txtNormalMicrobes.Width = txtNormalMicrobe.Width
    With frmChe
        .Top = frmNom.Top + frmNom.Height + 20
        .Left = 60
         .Width = PicNegative.ScaleWidth - 60
    End With

    If chkMicroscope.value = 1 Then
        fraOne.Visible = True
        With fraOne
            .Top = frmChe.Top + frmChe.Height + 20
            .Left = 60
             .Width = PicNegative.ScaleWidth - 60
        End With
        txtMicroscope.Width = fraOne.Width - Label1.Width - 500
        txtMicroscopeFinded.Width = txtMicroscope.Width
        txtMicroscopeNOFind.Width = txtMicroscope.Width
        With fraTwo
            .Top = fraOne.Top + fraOne.Height + 20
            .Left = 60
            .Height = PicNegative.ScaleHeight - frmNom.Height - frmChe.Height - fraOne.Height - 300
             .Width = PicNegative.ScaleWidth - 60
        End With

    Else
        fraOne.Visible = False
        With fraTwo
            .Top = frmChe.Top + frmChe.Height + 20
            .Left = 60
            .Height = PicNegative.ScaleHeight - frmNom.Height - frmChe.Height - 300
             .Width = PicNegative.ScaleWidth - 60
        End With
    End If
    txtGermComment.Width = fraTwo.Width - 300
    txtGermComment.Height = fraTwo.Height - 800
End Sub


Private Sub PicPic_Resize()
    ImageTypeSet 9
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With fraRight
        .Top = -40
        .Left = 0
        .Width = picRight.ScaleWidth
        .Height = picRight.ScaleHeight
    End With
    With TabPage
        .Top = 120
        .Left = 30
        .Width = picRight.Width - 70
        .Height = picRight.Height - 220
    End With
End Sub

Private Sub ReadPatientList(lngID As Long, Optional ByVal blnLoadFrm As Boolean)
      '功能           按条件读出病人列表
      '               blnLoadFrm=True  初始化窗体时不加载数据，只设置VSF
          Dim rsTmp As ADODB.Recordset, rsOldLisData As ADODB.Recordset
          Dim strSQL As String
          Dim lngKey As Long
          Dim strDepts As String
          Dim strDept As String
          Dim lngPatiID As Long
          Dim strTemp As String
          Dim strWhere As String
          Dim lngLoop As Long
          Dim strTitle As String   '列表
          Dim var_tmp As Variant
          Dim var_SubTmp As Variant
          Dim blnReadData As Boolean
          Dim strTiredFind As String
          Dim strFindSQL As String

          '查找这先清空
1         On Error GoTo ReadPatientList_Error

          '获取参数值
2         strTitle = ComGetPara(Sel_Lis_DB, "报告中心显示列", 2500, 1013)

3         If Trim(txtPatiNo <> "") Then
4             If lblNo.Caption = "就诊卡↓" Then
5                 strSQL = "select 病人id from 病人信息 where 就诊卡号  = [1] "
6                 Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查找病人就诊卡", txtPatiNo)
7                 If rsTmp.RecordCount > 0 Then
8                     lngPatiID = rsTmp("病人id")
9                 Else
10                    lngPatiID = -1
11                End If
12            End If
13        End If

14        mlngKey = 0
15        ReadSampleBacteriology 0
16        ReadSampleBacteriology 0
17        ReadSampleVal 0
18        Call setIcoObj(imgVsf)

          '高峰时段限制查询
19        blnReadData = True
20        If blnLoadFrm = False Then
21            If mintIn = 0 Then
22                If Not funCheckRushHours(2500, 1013, "检验报告查询中心", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59")) Then
23                    blnReadData = False
24                Else
25                    blnReadData = True
26                End If
27            End If
28        End If

29        If blnReadData = True Then


30            strSQL = "Select a.Id, a.选择, a.姓名, a.性别, a.报告, a.打印, a.年龄, a.申请项目, a.住院号, a.床号, a.申请时间, a.病人id, a.核收时间, a.审核时间, a.备注, a.诊断, a.微生物," & vbNewLine & _
                     "          a.阳性报告, a.检验人, a.审核人, a.病人来源, a.申请人,申请科室, a.版本, a.结果次数, b.操作时间 打印时间, b.操作员 打印人,a.是否传染病,a.复核人,a.复核时间,a.样本条码,a.采样人,a.采样时间,a.病历号" & vbNewLine & _
                     "   From (Select a.Id, 0 选择, a.姓名, Decode(a.性别, 1, '男', 2, '女', 9, '未知', '') 性别, Decode(a.审核人, Null, '未出', '已出') 报告," & vbNewLine & _
                     "                 Decode(a.打印次数, Null, (Decode(a.医生站打印, Null, Decode(a.自助打印次数, Null, 0, 1), 1)), 1) 打印, a.年龄, c.名称 申请项目, a.住院号," & vbNewLine & _
                     "                 a.床号, b.申请时间, a.病人id, a.核收时间, a.审核时间, a.备注, a.诊断, a.微生物, a.阳性报告, a.检验人, a.审核人, Nvl(a.病人来源, 0) 病人来源, a.申请人,a.申请科室," & vbNewLine & _
                     "                 25 版本, 0 结果次数, (Select Max(ID) From 检验操作日志 D Where d.标本id = a.Id And d.操作性质 = '标本打印') 操作日志id,a.是否传染病,a.复核人,a.复核时间,a.样本条码,a.采样人,a.采样时间,a.病历号" & vbNewLine & _
                     "          From 检验报告记录 A, 检验申请组合 B, 检验组合项目 C" & vbNewLine & _
                     "          Where a.Id(+) = b.标本id And b.组合id = c.Id(+) And a.核收时间 Between [1] And [2] [条件]) A, 检验操作日志 B" & vbNewLine & _
                     "   Where a.操作日志id = b.Id(+)"

31            Select Case cboPrint.ListIndex

              Case 1
                  '已打印
32                strWhere = strWhere & "and ( a.打印次数 is not null or a.医生站打印 is not  null or a.自助打印次数 is not null ) "
33            Case 2
                  '未打印
34                strWhere = strWhere & " and a.打印次数 is null   and  a.医生站打印  is  null and a.自助打印次数 is  null "
35            End Select

36            If chkAudit(0).value = 1 And chkAudit(1).value = 0 Then
37                strWhere = strWhere & "And a.审核人 Is Not Null And a.审核时间 Is Not Null"
38            ElseIf chkAudit(0).value = 0 And chkAudit(1).value = 1 Then
39                strWhere = strWhere & "And a.审核人 Is  Null And a.审核时间 Is  Null"
40            ElseIf chkAudit(0).value = 1 And chkAudit(1).value = 1 Then

41            End If

              '传染病
42            Select Case Me.cboDiseases.Text
              Case "传染病"
43                strWhere = strWhere & " and a.是否传染病=1"
44            Case "非传染病"
45                strWhere = strWhere & " and (a.是否传染病<>1 or 是否传染病 is null)"
46            End Select

              '处理病人来源
47            strTemp = checkboxSource()
48            If strTemp <> "" Then
49                strWhere = strWhere & " and nvl(a.病人来源,0) in (" & strTemp & ")"
50            Else
                  '为选择病人来源时病人来源条件为-1
51                strWhere = strWhere & " and nvl(a.病人来源,0) in (-1)"
52            End If

53            If lngID = -1 Then
54                strWhere = strWhere & " and a.id = -1 "
55                strFindSQL = strFindSQL & " and a.id = -1 "
56            End If

57            If txtName <> "" Then
58                strWhere = strWhere & " and a.姓名 like '" & txtName & "%' "
59                strFindSQL = strFindSQL & " and a.姓名 like '" & txtName & "%' "
60            End If

61            If txtDoctor.Text <> "" Then
62                strWhere = strWhere & " and a.申请人 like '" & txtDoctor.Text & "%'"
63                strFindSQL = strFindSQL & " and a.申请人 like '" & txtDoctor.Text & "%'"
64            End If

65            If Trim(txtPatiNo <> "") Then
66                If lblNo.Caption = "住院号↓" Then
67                    strWhere = strWhere & " and a.住院号 = [3] "
68                    strFindSQL = strFindSQL & " and a.住院号 = [3] "
69                ElseIf lblNo.Caption = "门诊号↓" Then
70                    strWhere = strWhere & " and a.门诊号 = [3] "
71                    strFindSQL = strFindSQL & " and a.门诊号 = [3] "
72                ElseIf lblNo.Caption = "床号↓" Then
73                    strWhere = strWhere & " and a.床号 = [3] "
74                    strFindSQL = strFindSQL & " and a.床号 = [3] "
75                ElseIf lblNo.Caption = "就诊卡↓" Then
76                    strWhere = strWhere & " and a.HIS病人ID = [7] "
77                    strFindSQL = strFindSQL & " and a.HIS病人ID = [7] "
78                ElseIf lblNo.Caption = "条码号↓" Then
79                    strWhere = strWhere & " and a.样本条码 = [3] "
80                    strFindSQL = strFindSQL & " and a.样本条码 = [3] "
81                End If
82            End If
83            If cboDept <> "" Then
84                strDept = Mid(cboDept.Text, InStr(cboDept.Text, "-") + 1)
85                If InStr(strDept, "所有科室") > 0 Then

86                Else
87                    strWhere = strWhere & " and a.申请科室 =[6] "
88                    strFindSQL = strFindSQL & " and a.申请科室 =[6] "
89                End If
90            End If
91            DoEvents
92            If lngPatiID = 0 And mlngPatientID <> 0 Then
93                lngPatiID = mlngPatientID
94                strWhere = strWhere & " and a.HIS病人ID = [7] "
95                strFindSQL = strFindSQL & " and a.HIS病人ID = [7] "
96            End If

97            strSQL = Replace(strSQL, "[条件]", strWhere)


98            strTiredFind = "   union all    Select a.申请id  Id, 0 选择, a.姓名, a.性别, '已出' 报告, 0 打印, a.年龄, c.名称 申请项目, a.住院号, a.床号, a.申请时间, a.病人id, sysdate 核收时间, sysdate 审核时间, '' 备注, a.诊断," & vbNewLine & _
                           " 1 微生物, 3 阳性报告, a.送检人, '' 审核人, a.病人来源, a.申请人, 申请科室, 25 版本, 1 结果次数 , null 打印时间, '' 打印人, 0 是否传染病, '' 复核人, null 复核时间," & vbNewLine & _
                           " a.样本条码, '' 采样人, null 采样时间,null 病历号" & vbNewLine & _
                             "From 检验申请组合 A ,检验组合项目 C " & vbNewLine & _
                           " Where a.申请时间 Between [1] And [2] And a.申请状态 = 4 and a.组合id =c.id(+) "
99            strTiredFind = strTiredFind & strFindSQL
100           strSQL = strSQL & strTiredFind
101           strSQL = " select * from (" & strSQL & " ) order by 审核时间,病人id,id"
102           If mintIn = 1 Then
103               strSQL = Replace(strSQL, "And a.核收时间 Between [1] And [2]", "")

104           End If
105           Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入病人列表", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
                                     CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59"), txtPatiNo, mlngGetPatientID, strDepts, strDept, lngPatiID)
106       End If

107       With vsfLeft
108           If strTitle = "" Then
109               .Rows = 1
110               .Cols = 35
111               .FixedRows = 1
112               .ColKey(0) = "序号": .ColWidth(.ColIndex("序号")) = 500: .ColAlignment(.ColIndex("序号")) = flexAlignCenterCenter
113               .ColKey(1) = "id": .ColWidth(.ColIndex("id")) = 2000: .ColAlignment(.ColIndex("id")) = flexAlignCenterCenter: .ColHidden(.ColIndex("id")) = True
114               .ColKey(2) = "选择": .ColWidth(.ColIndex("选择")) = 250: .ColAlignment(.ColIndex("选择")) = flexAlignCenterCenter: .ColDataType(.ColIndex("选择")) = flexDTBoolean
115               .ColKey(3) = "打印": .ColWidth(.ColIndex("打印")) = 300: .ColAlignment(.ColIndex("打印")) = flexAlignCenterCenter
116               .ColKey(4) = "版本": .ColWidth(.ColIndex("版本")) = 300: .ColAlignment(.ColIndex("版本")) = flexAlignCenterCenter
117               .ColKey(5) = "报告": .ColWidth(.ColIndex("报告")) = 500: .ColAlignment(.ColIndex("报告")) = flexAlignCenterCenter

118               .ColKey(6) = "病人来源": .ColWidth(.ColIndex("病人来源")) = 420: .ColAlignment(.ColIndex("病人来源")) = flexAlignCenterCenter
119               .ColKey(7) = "姓名": .ColWidth(.ColIndex("姓名")) = 750: .ColAlignment(.ColIndex("姓名")) = flexAlignCenterCenter
120               .ColKey(8) = "性别": .ColWidth(.ColIndex("性别")) = 500: .ColAlignment(.ColIndex("性别")) = flexAlignCenterCenter
121               .ColKey(9) = "年龄": .ColWidth(.ColIndex("年龄")) = 500: .ColAlignment(.ColIndex("年龄")) = flexAlignCenterCenter
122               .ColKey(10) = "申请项目": .ColWidth(.ColIndex("申请项目")) = 2200: .ColAlignment(.ColIndex("申请项目")) = flexAlignCenterCenter
123               .ColKey(11) = "样本条码": .ColWidth(.ColIndex("样本条码")) = 1300: .ColAlignment(.ColIndex("样本条码")) = flexAlignCenterCenter
124               .ColKey(12) = "审核时间": .ColWidth(.ColIndex("审核时间")) = 2000: .ColAlignment(.ColIndex("审核时间")) = flexAlignCenterCenter
125               .ColKey(13) = "住院号": .ColWidth(.ColIndex("住院号")) = 750: .ColAlignment(.ColIndex("住院号")) = flexAlignCenterCenter
126               .ColKey(14) = "床号": .ColWidth(.ColIndex("床号")) = 500: .ColAlignment(.ColIndex("床号")) = flexAlignCenterCenter
127               .ColKey(15) = "申请时间": .ColWidth(.ColIndex("申请时间")) = 2000: .ColAlignment(.ColIndex("申请时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("申请时间")) = True
128               .ColKey(16) = "病人ID": .ColWidth(.ColIndex("病人ID")) = 2000: .ColAlignment(.ColIndex("病人ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("病人ID")) = True
129               .ColKey(17) = "核收时间": .ColWidth(.ColIndex("核收时间")) = 2000: .ColAlignment(.ColIndex("核收时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("核收时间")) = True
130               .ColKey(18) = "备注": .ColWidth(.ColIndex("备注")) = 2000: .ColAlignment(.ColIndex("备注")) = flexAlignCenterCenter: .ColHidden(.ColIndex("备注")) = True
131               .ColKey(19) = "诊断": .ColWidth(.ColIndex("诊断")) = 2000: .ColAlignment(.ColIndex("诊断")) = flexAlignCenterCenter: .ColHidden(.ColIndex("诊断")) = True
132               .ColKey(20) = "微生物": .ColWidth(.ColIndex("微生物")) = 2000: .ColAlignment(.ColIndex("微生物")) = flexAlignCenterCenter: .ColHidden(.ColIndex("微生物")) = True
133               .ColKey(21) = "阳性报告": .ColWidth(.ColIndex("阳性报告")) = 2000: .ColAlignment(.ColIndex("阳性报告")) = flexAlignCenterCenter: .ColHidden(.ColIndex("阳性报告")) = True
134               .ColKey(22) = "检验人": .ColWidth(.ColIndex("检验人")) = 750: .ColAlignment(.ColIndex("检验人")) = flexAlignCenterCenter
135               .ColKey(23) = "审核人": .ColWidth(.ColIndex("审核人")) = 750: .ColAlignment(.ColIndex("审核人")) = flexAlignCenterCenter
136               .ColKey(24) = "申请人": .ColWidth(.ColIndex("申请人")) = 750: .ColAlignment(.ColIndex("申请人")) = flexAlignCenterCenter
137               .ColKey(25) = "申请科室": .ColWidth(.ColIndex("申请科室")) = 750: .ColAlignment(.ColIndex("申请科室")) = flexAlignCenterCenter

138               .ColKey(26) = "打印人": .ColWidth(.ColIndex("打印人")) = 750: .ColAlignment(.ColIndex("打印人")) = flexAlignCenterCenter
139               .ColKey(27) = "打印时间": .ColWidth(.ColIndex("打印时间")) = 2000: .ColAlignment(.ColIndex("打印时间")) = flexAlignCenterCenter
140               .ColKey(28) = "结果次数": .ColWidth(.ColIndex("结果次数")) = 2000: .ColAlignment(.ColIndex("结果次数")) = flexAlignCenterCenter: .ColHidden(.ColIndex("结果次数")) = True
141               .ColKey(29) = "传染病": .ColWidth(.ColIndex("传染病")) = 750: .ColAlignment(.ColIndex("传染病")) = flexAlignCenterCenter: .ColHidden(.ColIndex("传染病")) = True
142               .ColKey(30) = "复核人": .ColWidth(.ColIndex("复核人")) = 750: .ColAlignment(.ColIndex("复核人")) = flexAlignCenterCenter: .ColHidden(.ColIndex("复核人")) = Not InStr(mstrPrivs, "查看传染病报告") > 0
143               .ColKey(31) = "复核时间": .ColWidth(.ColIndex("复核时间")) = 750: .ColAlignment(.ColIndex("复核时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("复核时间")) = Not InStr(mstrPrivs, "查看传染病报告") > 0
144               .ColKey(32) = "采样人": .ColWidth(.ColIndex("采样人")) = 750: .ColAlignment(.ColIndex("采样人")) = flexAlignCenterCenter: .ColHidden(.ColIndex("采样人")) = True
145               .ColKey(33) = "采样时间": .ColWidth(.ColIndex("采样时间")) = 750: .ColAlignment(.ColIndex("采样时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("采样时间")) = True
                  .ColKey(34) = "病历号": .ColWidth(.ColIndex("病历号")) = 750: .ColAlignment(.ColIndex("病历号")) = flexAlignCenterCenter: .ColHidden(.ColIndex("病历号")) = True

146           Else
147               If InStr(strTitle, "采样人") <= 0 Then
148                   strTitle = strTitle & ";采样人,750,1;采样时间,750,1"
149               End If
150               var_tmp = Split(strTitle, ";")
151               .Rows = 1
152               .FixedRows = 1
153               .Cols = UBound(var_tmp) + 1
154               For lngLoop = LBound(var_tmp) To UBound(var_tmp)
155                   var_SubTmp = Split(var_tmp(lngLoop), ",")
156                   .ColKey(lngLoop) = var_SubTmp(0): .ColWidth(.ColIndex(var_SubTmp(0))) = var_SubTmp(1): .ColAlignment(.ColIndex(var_SubTmp(0))) = flexAlignCenterCenter: .ColHidden(.ColIndex(var_SubTmp(0))) = Not (Val(var_SubTmp(2)) = 1)
157                   If var_SubTmp(0) = "复核人" Or var_SubTmp(0) = "复核时间" Then
158                       .ColHidden(.ColIndex(var_SubTmp(0))) = Not InStr(mstrPrivs, "查看传染病报告") > 0
159                   End If
160                   .ColDataType(.ColIndex("选择")) = flexDTNull
161               Next
162               .ColDataType(.ColIndex("选择")) = flexDTBoolean
163           End If
164           .Cell(flexcpPicture, 0, .ColIndex("序号")) = mObjIco.Obj序号
165           .TextMatrix(0, .ColIndex("选择")) = ""
166           .Cell(flexcpPicture, 0, .ColIndex("打印")) = mObjIco.Obj打印
167           .TextMatrix(0, .ColIndex("报告")) = "报告"
168           .TextMatrix(0, .ColIndex("姓名")) = "姓名"
169           .TextMatrix(0, .ColIndex("性别")) = "性别"
170           .TextMatrix(0, .ColIndex("年龄")) = "年龄"
171           .TextMatrix(0, .ColIndex("申请项目")) = "申请项目"
172           .TextMatrix(0, .ColIndex("样本条码")) = "样本条码"
173           .TextMatrix(0, .ColIndex("审核时间")) = "审核时间"
174           .TextMatrix(0, .ColIndex("住院号")) = "住院号"
175           .TextMatrix(0, .ColIndex("床号")) = "床号"
176           .TextMatrix(0, .ColIndex("检验人")) = "检验人"
177           .TextMatrix(0, .ColIndex("审核人")) = "审核人"
178           .TextMatrix(0, .ColIndex("病人来源")) = "来源"
179           .TextMatrix(0, .ColIndex("申请人")) = "开单人"
180           .TextMatrix(0, .ColIndex("申请科室")) = "申请科室"

181           .TextMatrix(0, .ColIndex("打印人")) = "打印人"
182           .TextMatrix(0, .ColIndex("打印时间")) = "打印时间"
183           .TextMatrix(0, .ColIndex("版本")) = "版本"
184           .TextMatrix(0, .ColIndex("结果次数")) = "结果次数"
185           .TextMatrix(0, .ColIndex("复核人")) = "复核人"
186           .TextMatrix(0, .ColIndex("复核时间")) = "复核时间"
187           .TextMatrix(0, .ColIndex("采样人")) = "采样人"
188           .TextMatrix(0, .ColIndex("采样时间")) = "采样时间"
              .TextMatrix(0, .ColIndex("病历号")) = "病历号"



189           .Row = 0: .Col = .ColIndex("选择"): .CellPicture = mObjIco.Obj选择
190           .ExplorerBar = flexExSortShow

191           If blnReadData Then
192               Do Until rsTmp.EOF
193                   If lngKey <> Val(rsTmp("id") & "") Then
194                       .Rows = .Rows + 1

195                       .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""
196                       If rsTmp("打印") & "" = 0 And rsTmp("报告") = "已出" Then
197                           .TextMatrix(.Rows - 1, .ColIndex("选择")) = 1
198                       Else
199                           .TextMatrix(.Rows - 1, .ColIndex("选择")) = 0
200                       End If
201                       If Val(rsTmp("打印") & "") <> 0 Then
                              '不为0,这已经打印,显示打印图标
202                           .Cell(flexcpPicture, .Rows - 1, .ColIndex("打印")) = mObjIco.Obj打印
203                       End If
204                       .Cell(flexcpPicture, .Rows - 1, .ColIndex("版本")) = mObjIco.Obj新版

205                       .TextMatrix(.Rows - 1, .ColIndex("报告")) = rsTmp("报告") & ""
206                       If mintIn = 1 Then
207                           If txtName = "" Then
208                               txtName = rsTmp("姓名") & ""
209                           End If
210                       End If

211                       .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsTmp("姓名") & ""
212                       .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsTmp("性别") & ""
213                       .TextMatrix(.Rows - 1, .ColIndex("年龄")) = rsTmp("年龄") & ""
214                       .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsTmp("申请项目") & ""
215                       .TextMatrix(.Rows - 1, .ColIndex("样本条码")) = rsTmp("样本条码") & ""
216                       .TextMatrix(.Rows - 1, .ColIndex("审核时间")) = rsTmp("审核时间") & ""
217                       .TextMatrix(.Rows - 1, .ColIndex("住院号")) = rsTmp("住院号") & ""
218                       .TextMatrix(.Rows - 1, .ColIndex("床号")) = rsTmp("床号") & ""
219                       .TextMatrix(.Rows - 1, .ColIndex("申请时间")) = rsTmp("申请时间") & ""
220                       .TextMatrix(.Rows - 1, .ColIndex("病人ID")) = rsTmp("病人ID") & ""
221                       .TextMatrix(.Rows - 1, .ColIndex("核收时间")) = Format(rsTmp("核收时间") & "", "yyyy-mm-dd HH:mm:ss")
222                       .TextMatrix(.Rows - 1, .ColIndex("备注")) = rsTmp("备注") & ""
223                       .TextMatrix(.Rows - 1, .ColIndex("诊断")) = rsTmp("诊断") & ""
224                       .TextMatrix(.Rows - 1, .ColIndex("微生物")) = rsTmp("微生物") & ""
225                       .TextMatrix(.Rows - 1, .ColIndex("阳性报告")) = rsTmp("阳性报告") & ""
226                       .TextMatrix(.Rows - 1, .ColIndex("检验人")) = rsTmp("检验人") & ""
227                       .TextMatrix(.Rows - 1, .ColIndex("审核人")) = rsTmp("审核人") & ""
228                       .TextMatrix(.Rows - 1, .ColIndex("申请人")) = rsTmp("申请人") & ""
229                       .TextMatrix(.Rows - 1, .ColIndex("申请科室")) = rsTmp("申请科室") & ""

230                       .TextMatrix(.Rows - 1, .ColIndex("打印人")) = rsTmp("打印人") & ""
231                       .TextMatrix(.Rows - 1, .ColIndex("打印时间")) = rsTmp("打印时间") & ""
232                       .TextMatrix(.Rows - 1, .ColIndex("版本")) = rsTmp("版本") & ""
233                       .TextMatrix(.Rows - 1, .ColIndex("结果次数")) = rsTmp("结果次数") & ""
234                       .TextMatrix(.Rows - 1, .ColIndex("传染病")) = rsTmp("是否传染病") & ""
235                       .TextMatrix(.Rows - 1, .ColIndex("复核人")) = rsTmp("复核人") & ""
236                       .TextMatrix(.Rows - 1, .ColIndex("复核时间")) = rsTmp("复核时间") & ""
237                       .TextMatrix(.Rows - 1, .ColIndex("采样人")) = rsTmp("采样人") & ""
238                       .TextMatrix(.Rows - 1, .ColIndex("采样时间")) = rsTmp("采样时间") & ""
                          .TextMatrix(.Rows - 1, .ColIndex("病历号")) = rsTmp("病历号") & ""



239                       .TextMatrix(.Rows - 1, .ColIndex("病人来源")) = chkSource(rsTmp("病人来源") & "").Caption
240                       If mlngGetPatientID > 0 Then
241                           txtPatiNo = rsTmp("住院号") & ""
242                       End If
243                   Else
244                       .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = .TextMatrix(.Rows - 1, .ColIndex("申请项目")) & "," & rsTmp("申请项目") & ""
245                   End If
246                   lngKey = Val(rsTmp("id") & "")
247                   rsTmp.MoveNext
248               Loop


249               Set rsOldLisData = GetOldLisData(lngID, lngPatiID)
250               Do Until rsOldLisData.EOF
251                   If lngKey <> Val(rsOldLisData("id") & "") Then
252                       .Rows = .Rows + 1
253                       .TextMatrix(.Rows - 1, .ColIndex("id")) = rsOldLisData("id") & ""
254                       If rsOldLisData("打印") & "" = "" And rsOldLisData("报告") & "" = "已出" Then
255                           .TextMatrix(.Rows - 1, .ColIndex("选择")) = 1
256                       Else
257                           .TextMatrix(.Rows - 1, .ColIndex("选择")) = 0
258                       End If
259                       If Val(rsOldLisData("打印") & "") <> 0 Then
                              '不为0,这已经打印,显示打印图标
260                           .Cell(flexcpPicture, .Rows - 1, .ColIndex("打印")) = mObjIco.Obj打印
261                       End If

262                       .Cell(flexcpPicture, .Rows - 1, .ColIndex("版本")) = mObjIco.Obj老版
263                       .TextMatrix(.Rows - 1, .ColIndex("报告")) = rsOldLisData("报告") & ""
264                       .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsOldLisData("姓名") & ""
265                       .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsOldLisData("性别") & ""
266                       .TextMatrix(.Rows - 1, .ColIndex("年龄")) = rsOldLisData("年龄") & ""
267                       .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsOldLisData("申请项目") & ""
268                       .TextMatrix(.Rows - 1, .ColIndex("样本条码")) = rsOldLisData("样本条码") & ""
269                       .TextMatrix(.Rows - 1, .ColIndex("审核时间")) = rsOldLisData("审核时间") & ""
270                       .TextMatrix(.Rows - 1, .ColIndex("住院号")) = rsOldLisData("住院号") & ""
271                       .TextMatrix(.Rows - 1, .ColIndex("床号")) = rsOldLisData("床号") & ""
272                       .TextMatrix(.Rows - 1, .ColIndex("申请时间")) = rsOldLisData("申请时间") & ""
273                       .TextMatrix(.Rows - 1, .ColIndex("病人ID")) = rsOldLisData("病人ID") & ""
274                       .TextMatrix(.Rows - 1, .ColIndex("核收时间")) = Format(rsOldLisData("核收时间") & "", "yyyy-mm-dd HH:mm:ss")
275                       .TextMatrix(.Rows - 1, .ColIndex("备注")) = rsOldLisData("备注") & ""
276                       .TextMatrix(.Rows - 1, .ColIndex("诊断")) = rsOldLisData("诊断") & ""
277                       .TextMatrix(.Rows - 1, .ColIndex("微生物")) = rsOldLisData("微生物") & ""
278                       .TextMatrix(.Rows - 1, .ColIndex("阳性报告")) = rsOldLisData("阳性报告") & ""
279                       .TextMatrix(.Rows - 1, .ColIndex("检验人")) = rsOldLisData("检验人") & ""
280                       .TextMatrix(.Rows - 1, .ColIndex("审核人")) = rsOldLisData("审核人") & ""
281                       .TextMatrix(.Rows - 1, .ColIndex("申请人")) = rsOldLisData("申请人") & ""
282                       .TextMatrix(.Rows - 1, .ColIndex("申请科室")) = rsOldLisData("申请科室") & ""

283                       .TextMatrix(.Rows - 1, .ColIndex("打印人")) = rsOldLisData("打印人") & ""
284                       .TextMatrix(.Rows - 1, .ColIndex("打印时间")) = rsOldLisData("打印时间") & ""
285                       .TextMatrix(.Rows - 1, .ColIndex("版本")) = rsOldLisData("版本") & ""
286                       .TextMatrix(.Rows - 1, .ColIndex("结果次数")) = rsOldLisData("结果次数") & ""
287                       .TextMatrix(.Rows - 1, .ColIndex("传染病")) = ""
288                       .TextMatrix(.Rows - 1, .ColIndex("复核人")) = ""
289                       .TextMatrix(.Rows - 1, .ColIndex("复核时间")) = ""
290                       .TextMatrix(.Rows - 1, .ColIndex("采样人")) = rsOldLisData("采样人") & ""
291                       .TextMatrix(.Rows - 1, .ColIndex("采样时间")) = rsOldLisData("采样时间") & ""
                          .TextMatrix(.Rows - 1, .ColIndex("病历号")) = rsOldLisData("病历号") & ""


292                       .TextMatrix(.Rows - 1, .ColIndex("病人来源")) = chkSource(rsOldLisData("病人来源") & "").Caption
293                       If mlngGetPatientID > 0 Then
294                           txtPatiNo = rsOldLisData("住院号") & ""
295                       End If
296                   Else
297                       .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = .TextMatrix(.Rows - 1, .ColIndex("申请项目")) & "," & rsOldLisData("申请项目") & ""
298                   End If
299                   lngKey = Val(rsOldLisData("id") & "")
300                   rsOldLisData.MoveNext
301               Loop


302               If rsTmp.RecordCount > 0 Or rsOldLisData.RecordCount > 0 Then
303                   .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
304                   .Row = 1
305               Else
306                   .Rows = 2
307                   .Row = 1
308               End If
309               .Cell(flexcpSort, .FixedRows, .ColIndex("核收时间"), .Rows - 1, .ColIndex("核收时间")) = 1
310           End If


311           If mlngSelRow <> 0 And Not mlngSelRow > .Rows - 1 Then
312               .Select mlngSelRow, .ColIndex("姓名")
313               .ShowCell mlngSelRow, .ColIndex("姓名")
314           End If

              '获取序号
315           For lngLoop = 1 To .Rows - 1
316               .TextMatrix(lngLoop, .ColIndex("序号")) = lngLoop
317               .Cell(flexcpBackColor, lngLoop, .ColIndex("序号")) = &HFFEBD7
318           Next

319       End With
320       If lngID <> -1 Then
321           If mintIn = 0 Then
322               mlngGetPatientID = 0
323               mlngPatientID = 0
324           End If
325       End If


326       Exit Sub
ReadPatientList_Error:
327       Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(ReadPatientList)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)

328       Err.Clear

End Sub

Private Function GetOldLisData(lngID As Long, Optional lngGetPatiID As Long) As ADODB.Recordset
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim strDepts As String
          Dim strDept As String
          Dim lngPatiID As Long
          Dim strTemp  As String
          Dim strWhere As String
1         lngPatiID = lngGetPatiID
      '    strSQL = "      Select a.Id, 0 选择,  Decode(a.审核人, Null, '未出', '已出') 报告, a.姓名, a.性别, a.年龄, c.医嘱内容 申请项目, a.住院号, a.门诊号, a.床号, a.申请时间, a.病人id, a.核收时间, a.审核时间, a.备注," & vbNewLine & _
      '           "    b.项目 || ':' || b.内容 诊断, a.微生物标本 微生物, 1 阳性报告, 0 查阅, a.医嘱id 申请id, a.打印次数 打印, a.申请人, a.检验人, a.审核人, Nvl(a.病人来源, 0) 病人来源, '' 打印人, '' 打印时间, 10 版本, a.报告结果 结果次数" & vbNewLine & _
      '           "     From 检验标本记录 A, 病人医嘱附件 B, 病人医嘱记录 C,部门表 D" & vbNewLine & _
      '           "     Where a.医嘱id = b.医嘱id(+) And a.医嘱id = c.Id(+) and a.申请科室id = d.id  And a.核收时间 Between [1] And [2]   "
2         On Error GoTo GetOldLisData_Error

3         strSQL = "Select Distinct a.Id, 0 选择, Decode(a.审核人, Null, '未出', '已出') 报告, a.姓名, a.性别, a.年龄, c.医嘱内容 申请项目, a.住院号, a.门诊号, a.床号," & vbNewLine & _
                  "                a.申请时间, a.病人id, a.核收时间, a.审核时间, a.检验备注 备注, f.诊断, a.微生物标本 微生物, 1 阳性报告, 0 查阅, a.医嘱id 申请id, a.打印次数 打印, a.申请人,a.病人科室 申请科室," & vbNewLine & _
                  "                a.检验人, a.审核人, Nvl(a.病人来源, 0) 病人来源, '' 打印人, '' 打印时间, 10 版本, a.报告结果 结果次数,a.样本条码,a.采样人,a.采样时间,a.标识号 病历号" & vbNewLine & _
                  "From 检验标本记录 A, 病人医嘱记录 C, 部门表 D," & vbNewLine & _
                  "     (Select b.医嘱id 医嘱id, f_List2str(Cast(Collect(b.项目 || ':' || b.内容) As t_Strlist)) 诊断" & vbNewLine & _
                  "       From 检验标本记录 A, 病人医嘱附件 B" & vbNewLine & _
                  "       Where a.医嘱id = b.医嘱id and a.核收时间 between [1] and [2] [条件]" & vbNewLine & _
                  "       Group By b.医嘱id) F" & vbNewLine & _
                  "Where a.医嘱id = c.Id(+) And a.医嘱id = f.医嘱id(+) And a.申请科室id = d.Id and a.核收时间 between [1] and [2] "

4         Select Case cboPrint.ListIndex

              Case 1
                  '已打印
5                 strWhere = strWhere & "and  a.打印次数 is not null  "
6                 strSQL = strSQL & "and  a.打印次数 is not null  "

7             Case 2
                  '未打印
8                 strWhere = strWhere & " and a.打印次数 is null  "
9                 strSQL = strSQL & "and  a.打印次数 is  null  "

10        End Select

11        If chkAudit(0).value = 1 And chkAudit(1).value = 0 Then
12            strSQL = strSQL & "And a.审核人 Is Not Null And a.审核时间 Is Not Null"
13        ElseIf chkAudit(0).value = 0 And chkAudit(1).value = 1 Then
14            strSQL = strSQL & "And a.审核人 Is  Null And a.审核时间 Is  Null"
15        ElseIf chkAudit(0).value = 1 And chkAudit(1).value = 1 Then

16        End If


          '传染病
17        Select Case Me.cboDiseases.Text
              Case "传染病"
18                strSQL = strSQL & " and a.id = -1"
19        End Select

          '处理病人来源
20        strTemp = checkboxSource()
21        If strTemp <> "" Then
22            strWhere = strWhere & " and nvl(a.病人来源,0) in (" & strTemp & ")"
23            strSQL = strSQL & " and nvl(a.病人来源,0) in (" & strTemp & ")"

24        Else
              '为选择病人来源时病人来源条件为-1
25            strWhere = strWhere & " and nvl(a.病人来源,0) in (-1)"
26            strSQL = strSQL & " and nvl(a.病人来源,0) in (-1)"

27        End If

28        If lngID = -1 Then
29            strSQL = strSQL & " and a.id = -1 "
30        End If

31        If txtName <> "" Then
32            strWhere = strWhere & " and a.姓名 like '" & txtName & "%' "
33            strSQL = strSQL & " and a.姓名 like '" & txtName & "%' "

34        End If

35        If txtDoctor.Text <> "" Then
36            strWhere = strWhere & " and a.申请人 like '" & txtDoctor.Text & "%'"
37            strSQL = strSQL & " and a.申请人 like '" & txtDoctor.Text & "%'"

38        End If

39        If Trim(txtPatiNo <> "") Then
40            If lblNo.Caption = "住院号↓" Then
41                strWhere = strWhere & " and a.住院号 = [3] "
42                strSQL = strSQL & " and a.住院号 = [3] "

43            ElseIf lblNo.Caption = "门诊号↓" Then
44                strWhere = strWhere & " and a.门诊号 = [3] "
45                strSQL = strSQL & " and a.门诊号 = [3] "

46            ElseIf lblNo.Caption = "床号↓" Then
47                strWhere = strWhere & " and a.床号 = [3] "
48                strSQL = strSQL & " and a.床号 = [3] "
49            ElseIf lblNo.Caption = "就诊卡↓" Then
50                strWhere = strWhere & " and a.病人ID = [7] "
51                strSQL = strSQL & " and a.病人ID = [7] "
52            ElseIf lblNo.Caption = "条码号↓" Then
53                strWhere = strWhere & " and a.样本条码 = [3] "
54                strSQL = strSQL & " and a.样本条码 = [3] "
55            End If
56        End If
57        If cboDept <> "" Then
58            strDept = Mid(cboDept.Text, InStr(cboDept.Text, "-") + 1)
59            If InStr(strDept, "所有科室") > 0 Then

60            Else
61                strSQL = strSQL & " and d.名称 =[6] "
62            End If
63        End If
64        If lngPatiID = 0 And mlngPatientID <> 0 Then
65            lngPatiID = mlngPatientID
66            strWhere = strWhere & " and a.病人ID = [7] "
67            strSQL = strSQL & " and a.病人ID = [7] "
68        Else
69            If lngPatiID <> 0 Then
70                strWhere = strWhere & " and a.病人ID = [7] "
71                strSQL = strSQL & " and a.病人ID = [7] "
72            End If
73        End If
74        strSQL = Replace(strSQL, "[条件]", strWhere)
75        strSQL = strSQL & " order by a.审核时间,a.病人id,a.id"
76        If mintIn = 1 Then
77            strSQL = Replace(strSQL, "and a.核收时间 between [1] and [2]", "")
78        End If
79        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入病人列表", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
                                  CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59"), txtPatiNo, mlngGetPatientID, strDepts, strDept, lngPatiID)
80        Set GetOldLisData = rsTmp


81        Exit Function
GetOldLisData_Error:
82        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(GetOldLisData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
83        Err.Clear

End Function



'----------三方微生物报告处理
Private Sub beginPrint()
    Dim strFileSource As String
    Dim lng报告ID As String
    strFileSource = GetLisRptFile(mstrTag)
    lng报告ID = Split(mstrTag, ";")(0)
    Call FunFastPrint(strFileSource, lng报告ID)

End Sub

Private Sub picRpt_Resize()
    On Error Resume Next
    webSub.Move 0, 0, picRpt.Width, picRpt.Height
End Sub


Private Function GetLisRptFile(ByVal strTag As String) As String
'功能：打开LIS报告文件查看，获取临时文件路径
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    Dim lng报告ID As String
    Dim str报告名 As String
    Dim lng类型 As String
    Dim varTmp As Variant
    Dim strSuffix As String '文件后缀名

    Screen.MousePointer = 11

    varTmp = Split(strTag, ";")
    lng报告ID = varTmp(0)
    strTmp = Replace(strTag, varTmp(0) & ";" & varTmp(1) & ";", "")
    varTmp = Split(strTmp, "<sTab>")
    lng类型 = varTmp(0)
    If lng类型 = 0 Then
        strSuffix = "pdf"
    ElseIf lng类型 = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    str报告名 = varTmp(1)

    strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng报告ID & "." & strSuffix
    If Not objFile.FileExists(strFile) Then
        strFile = ReadLob(100, 22, lng报告ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "文件内容读取失败！", vbInformation, "中联信息":
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function


Private Sub FunFastPrint(ByVal strFile As String, ByVal lngRptID As Long)
'功能：API调用快速打印PDF文件
'参数：strFile 文件路径
    Dim RetVal As Long
    Dim strSQL As String
    Dim ShExInfo As SHELLEXECUTEINFO

     On Error GoTo errH
    With ShExInfo
        .cbSize = Len(ShExInfo)
        .fMask = &H40
        .hWnd = 0
        .lpVerb = "print"
        .lpFile = strFile
        .lpParameters = ""
        .lpDirectory = vbNullChar
        .nShow = 2
    End With
    RetVal = ShellExecuteEx(ShExInfo)
    If RetVal = 0 Then
        Exit Sub
    End If
'    strSQL = "Zl_医嘱报告内容_Print(" & lngRptID & ",0)"
'    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
   Exit Sub
errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

Private Sub WebShow(ByVal strKey As String)
'功能：Web控件展示文件
    Dim strURL As String
    If strKey = "" Then
        Call webSub.Navigate("about:blank")
        webSub.Visible = False
'        mstrCurFile = ""
    Else
        strURL = GetLisRptFile(strKey)
        If strURL <> "" Then
            webSub.Navigate strURL
'            mstrCurFile = strURL
        End If
        webSub.Visible = True
    End If
End Sub

'-------------end---------

Private Sub tabPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    RefreshTab Item.Index
End Sub

Private Sub txtDoctor_GotFocus()
    Call selAllText(txtDoctor)
End Sub

Private Sub txtDoctor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtDoctor <> "" Then
            Call ReadPatientList(1)
            Call selAllText(txtDoctor)
        End If
    End If
End Sub

Private Sub txtName_GotFocus()
    Call selAllText(txtName)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtName <> "" Then
            If mintIn = 0 Then
                mlngPatientID = 0
            End If
            Call ReadPatientList(1)
            Call selAllText(txtName)
        End If
    End If
End Sub

Private Sub txtPatiNo_GotFocus()
    txtPatiNo.SelStart = 0
    txtPatiNo.SelLength = Len(txtPatiNo)
End Sub

Private Sub txtPatiNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPatiNo = ReplaseSpecial(txtPatiNo)
        If Trim(txtPatiNo) <> "" Then
            If mintIn = 0 Then
                mlngPatientID = 0
            End If
            Call ReadPatientList(1)
            Call selAllText(txtPatiNo)
        End If
    End If
End Sub

Private Sub VSFContrast_SelChange()
    Dim strErr As String
    Dim intType As Integer
    With Me.VSFContrast
        If .Row > 0 Then
            If Me.optContrast(0).value = True Then
                intType = 1
            Else
                intType = 2
            End If
            Call LoadVSFContrastToCht(Me.VSFContrast, Me.chtContrast, .Row, intType, strErr)
        End If
    End With
End Sub

Private Sub ReadSampleVal(lngSampleID As Long, Optional intVal As Integer = 25)
      '功能   读入结果信息
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim lngGroupId As Long
          Dim lngGroupMer As Long
          Dim strGroup As String
          Dim lngGroupRow As Long
          Dim strTitle As String
          Dim lngNo As Long
          Dim i As Integer

1         On Error GoTo ReadSampleVal_Error

2         If intVal = 25 Then
3             If gobjLiscomlib.IsTre(lngSampleID) Then
4                 strSQL = "select * from ( " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')' || decode(g.耐受时间,null,'', '(' || g.耐受时间 || ')') 检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间,B.ID 排序 " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e,耐受试验标本 F,检验耐受时间方案 G" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and d.id =e.组合id and b.ID=F.报告明细id(+) and F.耐受方案id=G.id(+) AND b.组合id is not null and e.组合id is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')' || decode(g.耐受时间,null,'', '(' || g.耐受时间 || ')') 检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间,B.ID 排序  " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e,耐受试验标本 F,检验耐受时间方案 G" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and b.ID=F.报告明细id(+) and F.耐受方案id=G.id(+) AND e.组合id is null and b.组合id is null and a.id = [1] ) order by 排序 desc" & vbNewLine
5             Else
6                 strSQL = "select * from ( " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')' 检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间 " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and d.id =e.组合id and  b.组合id is not null and e.组合id is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')' 检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间 " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and e.组合id is null and b.组合id is null and a.id = [1] ) order by 组合id,排列序号" & vbNewLine

7             End If
8             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", lngSampleID)

9         Else
10            strSQL = "   Select /*+ rule */" & vbNewLine & _
                     "  Distinct '' 序号,a.标本id, a.诊疗项目id, a.编码, a.排列序号, a.固定项目, a.Id, a.检验项目, a.临床意义, a.缩写 As 英文名, a.Cv," & vbNewLine & _
                     " 结果标志 , Decode(a.本次结果, '-', '阴性（-）', '+', '阳性（+）', '*', '*.**', a.本次结果) As 结果, Rownum As 序号, a.标志, a.仪器id, a.标本类别," & vbNewLine & _
                     "   a.核收时间, a.标本序号, a.标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号, a.当前床号, a.主页id, a.结果范围, Nvl(g.小数位数, 2) As 小数," & vbNewLine & _
                     "    a.警戒上限, a.警戒下限, a.单位," & vbNewLine & _
                     "   Trim(Replace(Replace(' ' ||" & vbNewLine & _
                     "                         Zlgetreference(a.Id, a.标本类型, Decode(a.性别, '男', 1, '女', 2, 0), a.出生日期, a.仪器id, a.年龄), ' .'," & vbNewLine & _
                     "                         '0.'), '～.', '～0.')) As 参考, a.Od, a.Cutoff, a.Cov, a.酶标板id, a.变异报警, a.变异警示, a.结果类型," & vbNewLine & _
                     "   A.结果参考,a.诊疗项目" & vbNewLine & _
                     "  From (Select a.Id As 标本id, b.诊疗项目id, LPad(Decode(d.排列序号, Null, Nvl(h.编码, c.编码), d.排列序号), 4, '0') As 编码," & vbNewLine & _
                     "        Nvl(b.排列序号, 9999) As 排列序号, Decode(b.诊疗项目id, Null, 0, 1) As 固定项目, b.检验项目id As ID," & vbNewLine & _
                     "       c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目, d.临床意义, d.缩写, b.原始结果, '' As 上次结果, '' As 上次时间, '' As Cv," & vbNewLine & _
                     "       b.结果标志, b.检验结果 As 本次结果, d.计算公式, d.结果类型, Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
                     "        Nvl(a.仪器id, -1) As 仪器id, Nvl(a.标本类别, 0) As 标本类别, a.核收时间, a.标本序号," & vbNewLine & _
                     "        Decode(a.仪器id, Null," & vbNewLine & _
                     "                To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000'), a.标本序号) As 标本号显示," & vbNewLine & _
                     "        a.检验备注, a.姓名, a.性别, a.年龄, a.标本类型, a.出生日期, a.门诊号, a.住院号, a.床号 As 当前床号, a.主页id, d.结果范围, d.警戒上限, d.警戒下限, d.单位," & vbNewLine & _
                     "        b.Od, b.Cutoff, b.Sco As Cov, b.酶标板id, d.变异报警率 As 变异报警, d.变异警示率 As 变异警示, b.结果参考,h.名称 诊疗项目" & vbNewLine & _
                     " From 检验标本记录 A, 检验普通结果 B, 诊治所见项目 C, 检验项目 D, 诊疗项目目录 H" & vbNewLine & _
                     " Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And b.诊疗项目id = h.Id(+) And b.记录类型 = a.报告结果 And" & vbNewLine & _
                     "       A.ID = [1]" & vbNewLine & _
                     " Union All" & vbNewLine & _
                     " Select a.Id As 标本id, b.诊疗项目id, LPad(Decode(d.排列序号, Null, Nvl(h.编码, c.编码), d.排列序号), 4, '0') As 编码," & vbNewLine & _
                  "       Nvl(b.排列序号, 9999) As 排列序号, Decode(b.诊疗项目id, Null, 0, 1) As 固定项目, b.检验项目id As ID,"
11            strSQL = strSQL & "        c.中文名 || Decode(d.缩写, Null, '', '(' || d.缩写 || ')') As 检验项目, d.临床意义, d.缩写, b.原始结果, '' As 上次结果, '' As 上次时间, '' As Cv," & vbNewLine & _
                     "     b.结果标志,   b.检验结果 As 本次结果, d.计算公式, d.结果类型, Decode(b.结果标志, 3, '↑', 2, '↓', 1, '', 4, '异常', 5, '↓↓', 6, '↑↑', '') As 标志," & vbNewLine & _
                     "        Nvl(a.仪器id, -1) As 仪器id, Nvl(a.标本类别, 0) As 标本类别, a.核收时间, a.标本序号," & vbNewLine & _
                     "        Decode(a.仪器id, Null," & vbNewLine & _
                     "                To_Char(Trunc(a.标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(a.标本序号, 10000), '0000'), a.标本序号) As 标本号显示," & vbNewLine & _
                     "        a.检验备注, a.姓名, a.性别, a.年龄, a.标本类型, a.出生日期, a.门诊号, a.住院号, a.床号 As 当前床号, a.主页id, d.结果范围, d.警戒上限, d.警戒下限, d.单位," & vbNewLine & _
                     "        b.Od, b.Cutoff, b.Sco As Cov, b.酶标板id, d.变异报警率 As 变异报警, d.变异警示率 As 变异警示, b.结果参考,h.名称 诊疗项目" & vbNewLine & _
                     " From 检验标本记录 A, 检验标本记录 E, 检验普通结果 B, 诊治所见项目 C, 检验项目 D, 检验仪器项目 G, 诊疗项目目录 H" & vbNewLine & _
                     " Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And b.诊疗项目id = h.Id(+) And b.记录类型 = a.报告结果 And" & vbNewLine & _
                     "       e.Id = a.合并id And e.id = [1]) A, 检验仪器项目 G" & vbNewLine & _
                     "  Where a.仪器id = g.仪器id(+) And a.Id = g.项目id(+)" & vbNewLine & _
                     "  Order By a.诊疗项目ID, a.排列序号"
12            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验技师站", lngSampleID)

13        End If
14        With vsfCenter
15            .MergeCells = flexMergeRestrictColumns
16            For i = 0 To rsTmp.Fields.Count - 1
17                strTitle = strTitle & ";" & rsTmp.Fields(i).Name & ",0," & flexAlignLeftCenter
18            Next
19            If Mid(strTitle, 1, 1) = ";" Then strTitle = Mid(strTitle, 2)
20            Call vfgSetting(0, vsfCenter, strTitle)

21            .Rows = 1
22            Do While Not rsTmp.EOF
23                If intVal = 25 Then
24                    lngGroupId = Val(rsTmp("组合ID") & "")
25                    strGroup = rsTmp("组合名称") & ""
26                Else
27                    lngGroupId = Val(rsTmp("诊疗项目ID") & "")
28                    strGroup = rsTmp("诊疗项目") & ""
29                End If
30                If lngGroupMer <> lngGroupId Then
                      '                If lngGroupRow <> 0 Then .Cell(flexcpText, lngGroupRow, 0, lngGroupRow, .Cols - 1) = strGroup & "(共" & lngNo & "项)"
31                    lngNo = 0
32                    .Rows = .Rows + 1
33                    lngGroupRow = .Rows - 1
34                    .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, .Cols - 1) = strGroup
35                    .Cell(flexcpFontBold, .Rows - 1, 0, .Rows - 1, .Cols - 1) = True
36                    .MergeRow(.Rows - 1) = True
37                    .MergeCells = flexMergeRestrictRows
38                End If

39                lngNo = lngNo + 1
40                .Rows = .Rows + 1
41                For i = 0 To rsTmp.Fields.Count - 1
42                    .TextMatrix(.Rows - 1, .ColIndex(rsTmp.Fields(i).Name)) = rsTmp.Fields(i).value & ""
43                    .TextMatrix(.Rows - 1, .ColIndex("序号")) = lngNo
44                Next
45                lngGroupMer = lngGroupId
46                If lngGroupRow <> 0 Then .Cell(flexcpText, lngGroupRow, 0, lngGroupRow, .Cols - 1) = strGroup & "(共" & lngNo & "项)"
47                rsTmp.MoveNext
48            Loop

49            .ColWidth(.ColIndex("序号")) = 500: .ColHidden(.ColIndex("序号")) = False
50            .ColWidth(.ColIndex("检验项目")) = 1800: .ColHidden(.ColIndex("检验项目")) = False
51            .ColWidth(.ColIndex("结果")) = 800: .ColHidden(.ColIndex("结果")) = False
52            If intVal = 25 Then
53                .ColWidth(.ColIndex("上次")) = 800: .ColHidden(.ColIndex("上次")) = False
54            Else
55                .ColWidth(.ColIndex("标志")) = 800: .ColHidden(.ColIndex("标志")) = False
56            End If
57            .ColWidth(.ColIndex("单位")) = 900: .ColHidden(.ColIndex("单位")) = False
58            .ColWidth(.ColIndex("参考")) = 1000: .ColHidden(.ColIndex("参考")) = False

59            For i = 1 To .Rows - 1
60                If .Cell(flexcpFontBold, i, .ColIndex("序号")) = True Then
61                    .RowHidden(i) = IIf(Me.chkGroup.value = 1, False, True)
62                End If
63            Next
64        End With

65        CalcReferenceColour


66        Exit Sub
ReadSampleVal_Error:
67        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(ReadSampleVal)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
68        Err.Clear
End Sub

Private Sub optReportShow()
    If optReport(0).value = True Then
        txtNormalMicrobe.ForeColor = vbRed
        txtNormalMicrobe.FontBold = True
        txtNoFindMicrobe.ForeColor = vbRed
        txtNoFindMicrobe.FontBold = True
        txtNormalMicrobes.ForeColor = vbRed
        txtNormalMicrobes.FontBold = True
        txtMicroscope.ForeColor = vbRed
        txtMicroscope.FontBold = True
        txtMicroscopeFinded.ForeColor = vbRed
        txtMicroscopeFinded.FontBold = True
        txtMicroscopeNOFind.ForeColor = vbRed
        txtMicroscopeNOFind.FontBold = True
        txtGermComment.ForeColor = vbRed
        txtGermComment.FontBold = True
    Else
        txtNormalMicrobe.ForeColor = vbBlack
        txtNormalMicrobe.FontBold = False
        txtNoFindMicrobe.ForeColor = vbBlack
        txtNoFindMicrobe.FontBold = False
        txtNormalMicrobes.ForeColor = vbBlack
        txtNormalMicrobes.FontBold = False
        txtMicroscope.ForeColor = vbBlack
        txtMicroscope.FontBold = False
        txtMicroscopeFinded.ForeColor = vbBlack
        txtMicroscopeFinded.FontBold = False
        txtMicroscopeNOFind.ForeColor = vbBlack
        txtMicroscopeNOFind.FontBold = False
        txtGermComment.ForeColor = vbBlack
        txtGermComment.FontBold = False
    End If

End Sub

Private Sub ReadSampleBacteriology(lngSampleID As Long, Optional intVal As Integer = 25)
    '功能   读入结果信息
    Dim strErr As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ReadSampleBacteriology_Error

    If intVal = 25 Then
        strSQL = "select b.id,b.中文名 || '(' || b.英文名 || ')' 细菌名,a.培养描述 描述," & vbNewLine & _
                "       a.耐药机制,a.组合id," & vbNewLine & _
                "a.培养时间,a.正常菌,a.未检出,a.补充描述,a.无致病菌,a.无细菌,a.镜检设备,a.镜检检出," & _
                "a.镜检未检出,a.阳性评语,a.阴性评语,a.结果标志,a.细菌ID,a.是否镜检结果,a.结果性质 " & vbNewLine & _
                "from 检验报告细菌 a,检验细菌记录 b" & vbNewLine & _
                "where a.细菌id = b.id(+) and a.标本id = [1] "

        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", lngSampleID)
    Else
        strSQL = "SELECT Distinct B.编码, B.ID 细菌id ,D.报告结果,B.中文名 AS 细菌名, " & _
                    "A.检验结果 AS 检验结果,A.培养描述 as 描述, d.检验备注,d.备注 " & _
                    "FROM 检验普通结果 A,检验细菌 B,检验标本记录 D  " & _
                    "WHERE A.细菌id = B.ID And D.审核人 is Not null  " & _
                        "AND A.记录类型 = [1]  " & _
                        "AND D.ID=A.检验标本ID  " & _
                        "AND D.ID= [2] Order by B.编码"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验技师站", mlngValueC, lngSampleID)
    End If
'    rsTmp.Sort = "排列序号"
    If Not vfgLoadFromRecord(VsfMicrobe, rsTmp, strErr, imgVsf) Then Exit Sub

    With VsfMicrobe
       If intVal = 25 Then
            .ColWidth(.ColIndex("细菌名")) = 1800: .ColHidden(.ColIndex("细菌名")) = False
            .ColWidth(.ColIndex("描述")) = 1500: .ColHidden(.ColIndex("描述")) = False
            .ColWidth(.ColIndex("耐药机制")) = 1500: .ColHidden(.ColIndex("耐药机制")) = False
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst


                Me.txtNormalMicrobe = rsTmp("正常菌") & ""
                Me.txtNoFindMicrobe = rsTmp("未检出") & ""
                Me.txtNormalMicrobes = rsTmp("补充描述") & ""
                Me.chkPathopoiesiaGerm.value = IIf(rsTmp("无致病菌") = 1, 1, 0)
                Me.chkNoGerm.value = IIf(rsTmp("无细菌") = 1, 1, 0)
                Me.txtMicroscope = rsTmp("镜检设备") & ""
                Me.txtMicroscopeFinded = rsTmp("镜检检出") & ""
                Me.txtMicroscopeNOFind = rsTmp("镜检未检出") & ""
                Me.txtMicrobePositiveComment = rsTmp("阳性评语") & ""
                Me.txtGermComment = rsTmp("阴性评语") & ""
                If Val(rsTmp("是否镜检结果") & "") = 0 Then
                    chkMicroscope.value = 0
                Else
                    chkMicroscope.value = 1
                End If
                If Val(rsTmp("结果性质") & "") = 0 Then
                    optReport(1).value = True
                Else
                    optReport(0).value = True
                End If
                optReportShow

                ReadSampleAntibiotic mlngKey, Val(rsTmp("细菌ID") & "")
            End If
        Else
            .ColWidth(.ColIndex("细菌名")) = 1800: .ColHidden(.ColIndex("细菌名")) = False
            .ColWidth(.ColIndex("检验结果")) = 1500: .ColHidden(.ColIndex("检验结果")) = False
            .ColWidth(.ColIndex("描述")) = 1500: .ColHidden(.ColIndex("描述")) = False
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                txtMicrobePositiveComment = rsTmp("备注") & ""
                txtComment = rsTmp("检验备注") & ""
                ReadSampleAntibiotic mlngKey, Val(rsTmp("细菌ID") & ""), 10
            End If
        End If
    End With


    Exit Sub
ReadSampleBacteriology_Error:
    Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(ReadSampleBacteriology)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear

End Sub
Private Sub ReadSampleAntibiotic(lngSampleID As Long, lngBacteriologyID As Long, Optional intVal As Integer = 25)
          '功能           读入抗生素写入VSF
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strErr As String

1         On Error GoTo ReadSampleAntibiotic_Error

2         If intVal = 25 Then
3             strSQL = "select c.id,a.细菌id,c.中文名 || '(' || c.英文名 || ')' 抗生素名,b.结果," & vbNewLine & _
                      "b.结果类型,b.药敏方法,b.复查次数,b.参考描述,b.药敏组ID,a.细菌ID " & vbNewLine & _
                      "from 检验报告细菌 a,检验报告药敏 b,检验药敏 c" & vbNewLine & _
                      "where a.id = b.结果id and b.药敏id = c.id and a.标本ID = [1] and a.细菌id = [2]  "
4             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", lngSampleID, lngBacteriologyID)
5         Else
6             strSQL = "SELECT C.细菌ID AS Key,B.ID,B.中文名 AS 抗生素名, A.结果 AS 结果,  " & _
                  "DECODE(A.结果类型,'R','R-耐药','I','I-中介','S','S-敏感','') AS 结果类型, " & _
                  "DECODE(A.药敏方法,1,'1-MIC',2,'2-DISK',3,'3-K-B','') As 药敏方法  " & _
                   "FROM 检验药敏结果 A, 检验用抗生素 B,检验普通结果 C  " & _
                  "Where A.抗生素ID = B.ID And C.ID=A.细菌结果ID AND C.记录类型=A.记录类型 AND C.检验标本id= [1] AND C.记录类型= [2] And C.细菌ID=[3] Order By B.编码"
7            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验技师站", lngSampleID, mlngValueC, lngBacteriologyID)
8         End If
      '    rsTmp.Sort = "排列序号"
9         If Not vfgLoadFromRecord(VsfAntibiotic, rsTmp, strErr, imgVsf) Then Exit Sub

10        With VsfAntibiotic
11            .ColWidth(.ColIndex("抗生素名")) = 1800: .ColHidden(.ColIndex("抗生素名")) = False
12            .ColWidth(.ColIndex("结果")) = 1500: .ColHidden(.ColIndex("结果")) = False
13            .ColWidth(.ColIndex("结果类型")) = 1500: .ColHidden(.ColIndex("结果类型")) = False
14        End With


15        Exit Sub
ReadSampleAntibiotic_Error:
16        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(ReadSampleAntibiotic)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
17        Err.Clear

End Sub
Private Function LoadContrastDBWriteVSF(VSFList As VSFlexGrid, lngSampleID As Long, lngPatientID As Long, SampleReportDate As Date, _
                                        intMaxDay As Integer, Optional strErr As String) As Boolean
      '功能                   从数据库中读出比对数据写入VSF中
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim lngItemid As Long
          Dim intCol As Integer
          Dim dblTmp As Double
          Dim blnTre As Boolean       '是否是耐受试验标本

1         On Error GoTo LoadContrastDBWriteVSF_Error



2         If mintVer = 25 Then
3             blnTre = gobjLiscomlib.IsTre(lngSampleID)

4             If blnTre Then
5                 strSQL = "Select b.id, b.中文名, b.英文名, b.单位, a.id 次数, c.报告时间, a.检验结果, e.耐受时间, b.变异报警率, b.结果类型, a.结果标志" & vbNewLine & _
                         "   From 检验报告明细 A, 检验指标 B, 检验报告记录 C, 耐受试验标本 D, 检验耐受时间方案 E" & vbNewLine & _
                         "   Where A.项目ID = B.ID And A.标本ID = C.ID And A.ID = D.报告明细id And D.耐受方案id = e.ID And A.标本ID = [1]" & vbNewLine & _
                         "   Order By a.id Desc"
6                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入比对数据", lngSampleID)
7             Else
8                 strSQL = "Select " & vbNewLine & _
                         " B.Id, B.中文名, B.英文名, B.单位, A.次数, A.报告时间, A.检验结果, B.变异报警率, B.结果类型" & vbNewLine & _
                           "From (Select B.项目id 检验项目id, B.次数, B.报告时间, B.检验结果" & vbNewLine & _
                         "       From (Select A.Id 次数, A.病人id, A.标本类型, A.报告时间, B.项目id" & vbNewLine & _
                         "              From 检验报告记录 A, 检验报告明细 B" & vbNewLine & _
                         "              Where A.Id = B.标本id And A.Id = [1] and b.检验结果 is not null ) A," & vbNewLine & _
                         "            (Select A.Id 次数, A.病人id, A.标本类型, A.报告时间, B.项目id, B.检验结果" & vbNewLine & _
                         "              From 检验报告记录 A, 检验报告明细 B" & vbNewLine & _
                         "              Where A.Id = B.标本id And A.病人id = [2] And 报告时间+0 Between [3] And [4] and a.id <= [1] and b.检验结果 is not null ) B" & vbNewLine & _
                         "       Where A.病人id = B.病人id And A.项目id + 0 = B.项目id And Nvl(A.标本类型, 0) = Nvl(B.标本类型, 0) ) A, 检验指标 B" & vbNewLine & _
                           "Where A.检验项目id = B.Id" & vbNewLine & _
                           "Order By LPad(B.排列序号, 10, '0'),b.id, A.次数 desc "

9                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入比对数据", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                         CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
10            End If
11        Else
12            strSQL = "    Select " & vbNewLine & _
                     "       i.Id, i.名称 As 中文名, v.缩写 As 英文名, i.计算单位 As 单位, a.次数, a.报告时间, a.检验结果, v.变异报警率, v.结果类型" & vbNewLine & _
                     "       From (Select b.检验项目id, b.次数, b.报告时间, b.检验结果" & vbNewLine & _
                     "              From (Select a.Id 次数, a.病人id, a.标本类型, a.审核时间 报告时间, b.检验项目id, b.检验结果" & vbNewLine & _
                     "                     From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                     "                     Where a.Id = b.检验标本id And a.Id = [1] And a.病人id = [2] And b.检验结果 Is Not Null) A," & vbNewLine & _
                     "                   (Select a.Id 次数, a.病人id, a.标本类型, a.审核时间 报告时间, b.检验项目id, b.检验结果" & vbNewLine & _
                     "                     From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                     "                     Where a.Id = b.检验标本id And a.Id < [1] And a.病人id = [2]  And  a.检验时间+0 Between [3] And [4]  And b.检验结果 Is Not Null) B" & vbNewLine & _
                     "              Where a.病人id = b.病人id And a.检验项目id + 0 = b.检验项目id) A, 检验项目 V, 检验报告项目 R, 诊疗项目目录 I" & vbNewLine & _
                     "       Where A.检验项目id = v.诊治项目id And A.检验项目id = r.报告项目id And r.诊疗项目id = i.ID And i.组合项目 <> 1"
13            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入比对数据", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                     CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
14        End If
15        vfgSetting 0, VSFList
16        With VSFList

17            .Rows = 1
18            .Cols = 1
19            .FixedRows = 1
              '        .FixedCols = 1
20            .TextMatrix(0, 0) = "检验项目": .ColWidth(0) = 2500
21            Do Until rsTmp.EOF
22                If lngItemid <> rsTmp("ID") Then
23                    .Rows = .Rows + 1
24                    intCol = 0
25                    If .Cols - 1 < intCol Then
26                        .Cols = .Cols + 1
27                        .ColWidth(intCol) = 1500
28                    End If

29                    If intCol = 0 Then
                          '写入项目
30                        .TextMatrix(.Rows - 1, intCol) = rsTmp("中文名") & "(" & rsTmp("英文名") & ")"

31                    End If
32                    intCol = intCol + 1
33                    If .Cols - 1 < intCol Then
34                        .Cols = .Cols + 1
35                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
36                        If Not blnTre Then
37                            .TextMatrix(0, intCol) = "本次"
38                        Else
39                            .TextMatrix(0, intCol) = rsTmp("耐受时间") & ""
40                        End If
41                    End If
                      '写入内容
42                    .TextMatrix(.Rows - 1, intCol) = rsTmp("检验结果") & ""
43                Else
44                    intCol = intCol + 1
45                    If .Cols - 1 < intCol Then
46                        .Cols = .Cols + 1
47                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
48                        If blnTre Then
49                            .TextMatrix(0, intCol) = rsTmp("耐受时间") & ""
50                        Else
51                            .TextMatrix(0, intCol) = "上" & intCol - 1 & "次"
52                        End If
53                        dblTmp = Val(CalcVolatility(.TextMatrix(.Rows - 1, 1), .TextMatrix(.Rows - 1, intCol)))
54                        If dblTmp <> 0 And Val(rsTmp("变异报警率") & "") <> 0 Then
55                            If dblTmp > Val(rsTmp("变异报警率") & "") Then
56                                .Cell(flexcpBackColor, .Rows - 1, intCol) = RGB(248, 194, 169)
57                            End If
58                        End If
59                    End If
                      '写入内容
60                    .TextMatrix(.Rows - 1, intCol) = rsTmp("检验结果") & ""
61                End If
62                lngItemid = rsTmp("ID")
63                rsTmp.MoveNext
64            Loop
65        End With

66        LoadContrastDBWriteVSF = True


67        Exit Function
LoadContrastDBWriteVSF_Error:
68        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(LoadContrastDBWriteVSF)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
69        Err.Clear

End Function

Private Sub ReadHistorData()
    '功能           读出历次的数据
    Dim strErr As String
    Call LoadContrastDBWriteVSF(VSFContrast, mlngKey, mlngPatientID, mdteReportDate, 60, strErr)
End Sub
Private Function CalcVolatility(strCalcA As String, strCalcB As String) As String
    '计算变异率

    On Error Resume Next

    If strCalcA = "" Or strCalcB = "" Then
        CalcVolatility = ""
        Exit Function
    End If
    If Val(strCalcA) = 0 Or Val(strCalcB) = 0 Then
        CalcVolatility = ""
    End If

    '计算
    CalcVolatility = (Val(strCalcB) - Val(strCalcA)) / Val(strCalcA) * 100
End Function
Private Function LoadVSFContrastToCht(VSFList As VSFlexGrid, chtObj As Chart2D, intRow As Integer, intType As Integer, strErr As String) As Boolean
          '功能           从VSF读出数据写入Cht控件
          Dim intCol As Integer
          Dim dblMax As Double

1         On Error GoTo LoadVSFContrastToCht_Error

2         chtObj.ChartGroups(1).Data.NumSeries = 0
3         With chtObj.ChartGroups(1)
4             .ChartType = oc2dTypePlot  '折线
5             .Styles(oc2dTypePlot).Symbol.Shape = oc2dShapeBox
6             With .Data
7                 .Layout = oc2dDataArray
8                 .NumSeries = 1
9                 .NumPoints(1) = VSFList.Cols - 1
10            End With
11        End With

12        With chtObj.ChartArea
13            .Axes("X").MajorGrid.Spacing.IsDefault = True
14            .Axes("Y").MajorGrid.Spacing.IsDefault = True
15            .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '横坐标显示值提示

16        End With
17        With chtObj.ChartGroups(1).Data
18            For intCol = 1 To VSFList.Cols - 1
19                Select Case intType
                      Case 1
20                        If intCol = 1 Then
21                            If VSFList.TextMatrix(VSFList.Row, 1) <> "" Then
22                                .Y(1, intCol) = 0
23                            End If
24                        Else
25                            If IsNumeric(VSFList.TextMatrix(VSFList.Row, 1)) = True And IsNumeric(VSFList.TextMatrix(VSFList.Row, intCol)) = True Then
26                                If CalcVolatility(VSFList.TextMatrix(VSFList.Row, 1), VSFList.TextMatrix(VSFList.Row, intCol)) <> "" Then
27                                    .Y(1, intCol) = Val(CalcVolatility(VSFList.TextMatrix(VSFList.Row, 1), VSFList.TextMatrix(VSFList.Row, intCol)))
28                                Else
29                                    .Y(1, intCol) = 1E+308
30                                End If
31                            End If
32                        End If
33                    Case 2
34                        If IsNumeric(VSFList.TextMatrix(VSFList.Row, intCol)) = True Then
35                            .Y(1, intCol) = IIf(VSFList.TextMatrix(VSFList.Row, intCol) = "", 1E+308, VSFList.TextMatrix(VSFList.Row, intCol))
36                        End If
37                End Select
38                If Abs(.Y(1, intCol)) > Abs(dblMax) And .Y(1, intCol) <> 1E+308 Then
39                    dblMax = .Y(1, intCol)
40                End If
41            Next
42        End With

43        With chtObj.ChartArea
44            Select Case intType
                  Case 1              '变异率
45                    .Axes("Y").DataMax = Abs(dblMax)
46                    .Axes("Y").DataMin = Abs(dblMax) * -1
47                    .Axes("Y").Origin = 0
48                Case 2              '结果值
49                    .Axes("Y").DataMax = Abs(dblMax) + Abs(dblMax) / 100 * 10
50                    .Axes("Y").DataMin = 0
51                    .Axes("Y").Origin = 0
52            End Select
53        End With
54        LoadVSFContrastToCht = True


55        Exit Function
LoadVSFContrastToCht_Error:
56        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(LoadVSFContrastToCht)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
57        Err.Clear

End Function
Private Sub ReadContrastToVsf()
    '功能       读入历次比对到VSF
    Dim strErr As String

    Me.VSFContrast.Rows = 1: Me.VSFContrast.Rows = 2


    '没有病人ID时退出
    If mlngPatientID = 0 Then Exit Sub

    Call LoadContrastDBWriteVSF(Me.VSFContrast, mlngKey, mlngPatientID, mdteReportDate, Val(txtMaxDay), strErr)
    Call VSFContrast_SelChange
End Sub

Private Sub InitFace()
    '功能           初始化界面
    '========================================显示颜色设置============================================
    '显示颜色设置
    mSampleShowColour.正常 = &H80000005
    mSampleShowColour.偏高 = Val(ComGetPara(Sel_Lis_DB, "显示偏高颜色", 2500, 2500, "8438015"))
    mSampleShowColour.偏低 = Val(ComGetPara(Sel_Lis_DB, "显示偏低颜色", 2500, 2500, "8454143"))
    mSampleShowColour.警示偏高 = Val(ComGetPara(Sel_Lis_DB, "显示警示偏高颜色", 2500, 2500, "255"))
    mSampleShowColour.警示偏低 = Val(ComGetPara(Sel_Lis_DB, "显示警示偏低颜色", 2500, 2500, "255"))
    mSampleShowColour.复查偏高 = Val(ComGetPara(Sel_Lis_DB, "显示复查偏高颜色", 2500, 2500, "65280"))
    mSampleShowColour.复查偏低 = Val(ComGetPara(Sel_Lis_DB, "显示复查偏低颜色", 2500, 2500, "12648384"))
    mSampleShowColour.异常 = Val(ComGetPara(Sel_Lis_DB, "显示异常颜色", 2500, 2500, "16576"))

    '是否显示未出复选框
    chkAudit(0).Visible = InStr(";" & mstrPrivs & ";", ";查看未出报告;") > 0
    If chkAudit(0).Visible = False Then chkAudit(0).value = 1
    chkAudit(1).Visible = InStr(";" & mstrPrivs & ";", ";查看未出报告;") > 0
    If chkAudit(1).Visible = False Then chkAudit(1).value = 0
End Sub
Private Function GetValColour(intValType As Integer) As Double
    '功能               传入对应的结果类型1-正常、2-偏低、3-偏高、4-阳性(异常)、5-警示下限、6-警示上限、7-复查下限、8-复查上限
    '返回               对应的颜色
    Select Case intValType
        Case 1, 0
            GetValColour = mSampleShowColour.正常
        Case 2
            GetValColour = mSampleShowColour.偏低
        Case 3
            GetValColour = mSampleShowColour.偏高
        Case 4
            GetValColour = mSampleShowColour.异常
        Case 5
            GetValColour = mSampleShowColour.警示偏低
        Case 6
            GetValColour = mSampleShowColour.警示偏高
        Case 7
            GetValColour = mSampleShowColour.复查偏低
        Case 8
            GetValColour = mSampleShowColour.复查偏高
    End Select
End Function

Private Sub CalcReferenceColour()
          '功能           计算结果的颜色
          Dim intCol As Integer
          Dim intRow As Integer

1         On Error GoTo CalcReferenceColour_Error

2         With vsfCenter
3             For intRow = 1 To .Rows - 1
4                 If .ColIndex("id") <> -1 Then
5                     If Val(.TextMatrix(intRow, .ColIndex("id"))) <> 0 Then
      '                    If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("版本"))) = 25 Then
6                             .SelectionMode = flexSelectionFree
7                             If intRow = .Row Then
8                                 For intCol = 0 To .Cols - 1
9                                     .Cell(flexcpBackColor, intRow, intCol, intRow, intCol) = &HFFEBD7
10                                Next
11                            End If
      '                    Else
      '                        .SelectionMode = flexSelectionByRow
      '                        .BackColorSel = &HFFEBD7
      '                    End If
12                        .Cell(flexcpBackColor, intRow, .ColIndex("结果"), intRow, .ColIndex("结果")) = GetValColour(Val(.TextMatrix(intRow, .ColIndex("结果标志"))))
13                    End If
14                End If
15            Next
16        End With


17        Exit Sub
CalcReferenceColour_Error:
18        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(CalcReferenceColour)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
19        Err.Clear
End Sub

Private Sub vsfCenter_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
          Dim intCol As Integer

1         On Error GoTo vsfCenter_AfterRowColChange_Error

2         If OldRow <> NewRow Or OldCol <> NewCol Then
3             With vsfCenter
4                 If .ColIndex("id") <> -1 Then
5                     If Val(.TextMatrix(.Row, .ColIndex("id"))) <> 0 Then
6                         For intCol = 0 To .Cols - 1
7                             If intCol <> .ColIndex("结果") Then
8                                 If OldRow <> 0 Then .Cell(flexcpBackColor, OldRow, intCol, OldRow, intCol) = mSampleShowColour.正常
9                                 .Cell(flexcpBackColor, NewRow, intCol, NewRow, intCol) = &HFFEBD7
10                            End If
11                        Next
12                        If .Col = .ColIndex("结果") Then
13                            .BackColorSel = GetValColour(Val(.TextMatrix(NewRow, .ColIndex("结果标志"))))
14                        Else
15                            .BackColorSel = &HFFEBD7
16                        End If

17                        txtSignificance.Text = .TextMatrix(.Row, .ColIndex("临床意义"))
18                    End If
19                End If
20            End With
21        End If


22        Exit Sub
vsfCenter_AfterRowColChange_Error:
23        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(vsfCenter_AfterRowColChange)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
24        Err.Clear
End Sub

Private Sub vsfLeft_AfterSort(ByVal Col As Long, Order As Integer)
    vsfLeft_SelChange
End Sub

Private Sub vsfLeft_Click()
    With Me.vsfLeft
        If .Row < 1 Then Exit Sub
        Me.cbrthis.FindControl(, conFun_Sample_Auditing).Enabled = IIf(Val(.TextMatrix(.Row, .ColIndex("传染病"))) = 1, True, False) _
                                                                    And .TextMatrix(.Row, .ColIndex("复核人")) = "" And .TextMatrix(.Row, .ColIndex("复核时间")) = ""
        Me.cbrthis.FindControl(, conFun_Sample_unAuditing).Enabled = IIf(Val(.TextMatrix(.Row, .ColIndex("传染病"))) = 1, True, False) _
                                                                    And .TextMatrix(.Row, .ColIndex("复核人")) <> "" And .TextMatrix(.Row, .ColIndex("复核时间")) <> ""
        mlngSelRow = .Row
    End With
End Sub

Private Sub vsfLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim Point As POINTAPI
    Dim strTitle As String

    With Me.vsfLeft
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow < 0 Or lngCol < 0 Then Exit Sub
        If Button = 1 Then
            If lngCol = .ColIndex("选择") Then
                If .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 1 Then
                    .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 2
                Else
                    .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 1
                End If
            End If
        End If

        '点击鼠标右键,弹出设置窗体
        If Button = 2 Then
            If lngRow = 0 Then
                Call GetCursorPos(Point)
                strTitle = SetVsfColHiden(Me, Me.vsfLeft, Point.X * 15, Point.Y * 15, "报告中心显示列", 2500, 1013, "结果次数,id")
                If strTitle <> "" Then
                    SaveDBLog 18, 6, 0, "检验报告查询", "设置表格列的显示和排序:" & strTitle, 2500, "临床实验室管理"
                    Call ReadPatientList(1)
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfLeft_SelChange()
1         On Error GoTo vsfLeft_SelChange_Error

2         With vsfLeft
3             If .ColIndex("id") <> -1 And .ColIndex("病人ID") <> -1 And .ColIndex("核收时间") <> -1 Then
4                 If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 Then
5                     If Val(.TextMatrix(.Row, .ColIndex("ID"))) <> mlngKey Then
6                         Call clsAllEdit
                          mstrTag = ""
7                         mlngKey = Val(.TextMatrix(.Row, .ColIndex("ID")))
8                         mlngPatientID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
9                         mdteReportDate = .TextMatrix(.Row, .ColIndex("核收时间"))
10                        txtComment = .TextMatrix(.Row, .ColIndex("备注"))
11                        txtDiagnose = .TextMatrix(.Row, .ColIndex("诊断"))
12                        mlngValueC = .TextMatrix(.Row, .ColIndex("结果次数"))
13                        mintVer = .TextMatrix(.Row, .ColIndex("版本"))
14                        If Val(.TextMatrix(.Row, .ColIndex("微生物"))) = 1 Then



15                            If Val(.TextMatrix(.Row, .ColIndex("阳性报告"))) = 1 Then
16                                picGeneral.Visible = False
17                                picMicrobePositive.Visible = True
18                                PicNegative.Visible = False
                                  picRpt.Visible = False
19                                If Val(.TextMatrix(.Row, .ColIndex("版本"))) = 25 Then
20                                    ReadSampleBacteriology mlngKey, 25
21                                Else
22                                    ReadSampleBacteriology mlngKey, 10
23                                End If
24                            ElseIf Val(.TextMatrix(.Row, .ColIndex("阳性报告"))) = 3 Then
                                    picGeneral.Visible = False
                                    picMicrobePositive.Visible = False
                                    PicNegative.Visible = False
                                    picRpt.Visible = True
                                    findThirdReport (mlngKey)
                              Else
                                 picRpt.Visible = False
25                                picGeneral.Visible = False
26                                picMicrobePositive.Visible = False
27                                PicNegative.Visible = True
28                                ReadSampleBacteriology mlngKey
29                            End If
30

                        Else
31                            picGeneral.Visible = True
32                            picMicrobePositive.Visible = False
33                            PicNegative.Visible = False
                              picRpt.Visible = False
34                            If Val(.TextMatrix(.Row, .ColIndex("版本"))) = 25 Then
35                                ReadSampleVal mlngKey, 25
36                            Else
37                                ReadSampleVal mlngKey, 10
38                            End If
39                        End If

40                        RefreshTab Me.TabPage.Selected.Index
41                    End If
42                End If
43            End If
44        End With


45        Exit Sub
vsfLeft_SelChange_Error:
46        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(vsfLeft_SelChange)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
47        Err.Clear
End Sub

Private Sub findThirdReport(ByVal strAdvice As String)
    Dim strSQL As String
    Dim rsTemp As Recordset
    '三方LIS报告
    Dim strTag As String

    mstrTag = ""
    strSQL = "select b.id as 报告ID,b.报告名,b.报告名||','||To_Char(A.开嘱时间,'YYYY-MM-DD HH24:MI') as 文档标题,c.医嘱ID,b.类型,b.打印次数 from 病人医嘱记录 a, 医嘱报告内容 b,病人医嘱报告 c where b.id=c.报告id and a.id=c.医嘱id and c.报告id is not null and b.类型 in (0,2) and a.id =[1]"

    Set rsTemp = OpenSQLRecord(Sel_His_DB, strSQL, Me.Caption, strAdvice)
    If rsTemp.RecordCount > 0 Then
        strTag = rsTemp!报告ID & ";" & rsTemp!医嘱id & ";" & rsTemp!类型 & "<sTab>" & rsTemp!报告名
        mstrTag = strTag
        Call WebShow(strTag)
    End If


End Sub

Private Sub ImageTypeSet(intCount As Integer, Optional blnNoRead As Boolean = False)
    '功能           图像排版(最大9幅图)
    Dim intloop As Integer
    '先隐藏所有图像后才行按排列
    For intloop = 0 To 8
        If chtPic.Count - 1 < intloop Then
            Load chtPic(intloop)
        End If
        chtPic(intloop).Visible = False
        If blnNoRead = True Then
            chtPic(intloop).Reset
            chtPic(intloop).ChartGroups(1).Data.NumPoints(1) = 0
        End If
        chtPic(intloop).Interior.Image.Layout = oc2dImageStretched
        chtPic(intloop).Border.Type = oc2dBorderPlain
        chtPic(intloop).Border.Width = 1
        chtPic(intloop).IsBatched = False
    Next

    If intCount <= 4 Then
        '按4幅图进行排列
        chtPic(0).Left = 25
        chtPic(0).Top = 25
        chtPic(0).Width = (Me.PicPic.ScaleWidth - 50) / 2
        chtPic(0).Height = (Me.PicPic.ScaleHeight - 50) / 2

        chtPic(1).Left = chtPic(0).Left + chtPic(0).Width + 25
        chtPic(1).Top = 25
        chtPic(1).Width = Me.PicPic.ScaleWidth - chtPic(1).Left - 25
        chtPic(1).Height = chtPic(0).Height

        chtPic(2).Left = 25
        chtPic(2).Top = chtPic(0).Top + chtPic(0).Height + 25
        chtPic(2).Height = chtPic(0).Height
        chtPic(2).Width = chtPic(0).Width

        chtPic(3).Left = chtPic(1).Left
        chtPic(3).Top = chtPic(2).Top
        chtPic(3).Height = chtPic(2).Height
        chtPic(3).Width = Me.PicPic.ScaleWidth - chtPic(3).Left - 25
    ElseIf intCount <= 6 Then
        chtPic(0).Left = 25
        chtPic(0).Top = 25
        chtPic(0).Width = (Me.PicPic.ScaleWidth - 100) / 3
        chtPic(0).Height = chtPic(0).Width

        chtPic(1).Left = chtPic(0).Left + chtPic(0).Width + 25
        chtPic(1).Top = 25
        chtPic(1).Width = chtPic(0).Width
        chtPic(1).Height = chtPic(0).Height

        chtPic(2).Left = chtPic(1).Left + chtPic(1).Width + 25
        chtPic(2).Top = 25
        chtPic(2).Width = Me.PicPic.ScaleWidth - chtPic(2).Left
        chtPic(2).Height = chtPic(0).Height

        chtPic(3).Left = 25
        chtPic(3).Top = chtPic(0).Top + chtPic(0).Height + 25
        chtPic(3).Width = chtPic(0).Width
        chtPic(3).Height = Me.PicPic.ScaleHeight - chtPic(3).Left

        chtPic(4).Left = chtPic(3).Left + chtPic(3).Width + 25
        chtPic(4).Top = chtPic(3).Top
        chtPic(4).Width = chtPic(3).Width
        chtPic(4).Height = chtPic(3).Height

        chtPic(5).Left = chtPic(4).Left + chtPic(4).Width + 25
        chtPic(5).Top = chtPic(3).Top
        chtPic(5).Width = chtPic(3).Width
        chtPic(5).Height = chtPic(3).Height
    ElseIf intCount <= 9 Then
        chtPic(0).Left = 25
        chtPic(0).Top = 25
        chtPic(0).Width = (Me.PicPic.ScaleWidth - 100) / 3
        chtPic(0).Height = (Me.PicPic.ScaleHeight - 100) / 3

        chtPic(1).Left = chtPic(0).Left + chtPic(0).Width + 25
        chtPic(1).Top = 25
        chtPic(1).Width = chtPic(0).Width
        chtPic(1).Height = chtPic(0).Height

        chtPic(2).Left = chtPic(1).Left + chtPic(1).Width + 25
        chtPic(2).Top = 25
        chtPic(2).Width = Me.PicPic.ScaleWidth - chtPic(2).Left
        chtPic(2).Height = chtPic(0).Height

        chtPic(3).Left = 25
        chtPic(3).Top = chtPic(0).Top + chtPic(0).Height + 25
        chtPic(3).Width = chtPic(0).Width
        chtPic(3).Height = chtPic(0).Height

        chtPic(4).Left = chtPic(3).Left + chtPic(3).Width + 25
        chtPic(4).Top = chtPic(0).Top + chtPic(0).Height + 25
        chtPic(4).Width = chtPic(3).Width
        chtPic(4).Height = chtPic(3).Height

        chtPic(5).Left = chtPic(4).Left + chtPic(4).Width + 25
        chtPic(5).Top = chtPic(4).Top
        chtPic(5).Width = PicPic.ScaleWidth - chtPic(5).Left
        chtPic(5).Height = chtPic(3).Height

        chtPic(6).Left = 25
        chtPic(6).Top = chtPic(3).Top + chtPic(3).Height + 25
        chtPic(6).Width = chtPic(0).Width
        chtPic(6).Height = PicPic.ScaleHeight - chtPic(6).Top

        chtPic(7).Left = chtPic(6).Left + chtPic(6).Width + 25
        chtPic(7).Top = chtPic(6).Top
        chtPic(7).Width = chtPic(6).Width
        chtPic(7).Height = chtPic(6).Height

        chtPic(8).Left = chtPic(7).Left + chtPic(7).Width + 25
        chtPic(8).Top = chtPic(6).Top
        chtPic(8).Width = Me.PicPic.ScaleWidth - chtPic(8).Left
        chtPic(8).Height = chtPic(6).Height
    End If

    For intloop = 0 To 8
        chtPic(intloop).Visible = True
    Next
End Sub

Private Sub ReadImages(lngSampleID As Long, Optional intVal As Integer = 25)
          '功能               读入当前标本的图形到Cht
          Dim strChart(0 To 8) As String
          Dim strErr As String
          Dim intloop As Integer

          '先排版
1         On Error GoTo ReadImages_Error

2         Call ImageTypeSet(9, True)
          '读入图像数据
3         If ReadSampleImage(lngSampleID, strChart, strErr, intVal) = False Then
4             Exit Sub
5         End If
6         For intloop = 0 To 8
7             If strChart(intloop) <> "" Then
8                 chtPic(intloop).Load (strChart(intloop))
9             End If
10        Next
          '读入完成再排版
11        Call ImageTypeSet(9)


12        Exit Sub
ReadImages_Error:
13        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(ReadImages)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
14        Err.Clear

End Sub

Private Sub RefreshTab(Index As Integer)
    '功能           刷新对应的页
    Select Case Index
        Case 1
            ReadHistorData
        Case 2
            ReadImages mlngKey, mintVer
    End Select
End Sub
Public Function PrintReport(objFrm As Object, lngSampleID As Long, Optional byRunMode As Byte = 2, Optional strErr As String) As Boolean
          '功能       打印报告
          Dim intCount As Integer
          Dim strNO As String
          Dim intSel As Integer
          Dim strChart(0 To 8) As String
          Dim strSQL As String
          Dim strTmp As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset

1         On Error GoTo PrintReport_Error

2         strSQL = "select b.id 仪器id ,b.名称 仪器名称,b.仪器类别,a.病人来源,a.报告时间,a.阳性报告,a.标本序号 from 检验报告记录 a,检验仪器记录 b where a.仪器id = b.id and a.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告打印", lngSampleID)

4         If rsTmp.RecordCount = 0 Then Exit Function

5         strSQL = "select id,编码,名称,门诊单据,住院单据,体检单据,院外单据,门诊格式,住院格式,体检格式,院外格式,格式数量," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊单据, '00000')) || '-2' 门诊单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院单据, '00000')) || '-2' 住院单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检单据, '00000')) || '-2' 体检单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外单据, '00000')) || '-2' 院外单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊格式, '00000')) || '-2' 门诊格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院格式, '00000')) || '-2' 住院格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检格式, '00000')) || '-2' 体检格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外格式, '00000')) || '-2' 院外格式号" & vbNewLine & _
                      "from 检验仪器记录 where id = [1] "

6         Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", Val(rsTmp("仪器ID") & ""))


7         rsReportFormat.Filter = "id=" & Val(rsTmp("仪器ID") & "")
8         If Val(rsTmp("仪器类别")) = 1 Then
9             If Val(rsTmp("阳性报告") & "") = 1 Then
                  '阳性
10                intSel = 0
11            Else
                  '阴性
12                intSel = 1
13            End If
14        Else
15            intCount = GetSampleValCount(lngSampleID)
              '没有结果时提示
16            If intCount = 0 Then
17                Exit Function
18            End If
19            If rsReportFormat.RecordCount > 0 Then
20                If Val(rsReportFormat("格式数量") & "") > 0 Then
21                    If intCount > Val(rsReportFormat("格式数量") & "") Then
22                        intSel = 0
23                    Else
24                        intSel = 1
25                    End If
26                End If
27            Else
28                intSel = 0
29            End If

30        End If
31        Select Case Val(rsTmp("病人来源") & "")
              Case 1
32                If intSel = 0 Then
33                    strNO = rsReportFormat("门诊单据号")
34                Else
35                    strNO = rsReportFormat("门诊格式号")
36                End If
37            Case 2
38                If intSel = 0 Then
39                    strNO = rsReportFormat("住院单据号")
40                Else
41                    strNO = rsReportFormat("住院格式号")
42                End If
43            Case 3
44                If intSel = 0 Then
45                    strNO = rsReportFormat("住院单据号")
46                Else
47                    strNO = rsReportFormat("住院格式号")
48                End If
49            Case 4
50                If intSel = 0 Then
51                    strNO = rsReportFormat("院外单据号")
52                Else
53                    strNO = rsReportFormat("院外格式号")
54                End If
55            Case Else
56                If intSel = 0 Then
57                    strNO = rsReportFormat("门诊单据号")
58                Else
59                    strNO = rsReportFormat("门诊格式号")
60                End If
61        End Select
62        If byRunMode = 3 Then
63            If strNO <> "" Then
64                FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
65            End If
66        Else
             '读图像
67            strTmp = "开始读入图像:" & Now & vbCrLf
68            If ReadSampleImage(lngSampleID, strChart, strErr) = False Then
69                Exit Function
70            End If
71            strTmp = strTmp & "读入图像完成:" & Now & vbCrLf

72            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "标本ID=" & lngSampleID, "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), _
                      "图形4=" & strChart(3), "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                      "图形9=" & strChart(8), byRunMode
73            strTmp = strTmp & "打印完成:" & Now & vbCrLf

              '对于审核过的标本标识
74            strSQL = "Zl_检验报告打印_Edit(1," & lngSampleID & ")"
75            Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
76            strTmp = strTmp & "完成打印:" & Now

77            SaveDBLog 18, 6, lngSampleID, "打印", "报告打印", 2500, "临床实验室管理"
78        End If

79        PrintReport = True

          '发送刷新科内概况已打印标签申请
80        Call SendMessage("RefreshDeptSurvey7")


81        Exit Function
PrintReport_Error:
82        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(PrintReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
83        Err.Clear
End Function

Private Sub GetDept(Optional intType As Integer, Optional ByVal strInfo As String)
          '功能               读入科室或病区
          '参数               intType 0=科室 1=病区
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim blnFind As Boolean

1         On Error GoTo GetDept_Error

2         If strInfo <> "" Then blnFind = True
3         If intType = 0 Then
4             strSQL = "Select Distinct a.编码, a.名称, a.简码 From 部门表 A, 部门性质说明 B" & _
                      " Where a.Id = b.部门id And a.撤档时间 Is Not Null And a.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd') And" & _
                      " (b.工作性质 = '临床' Or b.工作性质 = '治疗' Or b.工作性质 = '护理' Or b.工作性质 = '检验')" & _
                      IIf(blnFind, "AND ( A.编码 like [1] or A.名称 like [2] or A.简码 like [2])", "") & "order by a.编码"

5         Else
6             strSQL = ""
7         End If
8         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入科室", IIf(IsNumeric(strInfo), strInfo, ""), UCase(strInfo))
9         If rsTmp.EOF Then Exit Sub
10        With cboDept
11            .Clear
12            If blnFind = False Then
13                If InStr(mstrPrivs, "所有科室") > 0 Then
14                    .AddItem "00-所有科室"
15                End If
16            End If
17            Do Until rsTmp.EOF
18                .AddItem Trim(rsTmp("编码")) & "-" & Trim(rsTmp("名称")) & ""
19                rsTmp.MoveNext
20            Loop
21            .ListIndex = 0
22        End With


23        Exit Sub
GetDept_Error:
24        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(GetDept)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
25        Err.Clear
End Sub


Private Sub CboFind(objcbo As ComboBox, lngID As Long)
    '功能           找到cbo对应的id
    Dim intloop As Integer
    With objcbo
        For intloop = 0 To .ListCount - 1
            If .ItemData(intloop) = lngID Then
                .ListIndex = intloop
                Exit Sub
            End If
        Next
        .ListIndex = 0
    End With
End Sub
Public Function ShowMe(objFrm As Object, Optional strErr As String) As Boolean

    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 1013)
    Me.Show , objFrm
    ShowMe = True

    '是否显示传染病筛选框
    picDiseases.Visible = InStr(mstrPrivs, "查看传染病报告") > 0
    Me.cboDiseases.ListIndex = 2

    With Me.cbrthis
        .FindControl(, conFun_Sample_Auditing).Visible = InStr(mstrPrivs, "查看传染病报告") > 0
        .FindControl(, conFun_Sample_unAuditing).Visible = InStr(mstrPrivs, "查看传染病报告") > 0
    End With
    GetDept 0
    Exit Function
errH:
    strErr = "出错函数(ShowMe),出错信息:" & Err.Number & " " & Err.Description
End Function

Public Function BHShowMe(lngMain As Long, Optional strErr As String) As Boolean
    On Error GoTo errH
    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 1013)


    gobjLiscomlib.ShowChildWindow Me.hWnd, lngMain
    BHShowMe = True

     '是否显示传染病筛选框
    picDiseases.Visible = InStr(mstrPrivs, "查看传染病报告") > 0
    Me.cboDiseases.ListIndex = 2

    With Me.cbrthis
        .FindControl(, conFun_Sample_Auditing).Visible = InStr(mstrPrivs, "查看传染病报告") > 0
        .FindControl(, conFun_Sample_unAuditing).Visible = InStr(mstrPrivs, "查看传染病报告") > 0
    End With
    GetDept 0

    Exit Function
errH:
    strErr = "出错函数(ShowMe),出错信息:" & Err.Number & " " & Err.Description
End Function

Private Sub BatchPrint()
          '功能   批量打印
          Dim intRow As Integer
          Dim strMsgShow As String
          Dim blnPrint As Boolean '当勾选多个报告时,是否已经打印了已出的报告 TRUE=已打印已出报告,False=未打印已出报告

1         On Error GoTo BatchPrint_Error

2         If checkDiseases = False Then
3             Exit Sub
4         End If

5         With vsfLeft

              '判断是否有未出的报告
6             For intRow = 1 To .Rows - 1
7                 If .Cell(flexcpChecked, intRow, .ColIndex("选择"), intRow, .ColIndex("选择")) = 1 Then
8                     If Trim(.TextMatrix(intRow, .ColIndex("报告"))) = "已出" Then
9                         blnPrint = True
10                    End If
11                    If Trim(.TextMatrix(intRow, .ColIndex("报告"))) = "未出" Then
12                        strMsgShow = "报告未出,请耐心等待"
13                    End If
14                End If
15            Next
16            If blnPrint = True And strMsgShow <> "" Then
17                strMsgShow = "有未出报告,请耐心等待"
18            End If


19            For intRow = 1 To .Rows - 1
20                If .Cell(flexcpChecked, intRow, .ColIndex("选择"), intRow, .ColIndex("选择")) = 1 Then
21                    If Val(.TextMatrix(intRow, .ColIndex("id"))) > 0 And Trim(.TextMatrix(intRow, .ColIndex("报告"))) = "已出" Then
22                        If .TextMatrix(intRow, .ColIndex("版本")) = 25 Then
23                            If Val(.TextMatrix(intRow, .ColIndex("传染病"))) = 1 Then
24                                If .TextMatrix(intRow, .ColIndex("复核人")) <> "" And .TextMatrix(intRow, .ColIndex("复核时间")) <> "" Then
25                                    PrintReport Me, Val(.TextMatrix(intRow, .ColIndex("id")))
26                                End If
27                            Else
28                                PrintReport Me, Val(.TextMatrix(intRow, .ColIndex("id")))
29                            End If
30                        Else
31                            PtintOldReport Me, Val(.TextMatrix(intRow, .ColIndex("id"))), Val(.TextMatrix(intRow, .ColIndex("病人id")))
32                        End If
33                    End If
34                End If
35            Next
36        End With

37        If strMsgShow <> "" Then
38            MsgBox strMsgShow, vbInformation, Me.Caption
39        End If

          '刷新界面
40        Call ReadPatientList(1)
41        Me.txtPatiNo.SetFocus
42        Call txtPatiNo_GotFocus


43        Exit Sub
BatchPrint_Error:
44        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(BatchPrint)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
45        Err.Clear
End Sub

Private Function checkDiseases() As Boolean
          '打印之前检查是否存在没有复核的传染病报告,若有,则打印中断
          Dim intRow As Integer
          Dim blnFindDiseases As Boolean '是否查找到为复核的传染病报告

1         On Error GoTo checkDiseases_Error

2         blnFindDiseases = False
3         With Me.vsfLeft
4             For intRow = 1 To .Rows - 1
5                 If .Cell(flexcpChecked, intRow, .ColIndex("选择"), intRow, .ColIndex("选择")) = 1 Then
6                     If Val(.TextMatrix(intRow, .ColIndex("id"))) > 0 Then
7                         If .TextMatrix(intRow, .ColIndex("版本")) = 25 Then
8                             If Val(.TextMatrix(intRow, .ColIndex("传染病"))) = 1 Then
9                                 If .TextMatrix(intRow, .ColIndex("复核人")) = "" Or .TextMatrix(intRow, .ColIndex("复核时间")) = "" Then
10                                    blnFindDiseases = True
11                                    .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
12                                End If
13                            End If
14                        End If
15                    End If
16                End If
17            Next
18        End With
19        If blnFindDiseases = True Then
20            MsgBox "发现存在未复核的传染病报告,打印中断", vbInformation, Me.Caption
21            checkDiseases = False
22            Exit Function
23        End If
24        checkDiseases = True


25        Exit Function
checkDiseases_Error:
26        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(checkDiseases)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear

End Function

Private Function PtintOldReport(objFrm As Object, lngSampleID As Long, Optional lngPaintID As Long, Optional byRunMode As Byte = 2, Optional strErr As String) As Boolean
          '打印设置
          Dim strReportCode As String
          Dim strReportParaNo As String
          Dim bytReportParaMode As Byte
          Dim rsTmp As New ADODB.Recordset
          Dim blnCurrMoved As Boolean
          Dim lng医嘱ID As Long, lng发送号 As Long
          Dim strSQL As String
          Dim strChart(0 To 8) As String

1         On Error GoTo PtintOldReport_Error

2         strSQL = "select 发送号, a.医嘱id from 病人医嘱发送 a , 病人医嘱记录 b,检验标本记录  c where b.id = a.医嘱id and  a.医嘱id =c.医嘱id  and c.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "报告打印", lngSampleID)
4         If rsTmp.EOF = False Then
5             lng发送号 = Val("" & rsTmp("发送号"))
6             lng医嘱ID = Val("" & rsTmp("医嘱id"))
7         End If

8         If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
9             If byRunMode = 3 Then
10                FunReportPrintSetHis gcnHisOracle, 100, strReportCode, objFrm
11            Else
12                If ReadSampleImage(lngSampleID, strChart, strErr, 10) = False Then
13                    Exit Function
14                End If
15                Call FunReportOpenHis(gcnHisOracle, 100, strReportCode, objFrm, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱ID, _
                                      "病人ID=" & lngPaintID, "标本ID=" & lngSampleID, _
                                      "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), "图形4=" & strChart(3), _
                                      "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                                      "图形9=" & strChart(8), byRunMode)

                  '对于审核过的标本标识
16                strSQL = "Zl_检验标本记录_标本质控(" & lngSampleID & ",'',1)"
17                Call ComExecuteProc(Sel_His_DB, strSQL, "打印标本")
19            End If
20        End If


21        Exit Function
PtintOldReport_Error:
22        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "执行(PtintOldReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
23        Err.Clear
End Function



Private Sub VsfMicrobe_SelChange()
    With VsfMicrobe
        If .ColIndex("id") <> -1 And .ColIndex("细菌ID") <> -1 Then
            If Val(.TextMatrix(.Row, .ColIndex("细菌ID"))) > 0 Then
                If mintVer = 25 Then
                    ReadSampleAntibiotic mlngKey, Val(.TextMatrix(.Row, .ColIndex("细菌ID")))
                    txtMicrobePositiveComment.Text = .TextMatrix(.Row, .ColIndex("阳性评语"))
                Else
                    ReadSampleAntibiotic mlngKey, Val(.TextMatrix(.Row, .ColIndex("细菌ID"))), 10
                End If
            End If
        End If
    End With
End Sub

Private Sub clsAllEdit()
    '功能清除所有显示
    On Error Resume Next
    vsfCenter.Row = 1
    VsfMicrobe.Row = 1
    VsfAntibiotic.Row = 1
    VSFContrast.Row = 1
    txtComment.Text = ""
    txtDiagnose.Text = ""
    Me.txtNormalMicrobe = ""
    Me.txtNoFindMicrobe = ""
    Me.txtNormalMicrobes = ""
    Me.chkPathopoiesiaGerm.value = 0
    Me.chkNoGerm.value = 0
    Me.txtMicroscope = ""
    Me.txtMicroscopeFinded = ""
    Me.txtMicroscopeNOFind = ""
    Me.txtMicrobePositiveComment = ""
    Me.txtGermComment = ""
    Me.txtSignificance = ""
End Sub

Private Function checkboxSource() As String
    '选择病人来源字符串
    Dim chktemp As CheckBox
    For Each chktemp In chkSource
        If chktemp.value = Checked Then
            checkboxSource = checkboxSource & "," & chktemp.Index
        End If
    Next
    If checkboxSource <> "" Then checkboxSource = Mid(checkboxSource, 2)
End Function

Public Function ShowMeto_New(objFrm As Object, Optional strErr As String, Optional lngPatientID As Long, Optional intPaitTyope As String) As Boolean
    On Error GoTo errH

    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 1013)
    Me.Show , objFrm
    ShowMeto_New = True


    '是否显示传染病筛选框
    picDiseases.Visible = InStr(mstrPrivs, "查看传染病报告") > 0
    Me.cboDiseases.ListIndex = 2

    With Me.cbrthis
        .FindControl(, conFun_Sample_Auditing).Visible = InStr(mstrPrivs, "查看传染病报告") > 0
        .FindControl(, conFun_Sample_unAuditing).Visible = InStr(mstrPrivs, "查看传染病报告") > 0
    End With

    mintIn = 1
    mlngPatientID = Val(lngPatientID)
    Label1.Visible = False
    dtpS.Visible = False
    dtpE.Visible = False
    DoEvents
    Call ReadPatientList(1)
    Exit Function
errH:
    strErr = "出错函数(ShowMe),出错信息:" & Err.Number & " " & Err.Description
End Function

Private Sub selAllText(ByVal objCrl As Object)
    With objCrl
        .SelStart = 0
        .SelLength = Len(objCrl.Text)
    End With
End Sub


