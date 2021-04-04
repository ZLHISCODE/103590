VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatientReprotBrowse 
   AutoRedraw      =   -1  'True
   Caption         =   "检验结果浏览"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17085
   Icon            =   "frmPatientReprotBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   17085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   14670
      TabIndex        =   37
      Top             =   345
      Width           =   14670
      Begin VB.Frame fraTop 
         Height          =   1035
         Left            =   45
         TabIndex        =   38
         Top             =   -75
         Width           =   14580
         Begin VB.ComboBox cboPatients 
            Height          =   300
            Left            =   4650
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   195
            Width           =   1335
         End
         Begin VB.CheckBox chkVerifyDate 
            Height          =   255
            Left            =   13860
            TabIndex        =   56
            Top             =   240
            Width           =   300
         End
         Begin VB.CheckBox chkApplyDate 
            Height          =   255
            Left            =   9825
            TabIndex        =   51
            Top             =   240
            Value           =   1  'Checked
            Width           =   300
         End
         Begin VB.ComboBox cboPages 
            Height          =   300
            Left            =   7065
            TabIndex        =   43
            Top             =   600
            Width           =   1335
         End
         Begin VB.ComboBox cboDept 
            Height          =   300
            Left            =   1095
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   195
            Width           =   2550
         End
         Begin VB.ComboBox cbodor 
            Height          =   300
            Left            =   1080
            TabIndex        =   41
            Text            =   "Combo1"
            Top             =   600
            Width           =   2550
         End
         Begin VB.TextBox txtPatiNo 
            Height          =   300
            Left            =   4650
            TabIndex        =   40
            Top             =   600
            Width           =   1305
         End
         Begin VB.TextBox txtDay 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   9630
            MaxLength       =   4
            TabIndex        =   39
            Text            =   "7"
            Top             =   660
            Width           =   510
         End
         Begin MSComCtl2.DTPicker dtpE 
            Height          =   300
            Left            =   8490
            TabIndex        =   44
            Top             =   195
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   117243907
            CurrentDate     =   40954
         End
         Begin MSComCtl2.DTPicker dtpVS 
            Height          =   300
            Left            =   11085
            TabIndex        =   53
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   117243907
            CurrentDate     =   40954
         End
         Begin MSComCtl2.DTPicker dtpVE 
            Height          =   300
            Left            =   12525
            TabIndex        =   54
            Top             =   195
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   117243907
            CurrentDate     =   40954
         End
         Begin MSComCtl2.DTPicker dtpS 
            Height          =   300
            Left            =   7065
            TabIndex        =   58
            Top             =   195
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   117243907
            CurrentDate     =   40954
         End
         Begin VB.Label Label5 
            Caption         =   "审核日期"
            Height          =   240
            Left            =   10260
            TabIndex        =   55
            Top             =   270
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "申请日期"
            Height          =   240
            Left            =   6255
            TabIndex        =   52
            Top             =   270
            Width           =   750
         End
         Begin VB.Label lblPages 
            AutoSize        =   -1  'True
            Caption         =   "住院次数"
            Height          =   180
            Left            =   6255
            TabIndex        =   50
            Top             =   660
            Width           =   720
         End
         Begin VB.Line Line1 
            X1              =   9600
            X2              =   10170
            Y1              =   855
            Y2              =   855
         End
         Begin VB.Label lblTimeOut 
            Caption         =   "可查看已出院       天的检验记录"
            Height          =   225
            Left            =   8490
            TabIndex        =   49
            Top             =   660
            Width           =   2925
         End
         Begin VB.Label lblDor 
            AutoSize        =   -1  'True
            Caption         =   "开单医生"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人"
            Height          =   180
            Left            =   4125
            TabIndex        =   47
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lblNo 
            AutoSize        =   -1  'True
            Caption         =   "住院号↓"
            Height          =   180
            Left            =   3840
            TabIndex        =   46
            Top             =   660
            Width           =   720
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            Caption         =   "申请科室↓"
            Height          =   180
            Left            =   120
            TabIndex        =   45
            Top             =   255
            Width           =   900
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2820
      Width           =   1035
      Begin C1Chart2D8.Chart2D chtPic 
         Height          =   705
         Index           =   0
         Left            =   180
         TabIndex        =   27
         Top             =   150
         Width           =   615
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   1085
         _ExtentY        =   1244
         _StockProps     =   0
         ControlProperties=   "frmPatientReprotBrowse.frx":6852
      End
   End
   Begin VB.PictureBox picComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4560
      Left            =   11310
      ScaleHeight     =   4560
      ScaleWidth      =   3495
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3390
      Width           =   3495
      Begin VB.TextBox txtSignificance 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Top             =   3315
         Width           =   3255
      End
      Begin VB.TextBox txtDiagnose 
         Appearance      =   0  'Flat
         Height          =   1155
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
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
         TabIndex        =   23
         Top             =   300
         Width           =   3255
      End
      Begin VB.Label lblSignificance 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "临床意义:"
         Height          =   180
         Left            =   30
         TabIndex        =   88
         Top             =   3060
         Width           =   810
      End
      Begin VB.Label lblDiagnose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "诊断:"
         Height          =   180
         Left            =   30
         TabIndex        =   24
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注:"
         Height          =   180
         Left            =   30
         TabIndex        =   22
         Top             =   50
         Width           =   450
      End
   End
   Begin VB.PictureBox PICContrast 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   30
      ScaleHeight     =   3015
      ScaleWidth      =   3495
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7320
      Width           =   3495
      Begin VB.PictureBox PicContrast_Top 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   30
         ScaleHeight     =   975
         ScaleWidth      =   3150
         TabIndex        =   16
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
            TabIndex        =   17
            Text            =   "30"
            Top             =   60
            Width           =   705
         End
         Begin VSFlex8Ctl.VSFlexGrid VSFContrast 
            Height          =   1335
            Left            =   60
            TabIndex        =   18
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
            TabIndex        =   20
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
            TabIndex        =   19
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
         TabIndex        =   11
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   8
            Width           =   1500
         End
         Begin C1Chart2D8.Chart2D chtContrast 
            Height          =   975
            Left            =   60
            TabIndex        =   25
            Top             =   300
            Width           =   1005
            _Version        =   524288
            _Revision       =   7
            _ExtentX        =   1773
            _ExtentY        =   1720
            _StockProps     =   0
            ControlProperties=   "frmPatientReprotBrowse.frx":6DE7
         End
         Begin VB.Label lblCht 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "图形内容"
            Height          =   180
            Left            =   30
            TabIndex        =   15
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
      TabIndex        =   4
      Top             =   2490
      Width           =   45
   End
   Begin VB.Frame FraLC 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5040
      Left            =   4620
      MousePointer    =   9  'Size W E
      TabIndex        =   3
      Top             =   1590
      Width           =   45
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   10980
      ScaleHeight     =   4335
      ScaleWidth      =   3795
      TabIndex        =   2
      Top             =   2340
      Width           =   3795
      Begin XtremeSuiteControls.TabControl TabPage 
         Height          =   3495
         Left            =   300
         TabIndex        =   9
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
         TabIndex        =   7
         Top             =   900
         Width           =   1815
      End
   End
   Begin VB.PictureBox picCenter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7065
      Left            =   1890
      ScaleHeight     =   7065
      ScaleWidth      =   14430
      TabIndex        =   1
      Top             =   1350
      Width           =   14430
      Begin VB.PictureBox PicNegative 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   7665
         Left            =   6120
         ScaleHeight     =   7635
         ScaleWidth      =   7005
         TabIndex        =   63
         Top             =   210
         Visible         =   0   'False
         Width           =   7035
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
            TabIndex        =   84
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
               TabIndex        =   85
               Top             =   270
               Width           =   5115
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
            TabIndex        =   77
            Top             =   2160
            Width           =   5250
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
               TabIndex        =   80
               Text            =   "显微镜检查"
               Top             =   270
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
               TabIndex        =   79
               Top             =   1350
               Width           =   3915
            End
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
               TabIndex        =   78
               Top             =   690
               Width           =   3915
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
               TabIndex        =   83
               Top             =   300
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
               TabIndex        =   82
               Top             =   1290
               Width           =   960
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
               TabIndex        =   81
               Top             =   660
               Width           =   960
            End
         End
         Begin VB.Frame frmNom 
            Height          =   2655
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   5250
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
               TabIndex        =   73
               Top             =   210
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
               TabIndex        =   72
               Top             =   975
               Width           =   4065
            End
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
               TabIndex        =   71
               Top             =   1800
               Width           =   4065
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
               TabIndex        =   76
               Top             =   210
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
               TabIndex        =   75
               Top             =   930
               Width           =   960
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
               TabIndex        =   74
               Top             =   1800
               Width           =   960
            End
         End
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
            TabIndex        =   64
            Top             =   1080
            Width           =   5250
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
               TabIndex        =   69
               Top             =   600
               Value           =   -1  'True
               Width           =   885
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
               TabIndex        =   68
               Top             =   600
               Width           =   885
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
               TabIndex        =   67
               Top             =   300
               Width           =   1815
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
               TabIndex        =   66
               Top             =   300
               Width           =   1815
            End
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
               TabIndex        =   65
               Top             =   300
               Width           =   1305
            End
         End
      End
      Begin VB.Frame fraCenter 
         Height          =   5745
         Left            =   60
         TabIndex        =   6
         Top             =   -60
         Width           =   5655
         Begin VB.PictureBox picSupplement 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   3105
            Left            =   300
            ScaleHeight     =   3075
            ScaleWidth      =   5265
            TabIndex        =   89
            Top             =   3750
            Visible         =   0   'False
            Width           =   5295
            Begin VSFlex8Ctl.VSFlexGrid vsfSupplement 
               Height          =   1335
               Left            =   120
               TabIndex        =   90
               Top             =   330
               Width           =   5055
               _cx             =   8916
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
               FocusRect       =   2
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   2
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   3
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
            Begin VB.Label lblSupplement 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "补充报告"
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
               Left            =   30
               TabIndex        =   91
               Top             =   0
               Width           =   960
            End
         End
         Begin VB.PictureBox picMicrobePositive 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4095
            Left            =   5400
            ScaleHeight     =   4095
            ScaleWidth      =   2865
            TabIndex        =   30
            Top             =   150
            Visible         =   0   'False
            Width           =   2865
            Begin VB.TextBox txtMicrobePositiveComment 
               Appearance      =   0  'Flat
               Height          =   855
               Left            =   30
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   36
               Top             =   3030
               Width           =   2715
            End
            Begin VSFlex8Ctl.VSFlexGrid VsfAntibiotic 
               Height          =   975
               Left            =   30
               TabIndex        =   31
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
               TabIndex        =   34
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
               TabIndex        =   35
               Top             =   2820
               Width           =   450
            End
            Begin VB.Label lblAntibiotic 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "抗生素:"
               Height          =   180
               Left            =   30
               TabIndex        =   33
               Top             =   1560
               Width           =   630
            End
            Begin VB.Label lblMicrobe 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "细菌:"
               Height          =   180
               Left            =   30
               TabIndex        =   32
               Top             =   30
               Width           =   450
            End
         End
         Begin VB.PictureBox picGeneral 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3780
            Left            =   435
            ScaleHeight     =   3780
            ScaleWidth      =   3210
            TabIndex        =   28
            Top             =   360
            Visible         =   0   'False
            Width           =   3210
            Begin VB.CheckBox chkGroup 
               Caption         =   "显示组合项目"
               Height          =   255
               Left            =   0
               TabIndex        =   86
               Top             =   0
               Width           =   1665
            End
            Begin VB.PictureBox picLab 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   45
               ScaleHeight     =   360
               ScaleWidth      =   3120
               TabIndex        =   61
               Top             =   2325
               Width           =   3120
               Begin VB.Label Label8 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "结果说明"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   134
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Left            =   45
                  TabIndex        =   62
                  Top             =   60
                  Width           =   960
               End
            End
            Begin VB.PictureBox picResultComment 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   945
               Left            =   30
               ScaleHeight     =   945
               ScaleWidth      =   3120
               TabIndex        =   59
               Top             =   2715
               Width           =   3120
               Begin VB.TextBox txtResultComment 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   555
                  Left            =   195
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   60
                  Top             =   120
                  Width           =   2385
               End
            End
            Begin VSFlex8Ctl.VSFlexGrid vsfCenter 
               Height          =   1035
               Left            =   180
               TabIndex        =   29
               Top             =   480
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
      Left            =   150
      ScaleHeight     =   4275
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   2370
      Width           =   4425
      Begin VB.Frame fraLeft 
         Height          =   4335
         Left            =   30
         TabIndex        =   5
         Top             =   -60
         Width           =   4185
         Begin VSFlex8Ctl.VSFlexGrid vsfLeft 
            Height          =   4035
            Left            =   150
            TabIndex        =   8
            Top             =   180
            Width           =   3735
            _cx             =   6588
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
   Begin MSComctlLib.ImageList imgVsf 
      Left            =   810
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
            Picture         =   "frmPatientReprotBrowse.frx":737C
            Key             =   "选择"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":DBDE
            Key             =   "查阅"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":14440
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":1ACA2
            Key             =   "禁止打印"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":21504
            Key             =   "部分审核"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   30
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatientReprotBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const const_PicRectBackColour As Long = &HE0E0E0
Private mblnShow As Boolean                                         '窗体是否显示
Dim mlngKey As Long                                                 '当前选择的标本ID
Dim mlngPatientID As Long                                           '病人ID
Dim mReportDate As Date                                             '核收时间
Dim mlngGetPatientID As Long                                        '上级传来的病人ID
Dim mintPatientType As Integer                                      '病人来源
Dim mlngPatientPage As Long                                         '病案主页
Dim mlngValueC As Long                                              '微生物结果次数
Dim mintVer As Integer                                              '版本，25-新版 10-老板
Dim mrsPatientVal As ADODB.Recordset                                '病人信息

Private mrsAntibioticValType As Recordset                           '药敏结果类型
Private mstrPrivs As String                                         '传入的上级的权限


Private mobjFSO As New Scripting.FileSystemObject    'FSO对象
Private objImg As Object
Private mblnShowBorder As Boolean                    '是否显示窗体的border

Private Sub cboDept_Click()
    If mintPatientType = 2 Then
        Call GetPatientsList
        Call getPartDor
    End If
End Sub

Private Sub cbodor_Click()
    ReadPatientList
End Sub

Private Sub cboPages_Click()
    With cboPages
        mlngPatientPage = Val(Trim(Replace(Replace(.Text, "第", ""), "次", "")))
        If .Text = "所有" Then
            Call readDate(False)
        Else
            Call readDate(True)
        End If
    End With
    Call ReadPatientList
End Sub

Private Sub cboPages_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cboPatients_Click()
    With cboPatients
        mlngGetPatientID = Val(.ItemData(.ListIndex))
        txtPatiNo = ""
        If Val(.ItemData(.ListIndex)) = 0 Then
            cboPages.Text = ""
            cboPages.Enabled = False
            Me.dtpS = Currentdate - 7: Me.dtpE = Currentdate
            Me.dtpVS = Currentdate - 7: Me.dtpVE = Currentdate
        Else
            ReadPatientVal mlngGetPatientID
            cboPages.Enabled = True
        End If
        ReadPatientList
    End With
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_SelAll
            If MsgBox("是否选择已打印的?", vbYesNo + vbQuestion + vbDefaultButton2, "中联软件") = vbYes Then
                VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("选择"), 1, True
            Else
                VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("选择"), 1, False
            End If
        Case ConMenu_Browse_ClsAll
            VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("选择"), 2, True
        Case ConMenu_Browse_Refresh
            ReadPatientList
        Case ConMenu_Browse_Print                                       '打印
            BatchPrint (2)
        Case ConMenu_Browse_PrintView                                   '打印预览
            BatchPrint (1)
        Case ConMenu_Browse_PrintSet                                    '打印设置
            If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("版本"))) = 25 Then
                PrintReport Me, mlngKey, 3
            Else
                PtintOldReport Me, mlngKey, , 3
            End If
        Case ConMenu_Browse_Exit
            Unload Me
        Case ConMenu_pop_In
            lblNo.Caption = "住院号↓"
        Case ConMenu_pop_bed
            lblNo.Caption = "床号↓"
        Case ConMenu_pop_Dept
            lblDept.Caption = "申请科室↓"
            InitDepts 0
        Case ConMenu_pop_DeptDistrict
            lblDept.Caption = "申请病区↓"
            InitDepts 1
        Case ConMenu_Browse_Find                                        '历次检验
            If mlngGetPatientID = 0 Then
                MsgBox "请选择一个病人！", vbInformation, "中联信息"
                cboPatients.SetFocus
            Else
                Call ReadPatientList(1)
            End If
        Case ConMenu_Appfor_ClincHelp       '诊疗参考
            Call ShowClincHelp
        Case ConMenu_Browse_PrintViewAll        '预览本次住院报告
            Call PrintAll(1)
        Case ConMenu_Browse_PrintAll            '打印本次住院报告
            Call PrintAll(2)
        Case ConMenu_Browse_PrintSetAll         '打印设置
            Call PrintAll(3)
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '外挂功能执行
            Call ExePlugIn(Control.Parameter, mlngKey)
    End Select
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-08-13
'功    能:  预览/打印病人说有报告
'入    参:
'           intType 1=预览，2=打印，3=打印设置
'出    参:
'返    回:
'调整影响:
'调用注意:
'---------------------------------------------------------------------------------------
Private Sub PrintAll(ByVal intType As Integer)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strNewSampleIDs As String
          Dim strOldSampleIDs As String
          Dim strID() As String
          Dim i As Integer
          
1         On Error GoTo PrintAll_Error

2         If intType = 1 Then
3             Call frmShowPatientAllReport.ShowMe(Me, mlngGetPatientID, Val(Trim(Replace(Replace(cboPages.Text, "第", ""), "次", ""))))
4         ElseIf intType = 2 Then
              '新版报告
5             strSQL = "Select f_List2str(Cast(Collect(to_char(ID)) As t_Strlist)) 标本ID" & vbCrLf & _
                     "   From 检验报告记录" & vbCrLf & _
                     "   Where HIS病人ID = [1] And 主页ID = [2] and 审核人 is not null"
6             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "新版报告", mlngGetPatientID, Val(Trim(Replace(Replace(cboPages.Text, "第", ""), "次", ""))))
7             If Not rsTmp.EOF Then
8                 strNewSampleIDs = rsTmp("标本ID") & ""
9             End If

              '老版报告
10            strSQL = "Select f_List2str(Cast(Collect(to_char(ID)) As t_Strlist)) 标本ID" & vbCrLf & _
                     "   From 检验标本记录" & vbCrLf & _
                     "   Where 病人ID = [1] And 主页ID = [2] and 审核人 is not null"
11            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "老版报告", mlngGetPatientID, Val(Trim(Replace(Replace(cboPages.Text, "第", ""), "次", ""))))
12            If Not rsTmp.EOF Then
13                strOldSampleIDs = rsTmp("标本ID") & ""
14            End If
              
              '打印新版报告
15            If strNewSampleIDs <> "" Then
16                FunReportOpen gcnLisOracle, 2500, "ZL25_INSIDE_2500_109", Me, "标本ID=" & strNewSampleIDs, intType
17                strID = Split(strNewSampleIDs, ",")
18                For i = 0 To UBound(strID)
19                    strSQL = "Zl_检验报告打印_Edit(1," & Val(strID(i)) & ",1)"
20                    Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
21                Next
22            End If
              '打印老版报告
23            If strOldSampleIDs <> "" Then
24                FunReportOpen gcnHisOracle, 100, "ZL1_INSIDE_1208_9", Me, "标本ID=" & strOldSampleIDs, intType
25            End If
              
              
26        Else
27            FunReportPrintSet gcnLisOracle, 2500, "ZL25_INSIDE_2500_109", Me
28            FunReportPrintSet gcnHisOracle, 100, "ZL1_INSIDE_1208_9", Me
29        End If


30        Exit Sub
PrintAll_Error:
31        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(PrintAll)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
32        Err.Clear
End Sub

Private Function VsfColAllSelAllcls(objVSF As VSFlexGrid, intCol As Integer, Optional intSel As Integer, Optional blnSelect As Boolean, Optional strErr As String) As Boolean
          '功能               全选或全清选择框
          '参数               intSel 0=安批一行进行判断 1=全部选中 2=全部不选中

          Dim intRow As Integer
1         On Error GoTo VsfColAllSelAllcls_Error

2         With objVSF
3             If intSel = 0 Then
4                 If .Rows = 1 Then Exit Function
5                 intSel = .Cell(flexcpChecked, 1, intCol, 1, intCol)
6                 If intSel = 1 Then
7                     intSel = 2
8                 Else
9                     intSel = 1
10                End If
11            End If
12            For intRow = 1 To .Rows - 1
13                 If .Cell(flexcpFontBold, intRow, .ColIndex("选择")) = False Then
14                    If blnSelect = True Then


15                       .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = intSel
16                    Else
17                       If Val(.TextMatrix(intRow, .ColIndex("打印"))) = 0 Then
18                           .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = intSel
19                       End If
20                    End If
21                End If
22            Next
23        End With
24        VsfColAllSelAllcls = True


25        Exit Function
VsfColAllSelAllcls_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(VsfColAllSelAllcls)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
27        Err.Clear

End Function

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Exit        '退出
            Control.Visible = mblnShowBorder
        Case ConMenu_Appfor_ClincHelp       '诊疗参考
            Control.Visible = VerCompare(gSysInfo.VersionHIS, "10.35.120") <> -1
        Case ConMenu_Browse_PrintViewAll    '预览病人所有报告
            Control.Visible = lblDept.Caption = "申请病区↓" And VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1
        Case ConMenu_Browse_PrintSetAll     '打印设置
            Control.Visible = lblDept.Caption = "申请病区↓" And VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1
        Case ConMenu_Browse_PrintAll        '打印病人所有报告
            Control.Visible = lblDept.Caption = "申请病区↓" And VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1
    End Select
End Sub

Private Sub chkApplyDate_Click()
    If chkApplyDate.value = 1 Then
        dtpS.Enabled = True
        dtpE.Enabled = True
    Else
        dtpS.Enabled = False
        dtpE.Enabled = False
    End If
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

Private Sub chkVerifyDate_Click()
    If chkVerifyDate.value = 1 Then
        dtpVS.Enabled = True
        dtpVE.Enabled = True
    Else
        dtpVS.Enabled = False
        dtpVE.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    If mblnShow = False Then
        InitFace
        ReadPatientVal mlngGetPatientID
        mblnShow = True
        ReadPatientList
    End If
End Sub

Private Sub Form_Load()
    Dim blnPrintReport As Boolean   '打印按钮是否可用
    Dim strTemp As String

    If InStr(";" & mstrPrivs & ";", ";打印检验报告;") > 0 Then
        blnPrintReport = True
    Else
        blnPrintReport = False
    End If

    '参数控制过滤条件申请和审核日期的可用状态
    strTemp = ComGetPara(Sel_Lis_DB, "过滤条件控制", 2500, 2001, "1|0")
    If strTemp <> "" Then
        chkApplyDate.value = Val(Split(strTemp, "|")(0))
        chkVerifyDate.value = Val(Split(strTemp, "|")(1))
    End If

    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
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
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Refresh, "刷新")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "打印预览")
        cbrControl.BeginGroup = True
        cbrControl.Enabled = blnPrintReport
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "打印设置")
        cbrControl.Enabled = blnPrintReport
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Print, "打印")
        cbrControl.Enabled = blnPrintReport
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintViewAll, "预览住院报告"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_PrintAll, "打印住院报告")
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSetAll, "打印设置  ")
        End With

        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "历次检验")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ClincHelp, "诊疗参考")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "退出")
        cbrControl.BeginGroup = True
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



    dtpS = Now - 7: dtpE = Now
    dtpVS = Now - 7: dtpVE = Now

    picLeft.Width = GetSetWith(1)

    Me.chkGroup.value = Val(ComGetPara(Sel_Lis_DB, "是否显示组合项目", gSysInfo.SysNo, gSysInfo.ModlNo, 1))

    ReadSampleBacteriology 0
    ReadSampleBacteriology 0
    ReadSampleVal 0

    Set mrsAntibioticValType = GetDictType("药敏结果类型")
End Sub

Private Function GetSetWith(ByVal intType As Integer) As Long
    '读取/设置窗体左边部分的宽度
    '1-读取,2-设置
    If intType = 1 Then
        GetSetWith = ComGetPara(Sel_Lis_DB, "临床检验报宽", 2500, 2500, "5000")
    ElseIf intType = 2 Then
        Call ComSetPara(Sel_Lis_DB, "临床检验报宽", picLeft.Width, 2500, 2500)
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
    mlngGetPatientID = 0
    mintPatientType = 0
    mlngPatientPage = 0
    mlngValueC = 0
    mintVer = 0
    mstrPrivs = ""
    Set objImg = Nothing
    Set mobjFSO = Nothing
    Set mrsAntibioticValType = Nothing
    Set mrsPatientVal = Nothing
    Call GetSetWith(2)
    Call ComSetPara(Sel_Lis_DB, "过滤条件控制", chkApplyDate.value & "|" & chkVerifyDate.value, 2500, 2001)
    Call ComSetPara(Sel_Lis_DB, "是否显示组合项目", Me.chkGroup.value, gSysInfo.SysNo, gSysInfo.ModlNo)
'    With vsfLeft
'        .Rows = 0
'    End With
End Sub

Private Sub FraCR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim minWidth As Single
    minWidth = 3000
    If Button = 1 Then
        If picCenter.Width + X < minWidth Then Exit Sub
        If picRight.Width - X < minWidth Then Exit Sub
        picCenter.Width = picCenter.Width + X
        FraCR.Left = picCenter.Left + picCenter.Width
        Form_Resize
    End If
End Sub

Private Sub FraCR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_Resize
End Sub

Private Sub FraLC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim minWidth As Single
    minWidth = 3000
    If Button = 1 Then
        If picLeft.Width + X < minWidth Then Exit Sub
        If picRight.Width - X < minWidth Then Exit Sub
        picLeft.Width = picLeft.Width + X
        FraLC.Left = picLeft.Left + picLeft.Width
        Form_Resize
    End If
End Sub

Private Sub FraLC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_Resize
End Sub

Private Sub lblContrast_Click()
    Call ReadContrastToVsf
End Sub

Private Sub lblDept_Click()
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    Dim vPoint As POINTAPI
    On Error Resume Next

    Set objPopup = Me.cbrthis.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Dept, "申请科室")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_DeptDistrict, "申请病区")
    End With
    vPoint.X = lblDept.Left / Screen.TwipsPerPixelX
    vPoint.Y = (lblDept.Top + lblDept.Height + 30) / Screen.TwipsPerPixelY
    ClientToScreen picFilter.hWnd, vPoint

    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
End Sub

Private Sub lblNo_Click()
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    Dim vPoint As POINTAPI
    On Error Resume Next

    If mintPatientType <> 2 Then Exit Sub

    Set objPopup = Me.cbrthis.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_In, "住院号")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_bed, "床号")
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
    On Error Resume Next
    With fraTop
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth - 50
        .Height = Me.picFilter.ScaleHeight + 10
    End With
End Sub

Private Sub picGeneral_Resize()
    On Error Resume Next
    With Me.chkGroup
        .Left = 0
        .Top = 0
        .Width = Me.picGeneral.Width
    End With

    With Me.picResultComment
        .Top = picGeneral.Height - .Height
        .Left = 0
        .Width = picGeneral.Width
    End With

     With Me.picLab
        .Top = picResultComment.Top - .Height
        .Left = 0
        .Width = picGeneral.Width
    End With

    With picSupplement
        .Left = 0
        .Top = picLab.Top - .Height
        .Width = picGeneral.Width
    End With

    With vsfCenter
        .Top = chkGroup.Top + chkGroup.Height
        .Left = 0
        .Width = picGeneral.Width
        If picSupplement.Visible = True Then
            .Height = picSupplement.Top
        Else
            .Height = picLab.Top
        End If
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
        .Top = 120
        .Left = 30
        .Width = fraLeft.Width - 70
        .Height = fraLeft.Height - 150
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
    On Error Resume Next
    ImageTypeSet 9
End Sub

Private Sub picResultComment_Resize()
    On Error Resume Next
    With Me.txtResultComment
        .Top = 100
        .Left = 100
        .Width = Me.picResultComment.Width - 200
        .Height = Me.picResultComment.Height - 200
    End With
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

Private Sub ReadPatientList(Optional ByVal intType As Integer)
      '功能           按条件读出病人列表
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim lngKey As Long
          Dim strDepts As String
          Dim strDept As String
          Dim strPatients As String
          Dim intDeptType As Integer
          Dim rsOldLisData As ADODB.Recordset
          Dim strFenLei As String
          Dim i As Integer
          Dim stridSQL As String
          Dim strTitle As String
          Dim var_tmp As Variant
          Dim var_SubTmp As Variant
          Dim lngLoop As Long
          Dim blnReadData As Boolean       '提示大范围查询数据之后是否继续查询

1         On Error GoTo ReadPatientList_Error

2         If mblnShow = False Then Exit Sub

3         strTitle = ComGetPara(Sel_Lis_DB, "检验报告显示列", 2500, 2001)

4         Call ReadSampleVal(0)
5         RefreshTab Me.TabPage.Selected.Index

6         If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
7             strSQL = "Select a.id,c.分类,0 选择,A.姓名, Decode(A.性别, '1', '男', '2', '女', '9', '未知', '') 性别, A.年龄, C.名称 申请项目,b.标本类型 标本类型, " & _
                     " A.住院号,a.门诊号, A.床号,B.申请时间,a.病人ID,a.核收时间,a.审核时间,a.备注,a.诊断,a.微生物,a.阳性报告," & _
                     " a.查阅,b.申请Id,a.医生站打印 打印 ,b.申请人,a.检验人,a.审核人, 25 版本, 0 结果次数, a.是否传染病,a.结果说明,a.部分审核人,nvl(d.补充报告状态,0) 补充报告" & vbNewLine & _
                     " From 检验报告记录 A, 检验申请组合 B, 检验组合项目 C,检验报告补充记录 D" & vbNewLine & _
                     " Where A.Id = B.标本id And B.组合id = C.Id(+) and a.id=d.标本ID(+) and (a.审核人 is not null or a.部分审核人 is not null) And b.组合id Is Not Null "
8         Else
9             strSQL = "Select a.id,c.分类,0 选择,A.姓名, Decode(A.性别, '1', '男', '2', '女', '9', '未知', '') 性别, A.年龄, C.名称 申请项目,b.标本类型 标本类型, " & _
                     " A.住院号,a.门诊号, A.床号,B.申请时间,a.病人ID,a.核收时间,a.审核时间,a.备注,a.诊断,a.微生物,a.阳性报告," & _
                     " a.查阅,b.申请Id,a.医生站打印 打印 ,b.申请人,a.检验人,a.审核人, 25 版本, 0 结果次数, a.是否传染病,a.结果说明,a.部分审核人" & vbNewLine & _
                     " From 检验报告记录 A, 检验申请组合 B, 检验组合项目 C" & vbNewLine & _
                     " Where A.Id = B.标本id And B.组合id = C.Id(+) and (a.审核人 is not null or a.部分审核人 is not null) And b.组合id Is Not Null "
10        End If

          '如果没有浏览手工报告的权限，则不允许浏览手工申请报告
11        If InStr(";" & mstrPrivs & ";", ";浏览手工报告;") <= 0 Then
12            strSQL = strSQL & " And b.申请id Is Not Null "
13        End If
14        If mlngGetPatientID > 0 Then
15            strSQL = strSQL & " and a.HIS病人ID = [4] "
              '        If mlngPatientPage <> 0 Then
              '            strSQL = strSQL & " and a.主页id = [8] "
              '        End If
16        End If

17        If chkApplyDate.value = 1 Then
18            strSQL = strSQL & " and b.申请时间 between [1] and [2] "
19        End If

20        If chkVerifyDate.value = 1 Then
21            strSQL = strSQL & " and a.审核时间 between [10] and [11] "
22        End If

23        blnReadData = True
          '    If chkVerifyDate.value = 1 Or chkApplyDate.value = 1 Then
          '高峰时段限制查询
24        If chkVerifyDate.value = 0 And chkApplyDate.value = 1 Then
25            If Not funCheckRushHours(2500, 2001, "浏览检验结果", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59")) Then
26                blnReadData = False
27            Else
28                blnReadData = True
29            End If
30        ElseIf chkVerifyDate.value = 1 And chkApplyDate.value = 0 Then
31            If Not funCheckRushHours(2500, 2001, "浏览检验结果", CDate(Format(dtpVS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpVE, "yyyy-MM-dd") & " 23:59:59")) Then
32                blnReadData = False
33            Else
34                blnReadData = True
35            End If
36        ElseIf chkVerifyDate.value = 1 And chkApplyDate.value = 1 Then
37            If Not funCheckRushHours(2500, 2001, "浏览检验结果", CDate(Format(IIf(dtpVS.value < dtpS.value, dtpVS.value, dtpS.value), "yyyy-MM-dd") & " 00:00:00"), CDate(Format(IIf(dtpVE.value < dtpE.value, dtpVE.value, dtpE.value), "yyyy-MM-dd") & " 23:59:59")) Then
38                blnReadData = False
39            Else
40                blnReadData = True
41            End If
42        End If
          '    End If

43        If blnReadData Then

44            If mintPatientType = 2 Then
45                If cboDept <> "00-所有科室" Then
46                    If mlngGetPatientID <= 0 Then
47                        If lblDept.Caption = "申请病区↓" Then
48                            intDeptType = 2
49                        Else
50                            intDeptType = 1
51                        End If
52                        strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
                          '当strPatients长度大于3500时,需要分解
53                        If Len(strPatients) >= 3500 Then
54                            stridSQL = MidPatients(strPatients)
55                        Else
56                            stridSQL = "Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list('" & strPatients & "') As Zltools.T_Numlist)) b"
57                        End If

58                        strSQL = strSQL & " and a.his病人id in (" & stridSQL & ") "
59                    End If
                      '            If lblDept.Caption = "申请病区↓" Then
                      '        '                strDepts = GetDepts(cboDept.ItemData(cboDept.ListIndex))
                      '        '                If strDepts = "" Then
                      '        '                    strDepts = "0"
                      '        '                End If
                      '                'strSQL = strSQL & " and (b.申请科室编码 in  (Select * From Table(Cast(F_Num2list([5]) As Zltools.T_Numlist))) or b.病区编码 = [6] )"
                      '                strSQL = strSQL & " and a.病区编码 = [6] "
                      '                strDept = Mid(Me.cboDept.Text, 1, InStr(cboDept, "-") - 1)
                      '            Else
                      '                strSQL = strSQL & " and a.申请科室编码 = [6] "
                      '                strDept = Mid(Me.cboDept.Text, 1, InStr(cboDept, "-") - 1)
                      '            End If
60                End If


61            End If
62            If Trim(txtPatiNo <> "") Then
63                If lblNo.Caption = "住院号↓" Then
64                    strSQL = strSQL & " and a.住院号 = [3] "
65                ElseIf lblNo.Caption = "床号↓" Then
66                    strSQL = strSQL & " and a.床号 = [3] "
67                Else
68                    strSQL = strSQL & " and a.门诊号 = [3] "
69                End If
70            End If

71            If Trim(cboPages.Text) <> "所有" And Trim(cboPages.Text) <> "" Then
72                strSQL = strSQL & " And (a.主页id = [8] or a.主页id is null)"
73            End If

74            If Trim(cbodor.Text) <> "所有" And Trim(cbodor.Text) <> "" Then
75                strSQL = strSQL & " and b.申请人 = [9] "
76            End If

77            If chkVerifyDate.value = 1 Then
78                strSQL = strSQL & " Order By c.分类, a.id, a.姓名, a.审核时间 Desc "
79            Else
80                strSQL = strSQL & " Order By c.分类, a.id, a.姓名, b.申请时间 Desc "
81            End If

82            If intType = 1 Then
83                If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
84                    strSQL = "Select a.id,c.分类,0 选择,A.姓名, Decode(A.性别, 1, '男', 2, '女', 9, '未知', '') 性别, A.年龄, C.名称 申请项目,b.标本类型 标本类型, " & _
                             " A.住院号,a.门诊号, A.床号,B.申请时间,a.病人ID,a.核收时间,a.审核时间,a.备注,a.诊断,a.微生物,a.阳性报告," & _
                             " a.查阅,b.申请Id,a.医生站打印 打印 ,b.申请人,a.检验人,a.审核人, 25 版本, 0 结果次数, a.是否传染病,a.结果说明,a.部分审核人,nvl(d.补充报告状态,0) 补充报告 " & _
                             " From 检验报告记录 A, 检验申请组合 B, 检验组合项目 C,检验报告补充记录 D" & _
                             " Where A.Id = B.标本id And B.组合id = C.Id(+) and a.id=d.标本ID(+) and (a.审核人 is not null or a.部分审核人 is not null) and a.病人id = [1] " & _
                             " Order By c.分类, a.姓名, a.性别, a.年龄, a.审核时间, c.名称"
85                Else
86                    strSQL = "Select a.id,c.分类,0 选择,A.姓名, Decode(A.性别, 1, '男', 2, '女', 9, '未知', '') 性别, A.年龄, C.名称 申请项目,b.标本类型 标本类型, " & _
                             " A.住院号,a.门诊号, A.床号,B.申请时间,a.病人ID,a.核收时间,a.审核时间,a.备注,a.诊断,a.微生物,a.阳性报告," & _
                             " a.查阅,b.申请Id,a.医生站打印 打印 ,b.申请人,a.检验人,a.审核人, 25 版本, 0 结果次数, a.是否传染病,a.结果说明,a.部分审核人" & _
                             " From 检验报告记录 A, 检验申请组合 B, 检验组合项目 C" & _
                             " Where A.Id = B.标本id And B.组合id = C.Id(+) and (a.审核人 is not null or a.部分审核人 is not null) and a.病人id = [1] " & _
                             " Order By c.分类, a.姓名, a.性别, a.年龄, a.审核时间, c.名称"
87                End If
88                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入病人列表", mlngGetPatientID)
89            Else
90                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入病人列表", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
                                         CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59"), txtPatiNo, mlngGetPatientID, strDepts, strDept, _
                                         strPatients, mlngPatientPage, cbodor.Text, CDate(Format(dtpVS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpVE, "yyyy-MM-dd") & " 23:59:59"))
91            End If

92        End If

93        With vsfLeft
94            .ExplorerBar = flexExSort
95            .Rows = 1
96            .FixedRows = 1
97            .OutlineBar = flexOutlineBarComplete
98            .OutlineCol = 1
99            .SubtotalPosition = flexSTAbove
100           If strTitle = "" Then
101               .Cols = 28
102               .ColKey(0) = "id": .ColWidth(.ColIndex("id")) = 2000: .ColAlignment(.ColIndex("id")) = flexAlignCenterCenter: .ColHidden(.ColIndex("id")) = True
103               .ColKey(1) = "选择": .ColWidth(.ColIndex("选择")) = 500: .ColAlignment(.ColIndex("选择")) = flexAlignCenterCenter    ': .ColDataType(.ColIndex("选择")) = flexDTBoolean
104               .ColKey(2) = "查阅": .ColWidth(.ColIndex("查阅")) = 250: .ColAlignment(.ColIndex("查阅")) = flexAlignCenterCenter
105               .ColKey(3) = "打印": .ColWidth(.ColIndex("打印")) = 250: .ColAlignment(.ColIndex("打印")) = flexAlignCenterCenter
106               .ColKey(4) = "姓名": .ColWidth(.ColIndex("姓名")) = 750: .ColAlignment(.ColIndex("姓名")) = flexAlignCenterCenter
107               .ColKey(5) = "性别": .ColWidth(.ColIndex("性别")) = 500: .ColAlignment(.ColIndex("性别")) = flexAlignCenterCenter
108               .ColKey(6) = "年龄": .ColWidth(.ColIndex("年龄")) = 500: .ColAlignment(.ColIndex("年龄")) = flexAlignCenterCenter
109               .ColKey(7) = "申请项目": .ColWidth(.ColIndex("申请项目")) = 2200: .ColAlignment(.ColIndex("申请项目")) = flexAlignCenterCenter
110               .ColKey(8) = "标本类型": .ColWidth(.ColIndex("标本类型")) = 1200: .ColAlignment(.ColIndex("标本类型")) = flexAlignCenterCenter
111               .ColKey(9) = "审核时间": .ColWidth(.ColIndex("审核时间")) = 2000: .ColAlignment(.ColIndex("审核时间")) = flexAlignCenterCenter
112               .ColKey(10) = "住院号": .ColWidth(.ColIndex("住院号")) = 750: .ColAlignment(.ColIndex("住院号")) = flexAlignCenterCenter
113               .ColKey(11) = "床号": .ColWidth(.ColIndex("床号")) = 500: .ColAlignment(.ColIndex("床号")) = flexAlignCenterCenter
114               .ColKey(12) = "申请时间": .ColWidth(.ColIndex("申请时间")) = 2000: .ColAlignment(.ColIndex("申请时间")) = flexAlignCenterCenter
115               .ColKey(13) = "病人ID": .ColWidth(.ColIndex("病人ID")) = 2000: .ColAlignment(.ColIndex("病人ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("病人ID")) = True
116               .ColKey(14) = "核收时间": .ColWidth(.ColIndex("核收时间")) = 2000: .ColAlignment(.ColIndex("核收时间")) = flexAlignCenterCenter: .ColHidden(.ColIndex("核收时间")) = True
117               .ColKey(15) = "备注": .ColWidth(.ColIndex("备注")) = 2000: .ColAlignment(.ColIndex("备注")) = flexAlignCenterCenter: .ColHidden(.ColIndex("备注")) = True
118               .ColKey(16) = "诊断": .ColWidth(.ColIndex("诊断")) = 2000: .ColAlignment(.ColIndex("诊断")) = flexAlignCenterCenter: .ColHidden(.ColIndex("诊断")) = True
119               .ColKey(17) = "微生物": .ColWidth(.ColIndex("微生物")) = 2000: .ColAlignment(.ColIndex("微生物")) = flexAlignCenterCenter: .ColHidden(.ColIndex("微生物")) = True
120               .ColKey(18) = "阳性报告": .ColWidth(.ColIndex("阳性报告")) = 2000: .ColAlignment(.ColIndex("阳性报告")) = flexAlignCenterCenter: .ColHidden(.ColIndex("阳性报告")) = True
121               .ColKey(19) = "申请Id": .ColWidth(.ColIndex("申请Id")) = 2000: .ColAlignment(.ColIndex("申请Id")) = flexAlignCenterCenter: .ColHidden(.ColIndex("申请Id")) = True
122               .ColKey(20) = "申请人": .ColWidth(.ColIndex("申请人")) = 750: .ColAlignment(.ColIndex("申请人")) = flexAlignCenterCenter
123               .ColKey(21) = "检验人": .ColWidth(.ColIndex("检验人")) = 750: .ColAlignment(.ColIndex("检验人")) = flexAlignCenterCenter
124               .ColKey(22) = "审核人": .ColWidth(.ColIndex("审核人")) = 750: .ColAlignment(.ColIndex("审核人")) = flexAlignCenterCenter
125               .ColKey(23) = "版本": .ColWidth(.ColIndex("版本")) = 750: .ColAlignment(.ColIndex("版本")) = flexAlignCenterCenter: .ColHidden(.ColIndex("版本")) = True
126               .ColKey(24) = "结果次数": .ColWidth(.ColIndex("结果次数")) = 750: .ColAlignment(.ColIndex("结果次数")) = flexAlignCenterCenter: .ColHidden(.ColIndex("结果次数")) = True
127               .ColKey(25) = "结果说明": .ColWidth(.ColIndex("结果说明")) = 750: .ColAlignment(.ColIndex("结果说明")) = flexAlignCenterCenter: .ColHidden(.ColIndex("结果说明")) = True
128               .ColKey(26) = "部分审核人": .ColWidth(.ColIndex("部分审核人")) = 0: .ColAlignment(.ColIndex("部分审核人")) = flexAlignCenterCenter: .ColHidden(.ColIndex("部分审核人")) = True
129               .ColKey(27) = "补充报告": .ColWidth(.ColIndex("补充报告")) = 0: .ColAlignment(.ColIndex("补充报告")) = flexAlignCenterCenter: .ColHidden(.ColIndex("补充报告")) = True
130           Else
131               var_tmp = Split(strTitle, ";")
132               .Rows = 1
133               .Cols = UBound(var_tmp) + 1
134               For lngLoop = LBound(var_tmp) To UBound(var_tmp)
135                   var_SubTmp = Split(var_tmp(lngLoop), ",")
136                   .ColKey(lngLoop) = var_SubTmp(0): .ColWidth(.ColIndex(var_SubTmp(0))) = var_SubTmp(1): .ColAlignment(.ColIndex(var_SubTmp(0))) = flexAlignCenterCenter: .ColHidden(.ColIndex(var_SubTmp(0))) = Not (Val(var_SubTmp(2)) = 1)
                      '                If var_SubTmp(0) = "复核人" Or var_SubTmp(0) = "复核时间" Then
                      '                    .ColHidden(.ColIndex(var_SubTmp(0))) = Not InStr(mstrPrivs, "查看传染病报告") > 0
                      '                End If
137               Next
138               If .ColIndex("补充报告") < 0 Then
139                   .Cols = .Cols + 1
140                   .ColKey(.Cols - 1) = "补充报告": .ColWidth(.ColIndex("补充报告")) = 0: .ColAlignment(.ColIndex("补充报告")) = flexAlignCenterCenter: .ColHidden(.ColIndex("补充报告")) = True
141               End If
142           End If
143           .TextMatrix(0, .ColIndex("选择")) = ""
144           .TextMatrix(0, .ColIndex("查阅")) = ""
145           .TextMatrix(0, .ColIndex("打印")) = ""
146           .TextMatrix(0, .ColIndex("姓名")) = "姓名"
147           .TextMatrix(0, .ColIndex("性别")) = "性别"
148           .TextMatrix(0, .ColIndex("年龄")) = "年龄"
149           .TextMatrix(0, .ColIndex("申请项目")) = "申请项目"
150           .TextMatrix(0, .ColIndex("标本类型")) = "标本类型"
151           .TextMatrix(0, .ColIndex("审核时间")) = "审核时间"
152           .TextMatrix(0, .ColIndex("住院号")) = "住院号"
153           .TextMatrix(0, .ColIndex("床号")) = "床号"
154           .TextMatrix(0, .ColIndex("申请时间")) = "申请时间"
155           .TextMatrix(0, .ColIndex("申请人")) = "申请人"
156           .TextMatrix(0, .ColIndex("检验人")) = "检验人"
157           .TextMatrix(0, .ColIndex("审核人")) = "审核人"
158           .TextMatrix(0, .ColIndex("版本")) = "版本"
159           .TextMatrix(0, .ColIndex("结果次数")) = "结果次数"
160           .Row = 0: .Col = .ColIndex("选择"): .CellPicture = imgVsf.ListImages("选择").ExtractIcon
161           .Row = 0: .Col = .ColIndex("查阅"): .CellPicture = imgVsf.ListImages("查阅").ExtractIcon
162           .Row = 0: .Col = .ColIndex("打印"): .CellPicture = imgVsf.ListImages("打印").ExtractIcon

163           If blnReadData Then

164               Do Until rsTmp.EOF
165                   If rsTmp("申请人") & "" <> gUserInfo.Name And rsTmp("申请人") & "" <> "" And InStr(";" & mstrPrivs & ";", ";查看传染病报告;") <= 0 And Val(rsTmp("是否传染病") & "") = 1 Then
                          '仅有查看传染病报告权限时才允查看传染病报告,否则跳过传染病报告继续加载下一条数据
166                       rsTmp.MoveNext
167                   Else
168                       If InStr(strFenLei, rsTmp("分类") & ",") > 0 Then
169                           If lngKey <> Val(rsTmp("id") & "") Then
170                               .Rows = .Rows + 1
171                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""

172                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 2
173                               .TextMatrix(.Rows - 1, .ColIndex("打印")) = rsTmp("打印") & ""

174                               If rsTmp("查阅") & "" = 1 Then
175                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("查阅")) = imgVsf.ListImages("查阅").ExtractIcon
176                               End If

177                               If Val(rsTmp("打印") & "") > 0 Then
178                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("打印")) = imgVsf.ListImages("打印").ExtractIcon
179                               End If

180                               .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsTmp("姓名") & ""
181                               .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsTmp("性别") & ""
182                               .TextMatrix(.Rows - 1, .ColIndex("年龄")) = rsTmp("年龄") & ""
183                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsTmp("申请项目") & ""
184                               .TextMatrix(.Rows - 1, .ColIndex("标本类型")) = rsTmp("标本类型") & ""
185                               .TextMatrix(.Rows - 1, .ColIndex("审核时间")) = Format(rsTmp("审核时间") & "", "yyyy-mm-dd HH:mm:ss")
186                               .TextMatrix(.Rows - 1, .ColIndex("住院号")) = rsTmp("住院号") & ""
187                               .TextMatrix(.Rows - 1, .ColIndex("床号")) = rsTmp("床号") & ""
188                               .TextMatrix(.Rows - 1, .ColIndex("申请时间")) = rsTmp("申请时间") & ""
189                               .TextMatrix(.Rows - 1, .ColIndex("病人ID")) = rsTmp("病人ID") & ""
190                               .TextMatrix(.Rows - 1, .ColIndex("核收时间")) = rsTmp("核收时间") & ""
191                               .TextMatrix(.Rows - 1, .ColIndex("备注")) = rsTmp("备注") & ""
192                               .TextMatrix(.Rows - 1, .ColIndex("诊断")) = rsTmp("诊断") & ""
193                               .TextMatrix(.Rows - 1, .ColIndex("微生物")) = rsTmp("微生物") & ""
194                               .TextMatrix(.Rows - 1, .ColIndex("阳性报告")) = rsTmp("阳性报告") & ""
195                               .TextMatrix(.Rows - 1, .ColIndex("申请Id")) = rsTmp("申请Id") & ""
196                               .TextMatrix(.Rows - 1, .ColIndex("申请人")) = rsTmp("申请人") & ""
197                               .TextMatrix(.Rows - 1, .ColIndex("检验人")) = rsTmp("检验人") & ""
198                               .TextMatrix(.Rows - 1, .ColIndex("审核人")) = rsTmp("审核人") & ""
199                               .TextMatrix(.Rows - 1, .ColIndex("版本")) = rsTmp("版本") & ""
200                               .TextMatrix(.Rows - 1, .ColIndex("结果次数")) = rsTmp("结果次数") & ""
201                               .TextMatrix(.Rows - 1, .ColIndex("结果说明")) = rsTmp("结果说明") & ""
202                               .TextMatrix(.Rows - 1, .ColIndex("部分审核人")) = rsTmp("部分审核人") & ""
203                               If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then .TextMatrix(.Rows - 1, .ColIndex("补充报告")) = rsTmp("补充报告") & ""
204                           Else
205                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = .TextMatrix(.Rows - 1, .ColIndex("申请项目")) & "," & rsTmp("申请项目") & ""
206                           End If
207                           lngKey = Val(rsTmp("id") & "")
208                       Else
209                           If lngKey <> Val(rsTmp("id") & "") Then
210                               .Rows = .Rows + 2

211                               For i = 1 To .Cols - 1
212                                   .TextMatrix(.Rows - 2, i) = CStr(rsTmp("分类") & "") & "(新版)"
213                               Next

                                  '合并
214                               .MergeRow(.Rows - 2) = True
215                               .MergeCellsFixed = flexMergeRestrictRows

                                  '缩进
216                               .IsSubtotal(.Rows - 2) = True
217                               .RowOutlineLevel(.Rows - 2) = 1

                                  '加粗
218                               .Cell(flexcpFontBold, .Rows - 2, 1) = True

219                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""

220                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 2
221                               .TextMatrix(.Rows - 1, .ColIndex("打印")) = rsTmp("打印") & ""

222                               If rsTmp("查阅") & "" = 1 Then
223                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("查阅")) = imgVsf.ListImages("查阅").ExtractIcon
224                               End If

225                               If Val(rsTmp("打印") & "") > 0 Then
226                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("打印")) = imgVsf.ListImages("打印").ExtractIcon
227                               End If

228                               .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsTmp("姓名") & ""
229                               .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsTmp("性别") & ""
230                               .TextMatrix(.Rows - 1, .ColIndex("年龄")) = rsTmp("年龄") & ""
231                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsTmp("申请项目") & ""
232                               .TextMatrix(.Rows - 1, .ColIndex("标本类型")) = rsTmp("标本类型") & ""
233                               .TextMatrix(.Rows - 1, .ColIndex("审核时间")) = Format(rsTmp("审核时间") & "", "yyyy-mm-dd HH:mm:ss")
234                               .TextMatrix(.Rows - 1, .ColIndex("住院号")) = rsTmp("住院号") & ""
235                               .TextMatrix(.Rows - 1, .ColIndex("床号")) = rsTmp("床号") & ""
236                               .TextMatrix(.Rows - 1, .ColIndex("申请时间")) = rsTmp("申请时间") & ""
237                               .TextMatrix(.Rows - 1, .ColIndex("病人ID")) = rsTmp("病人ID") & ""
238                               .TextMatrix(.Rows - 1, .ColIndex("核收时间")) = rsTmp("核收时间") & ""
239                               .TextMatrix(.Rows - 1, .ColIndex("备注")) = rsTmp("备注") & ""
240                               .TextMatrix(.Rows - 1, .ColIndex("诊断")) = rsTmp("诊断") & ""
241                               .TextMatrix(.Rows - 1, .ColIndex("微生物")) = rsTmp("微生物") & ""
242                               .TextMatrix(.Rows - 1, .ColIndex("阳性报告")) = rsTmp("阳性报告") & ""
243                               .TextMatrix(.Rows - 1, .ColIndex("申请Id")) = rsTmp("申请Id") & ""
244                               .TextMatrix(.Rows - 1, .ColIndex("申请人")) = rsTmp("申请人") & ""
245                               .TextMatrix(.Rows - 1, .ColIndex("检验人")) = rsTmp("检验人") & ""
246                               .TextMatrix(.Rows - 1, .ColIndex("审核人")) = rsTmp("审核人") & ""
247                               .TextMatrix(.Rows - 1, .ColIndex("版本")) = rsTmp("版本") & ""
248                               .TextMatrix(.Rows - 1, .ColIndex("结果次数")) = rsTmp("结果次数") & ""
249                               .TextMatrix(.Rows - 1, .ColIndex("结果说明")) = rsTmp("结果说明") & ""
250                               .TextMatrix(.Rows - 1, .ColIndex("部分审核人")) = rsTmp("部分审核人") & ""
251                               If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then .TextMatrix(.Rows - 1, .ColIndex("补充报告")) = rsTmp("补充报告") & ""
252                           Else
253                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = .TextMatrix(.Rows - 1, .ColIndex("申请项目")) & "," & rsTmp("申请项目") & ""
254                           End If
255                           lngKey = Val(rsTmp("id") & "")
256                           strFenLei = strFenLei & rsTmp("分类") & ","
257                       End If
258                       rsTmp.MoveNext
259                   End If
260               Loop

261               strFenLei = ""
262               Set rsOldLisData = GetOldLisData(intType)
263               If rsOldLisData.RecordCount > 0 Then
264                   Do Until rsOldLisData.EOF
265                       If InStr(strFenLei, rsOldLisData("分类") & ",") > 0 Then
266                           If lngKey <> Val(rsOldLisData("id") & "") Then
267                               .Rows = .Rows + 1
268                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsOldLisData("id") & ""
269                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 2
270                               .TextMatrix(.Rows - 1, .ColIndex("打印")) = rsOldLisData("打印") & ""

271                               If rsOldLisData("查阅") & "" = 1 Then
272                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("查阅")) = imgVsf.ListImages("查阅").ExtractIcon
273                               End If

274                               If Val(rsOldLisData("打印") & "") > 0 Then
275                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("打印")) = imgVsf.ListImages("打印").ExtractIcon
276                               End If

277                               .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsOldLisData("姓名") & ""
278                               .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsOldLisData("性别") & ""
279                               .TextMatrix(.Rows - 1, .ColIndex("年龄")) = rsOldLisData("年龄") & ""
280                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsOldLisData("申请项目") & ""
281                               .TextMatrix(.Rows - 1, .ColIndex("标本类型")) = rsOldLisData("标本类型") & ""
282                               .TextMatrix(.Rows - 1, .ColIndex("审核时间")) = Format(rsOldLisData("审核时间") & "", "yyyy-mm-dd HH:mm:ss")
283                               .TextMatrix(.Rows - 1, .ColIndex("住院号")) = rsOldLisData("住院号") & ""
284                               .TextMatrix(.Rows - 1, .ColIndex("床号")) = rsOldLisData("床号") & ""
285                               .TextMatrix(.Rows - 1, .ColIndex("申请时间")) = rsOldLisData("申请时间") & ""
286                               .TextMatrix(.Rows - 1, .ColIndex("病人ID")) = rsOldLisData("病人ID") & ""
287                               .TextMatrix(.Rows - 1, .ColIndex("核收时间")) = rsOldLisData("核收时间") & ""
288                               .TextMatrix(.Rows - 1, .ColIndex("备注")) = rsOldLisData("备注") & ""
289                               .TextMatrix(.Rows - 1, .ColIndex("诊断")) = rsOldLisData("诊断") & ""
290                               .TextMatrix(.Rows - 1, .ColIndex("微生物")) = rsOldLisData("微生物") & ""
291                               .TextMatrix(.Rows - 1, .ColIndex("阳性报告")) = rsOldLisData("阳性报告") & ""
292                               .TextMatrix(.Rows - 1, .ColIndex("申请Id")) = rsOldLisData("申请Id") & ""
293                               .TextMatrix(.Rows - 1, .ColIndex("申请人")) = rsOldLisData("申请人") & ""
294                               .TextMatrix(.Rows - 1, .ColIndex("检验人")) = rsOldLisData("检验人") & ""
295                               .TextMatrix(.Rows - 1, .ColIndex("审核人")) = rsOldLisData("审核人") & ""
296                               .TextMatrix(.Rows - 1, .ColIndex("版本")) = rsOldLisData("版本") & ""
297                               .TextMatrix(.Rows - 1, .ColIndex("结果次数")) = rsOldLisData("结果次数") & ""
298                               .TextMatrix(.Rows - 1, .ColIndex("结果次数")) = rsOldLisData("结果次数") & ""
299                           Else
300                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = .TextMatrix(.Rows - 1, .ColIndex("申请项目")) & "," & rsOldLisData("申请项目") & ""
301                           End If
302                           lngKey = Val(rsOldLisData("id") & "")
303                       Else
304                           If lngKey <> Val(rsOldLisData("id") & "") Then
305                               .Rows = .Rows + 2
306                               For i = 1 To .Cols - 1
307                                   .TextMatrix(.Rows - 2, i) = CStr(rsOldLisData("分类") & "") & "(老版)"
308                               Next

                                  '合并
309                               .MergeRow(.Rows - 2) = True
310                               .MergeCellsFixed = flexMergeRestrictRows

                                  '缩进
311                               .IsSubtotal(.Rows - 2) = True
312                               .RowOutlineLevel(.Rows - 2) = 1

                                  '加粗
313                               .Cell(flexcpFontBold, .Rows - 2, 1) = True

314                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsOldLisData("id") & ""
315                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 2
316                               .TextMatrix(.Rows - 1, .ColIndex("打印")) = rsOldLisData("打印") & ""

317                               If rsOldLisData("查阅") & "" = 1 Then
318                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("查阅")) = imgVsf.ListImages("查阅").ExtractIcon
319                               End If

320                               If Val(rsOldLisData("打印") & "") > 0 Then
321                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("打印")) = imgVsf.ListImages("打印").ExtractIcon
322                               End If

323                               .TextMatrix(.Rows - 1, .ColIndex("姓名")) = rsOldLisData("姓名") & ""
324                               .TextMatrix(.Rows - 1, .ColIndex("性别")) = rsOldLisData("性别") & ""
325                               .TextMatrix(.Rows - 1, .ColIndex("年龄")) = rsOldLisData("年龄") & ""
326                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = rsOldLisData("申请项目") & ""
327                               .TextMatrix(.Rows - 1, .ColIndex("标本类型")) = rsOldLisData("标本类型") & ""
328                               .TextMatrix(.Rows - 1, .ColIndex("审核时间")) = Format(rsOldLisData("审核时间") & "", "yyyy-mm-dd HH:mm:ss")
329                               .TextMatrix(.Rows - 1, .ColIndex("住院号")) = rsOldLisData("住院号") & ""
330                               .TextMatrix(.Rows - 1, .ColIndex("床号")) = rsOldLisData("床号") & ""
331                               .TextMatrix(.Rows - 1, .ColIndex("申请时间")) = rsOldLisData("申请时间") & ""
332                               .TextMatrix(.Rows - 1, .ColIndex("病人ID")) = rsOldLisData("病人ID") & ""
333                               .TextMatrix(.Rows - 1, .ColIndex("核收时间")) = rsOldLisData("核收时间") & ""
334                               .TextMatrix(.Rows - 1, .ColIndex("备注")) = rsOldLisData("备注") & ""
335                               .TextMatrix(.Rows - 1, .ColIndex("诊断")) = rsOldLisData("诊断") & ""
336                               .TextMatrix(.Rows - 1, .ColIndex("微生物")) = rsOldLisData("微生物") & ""
337                               .TextMatrix(.Rows - 1, .ColIndex("阳性报告")) = rsOldLisData("阳性报告") & ""
338                               .TextMatrix(.Rows - 1, .ColIndex("申请Id")) = rsOldLisData("申请Id") & ""
339                               .TextMatrix(.Rows - 1, .ColIndex("申请人")) = rsOldLisData("申请人") & ""
340                               .TextMatrix(.Rows - 1, .ColIndex("检验人")) = rsOldLisData("检验人") & ""
341                               .TextMatrix(.Rows - 1, .ColIndex("审核人")) = rsOldLisData("审核人") & ""
342                               .TextMatrix(.Rows - 1, .ColIndex("版本")) = rsOldLisData("版本") & ""
343                               .TextMatrix(.Rows - 1, .ColIndex("结果次数")) = rsOldLisData("结果次数") & ""
344                           Else
345                               .TextMatrix(.Rows - 1, .ColIndex("申请项目")) = .TextMatrix(.Rows - 1, .ColIndex("申请项目")) & "," & rsOldLisData("申请项目") & ""
346                           End If
347                           lngKey = Val(rsOldLisData("id") & "")
348                           strFenLei = strFenLei & rsOldLisData("分类") & ","
349                       End If
350                       rsOldLisData.MoveNext
351                   Loop
352               End If

353               If (rsTmp.RecordCount > 0 Or rsOldLisData.RecordCount > 0) And .Rows > 1 Then
354                   .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
355                   .Row = 1
356               Else
357                   .Rows = 2
358                   .Row = 1
359               End If
360               mlngKey = 0


361               vsfLeft_SelChange

362               If intType = 0 Then
363                   If chkVerifyDate.value = 1 Then
364                       .Cell(flexcpSort, .FixedRows, .ColIndex("审核时间"), .Rows - 1, .ColIndex("审核时间")) = 2
365                       .Cell(flexcpSort, .FixedRows, .ColIndex("姓名"), .Rows - 1, .ColIndex("姓名")) = 1
366                   Else
367                       .Cell(flexcpSort, .FixedRows, .ColIndex("申请时间"), .Rows - 1, .ColIndex("申请时间")) = 2
368                       .Cell(flexcpSort, .FixedRows, .ColIndex("姓名"), .Rows - 1, .ColIndex("姓名")) = 1
369                   End If
370               End If
371           End If
372       End With


373       Exit Sub
ReadPatientList_Error:
374       Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(ReadPatientList)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
375       Err.Clear

End Sub

Private Function MidPatients(ByVal strPatients As String)
'当strPatients长度大于3500时,需要分解
    Dim strPatientsCut As String
    Dim strLast As String
    Dim stridSQL As String

    Do While Len(strPatients) >= 3500
        strPatientsCut = Mid(strPatients, 1, 3500)
        strPatients = Mid(strPatients, 3501)
        strLast = Mid(strPatientsCut, InStrRev(strPatientsCut, ","))
        strPatientsCut = Mid(strPatientsCut, 1, InStrRev(strPatientsCut, ",") - 1)
        strPatients = strLast & strPatients

        If Mid(strPatientsCut, 1, 1) = "," Then
            strPatientsCut = Mid(strPatientsCut, 2)
        End If

        stridSQL = stridSQL & " Union All " & "Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list('" & strPatientsCut & "') As Zltools.T_Numlist)) b"
    Loop

    If stridSQL <> "" Then
        If Mid(strPatients, 1, 1) = "," Then
            strPatients = Mid(strPatients, 2)
        End If
        stridSQL = Mid(stridSQL, 12) & " Union All " & "Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list('" & strPatients & "') As Zltools.T_Numlist)) b"
    End If

    MidPatients = stridSQL
End Function

Private Function GetOldLisData(Optional ByVal intType As Integer) As ADODB.Recordset
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim strDepts As String
          Dim strDept As String
          Dim strPatients As String
          Dim intDeptType As Integer
          Dim stridSQL As String
          Dim strWhere As String

1         On Error GoTo GetOldLisData_Error

2         strSQL = "Select Distinct a.Id, a.操作类型 分类, 0 选择, a.姓名, a.性别, a.年龄, a.检验项目 申请项目, a.标本类型, a.住院号, a.门诊号, a.床号, a.申请时间, a.病人id, a.核收时间," & vbNewLine & _
                  "                a.审核时间, a.检验备注 备注, d.诊断, a.微生物标本 微生物, 1 阳性报告, 0 查阅, a.医嘱id 申请id, a.打印次数 打印, a.申请人, a.检验人, a.审核人, 10 版本," & vbNewLine & _
                  "                a.报告结果 结果次数" & vbNewLine & _
                  "From 检验标本记录 A, 病人医嘱记录 C," & vbNewLine & _
                  "     (Select b.医嘱id 医嘱id, f_List2str(Cast(Collect(b.项目 || ':' || b.内容) As t_Strlist)) 诊断" & vbNewLine & _
                  "       From 检验标本记录 A, 病人医嘱附件 B" & vbNewLine & _
                  "       Where a.医嘱id = b.医嘱id  [条件]" & vbNewLine & _
                  "       Group By b.医嘱id) D" & vbNewLine & _
                  "Where a.医嘱id = c.Id(+) And a.医嘱id = d.医嘱id(+) And a.审核人 Is Not Null"

3         If mlngGetPatientID > 0 Then
4             strWhere = strWhere & " and a.病人ID = [4] "
5             strSQL = strSQL & " and a.病人ID = [4] "
6         End If

7         If chkApplyDate.value = 1 Then
8             strWhere = strWhere & " and a.申请时间 between [1] and [2] "
9             strSQL = strSQL & " and a.申请时间 between [1] and [2] "
10        End If

11        If chkVerifyDate.value = 1 Then
12            strWhere = strWhere & " and a.审核时间 between [10] and [11] "
13            strSQL = strSQL & " and a.审核时间 between [10] and [11] "
14        End If

15        If mintPatientType = 2 Then
16            If cboDept <> "00-所有科室" Then
17                If mlngGetPatientID <= 0 Then
18                    If lblDept.Caption = "申请病区↓" Then
19                        intDeptType = 2
20                    Else
21                        intDeptType = 1
22                    End If
23                    strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
                      '当strPatients长度大于3500时,需要分解
24                    If Len(strPatients) >= 3500 Then
25                        stridSQL = MidPatients(strPatients)
26                    Else
27                        stridSQL = "Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list('" & strPatients & "') As Zltools.T_Numlist)) b"
28                    End If

29                    strWhere = strWhere & " and a.病人id in (" & stridSQL & ") "
30                    strSQL = strSQL & " and a.病人id in (" & stridSQL & ") "

31                End If

32            End If


33        End If
34        If Trim(txtPatiNo <> "") Then
35            If lblNo.Caption = "住院号↓" Then
36                strWhere = strWhere & " and a.住院号 = [3] "
37                strSQL = strSQL & " and a.住院号 = [3] "

38            ElseIf lblNo.Caption = "床号↓" Then
39                strWhere = strWhere & " and a.床号 = [3] "
40                strSQL = strSQL & " and a.床号 = [3] "

41            Else
42                strWhere = strWhere & " and a.门诊号 = [3] "
43                strSQL = strSQL & " and a.门诊号 = [3] "

44            End If
45        End If

46        If Trim(cbodor.Text) <> "所有" And Trim(cbodor.Text) <> "" Then
47            strWhere = strWhere & " and a.申请人 = [9] "
48            strSQL = strSQL & " and a.申请人 = [9] "
49        End If

50        If chkVerifyDate.value = 1 Then
51            strSQL = strSQL & " Order By a.操作类型, a.id, a.姓名, a.审核时间 Desc "
52        Else
53            strSQL = strSQL & " Order By a.操作类型, a.id, a.姓名, a.申请时间 Desc "
54        End If

55        strSQL = Replace(strSQL, "[条件]", strWhere)

56        If intType = 1 Then
57            strSQL = "Select Distinct a.Id, a.操作类型 分类, 0 选择, a.姓名, a.性别, a.年龄, a.检验项目 申请项目, a.标本类型, a.住院号, a.门诊号, a.床号, a.申请时间, a.病人id, a.核收时间," & vbNewLine & _
                      "                a.审核时间, a.检验备注 备注, d.诊断, a.微生物标本 微生物, 1 阳性报告, 0 查阅, a.医嘱id 申请id, a.打印次数 打印, a.申请人, a.检验人, a.审核人, 10 版本," & vbNewLine & _
                      "                a.报告结果 结果次数" & vbNewLine & _
                      "From 检验标本记录 A, 病人医嘱记录 C," & vbNewLine & _
                      "     (Select b.医嘱id 医嘱id, f_List2str(Cast(Collect(b.项目 || ':' || b.内容) As t_Strlist)) 诊断" & vbNewLine & _
                      "       From 检验标本记录 A, 病人医嘱附件 B" & vbNewLine & _
                      "       Where a.医嘱id = b.医嘱id and a.病人id = [1]" & vbNewLine & _
                      "       Group By b.医嘱id) D" & vbNewLine & _
                      "Where a.医嘱id = c.Id(+) And a.医嘱id = d.医嘱id(+) And a.审核人 Is Not Null and a.病人id = [1]" & vbNewLine & _
                      " Order By a.操作类型, a.姓名, a.性别, a.年龄, a.审核时间, a.检验项目"
58            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入病人列表", mlngGetPatientID)
59        Else
60            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入病人列表", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
                                  CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59"), txtPatiNo, mlngGetPatientID, strDepts, strDept, _
                                  strPatients, mlngPatientPage, cbodor.Text, CDate(Format(dtpVS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpVE, "yyyy-MM-dd") & " 23:59:59"))
61        End If

62        Set GetOldLisData = rsTmp


63        Exit Function
GetOldLisData_Error:
64        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetOldLisData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
65        Err.Clear

End Function

Private Sub ReadPatientVal(lngSampleID As Long)
          Dim strSQL As String
1         On Error GoTo ReadPatientVal_Error

2         If mintPatientType = 2 Then
3             strSQL = "Select Distinct a.病人id,b.门诊号,a.住院号,a.入院病床,a.姓名, a.主页id," & _
                       " a.入院日期, a.出院日期  From 病案主页 a,病人信息 B where a.病人id=b.病人id and a.病人id =[1]  order by 主页ID"

4             Set mrsPatientVal = ComOpenSQL(Sel_His_DB, strSQL, "检验技师站", lngSampleID)

5             With Me.cboPages
6                 .Clear
7                 .AddItem "所有"
8                 .ItemData(.NewIndex) = 0
9                 Do Until mrsPatientVal.EOF
10                    .AddItem "第 " & mrsPatientVal("主页ID") & " 次"
11                    .ItemData(.NewIndex) = mrsPatientVal("病人id")
12                    mrsPatientVal.MoveNext
13                Loop
14                If mrsPatientVal.RecordCount > 0 Then
15                    mrsPatientVal.MoveLast
16                    mlngPatientPage = mrsPatientVal("主页ID")
17                    .Text = "第 " & mlngPatientPage & " 次"
18                End If
19                Call readDate(True)
20            End With
21        End If


22        Exit Sub
ReadPatientVal_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(ReadPatientVal)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
24        Err.Clear

End Sub

Private Sub readDate(ByVal blnFilter As Boolean)
1         On Error GoTo readDate_Error

2         If mrsPatientVal.RecordCount > 0 Then
3             mrsPatientVal.MoveFirst
4             If blnFilter = True Then
5                 mrsPatientVal.Filter = "主页ID=" & mlngPatientPage & ""
6                 If mrsPatientVal.RecordCount <> 0 Then
7                     Me.dtpS.value = IIf(IsNull(mrsPatientVal("入院日期")), Currentdate, mrsPatientVal("入院日期"))
8                     Me.dtpE.value = IIf(IsNull(mrsPatientVal("出院日期")), Currentdate, mrsPatientVal("出院日期"))
9                     If lblNo.Caption = "住院号↓" Then
10                        txtPatiNo.Text = mrsPatientVal("住院号") & ""
11                    ElseIf lblNo.Caption = "门诊号" Then
12                        txtPatiNo.Text = mrsPatientVal("门诊号") & ""
13                    Else
14                        txtPatiNo.Text = mrsPatientVal("入院病床") & ""
15                    End If
'16                    cboDept.ListIndex = 0
17                End If
18            Else
19                Me.dtpS.value = IIf(IsNull(mrsPatientVal("入院日期")), Currentdate, mrsPatientVal("入院日期"))
20                mrsPatientVal.MoveLast
21                Me.dtpE.value = IIf(IsNull(mrsPatientVal("出院日期")), Currentdate, mrsPatientVal("出院日期"))
22                txtPatiNo.Text = ""
'23                cboDept.ListIndex = 0
24                txtPatiNo = ""
25            End If
26            mrsPatientVal.Filter = ""
27        End If


28        Exit Sub
readDate_Error:
29        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(readDate)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
30        Err.Clear
End Sub


Private Sub dtpE_Change()
    Me.cboPages.Text = "所有"
End Sub

Private Sub dtpS_Change()
    Me.cboPages.Text = "所有"
End Sub

Private Sub picSupplement_Resize()
    On Error Resume Next
    With vsfSupplement
        .Left = 0
        .Width = Me.picSupplement.Width
        .Height = Me.picSupplement.Height - .Top
        .BorderStyle = flexBorderNone
    End With
End Sub

Private Sub tabPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    RefreshTab Item.Index
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        GetPatientsList
    End If
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub

End Sub

Private Sub txtPatiNo_GotFocus()
    txtPatiNo.SelStart = 0
    txtPatiNo.SelLength = Len(txtPatiNo)
End Sub

Private Sub txtPatiNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If Trim(txtPatiNo) <> "" Then
            ReadPatientList
'        End If
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
3             If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
4                 strSQL = "select * from ( " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')' || decode(g.耐受时间,null,'', '(' || g.耐受时间 || ')')  检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,b.是否审核," & vbNewLine & _
                         "       c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间,B.ID 排序 " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e,耐受试验标本 F,检验耐受时间方案 G" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and d.id =e.组合id and  b.ID=F.报告明细id(+) and F.耐受方案id=G.id(+) AND b.组合id is not null and e.组合id is not null and b.检验结果 is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')' || decode(g.耐受时间,null,'', '(' || g.耐受时间 || ')') 检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,b.是否审核," & vbNewLine & _
                         "       c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间,B.ID 排序 " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e,耐受试验标本 F,检验耐受时间方案 G" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and b.ID=F.报告明细id(+) and F.耐受方案id=G.id(+) AND e.组合id is null and b.组合id is null and b.检验结果 is not null and a.id = [1] ) order by 排序 desc" & vbNewLine
5             Else
6                 strSQL = "select * from ( " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')'  检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,b.是否审核," & vbNewLine & _
                         "       c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间,B.ID 排序 " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and d.id =e.组合id and b.组合id is not null and e.组合id is not null and b.检验结果 is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' 序号,c.id,c.中文名 || '(' || c.英文名 || ')' 检验项目,b.检验结果 结果,b.上次结果 上次," & vbNewLine & _
                         "       c.单位,b.结果参考 参考,a.申请来源 申请类型,e.医嘱id,e.组合id,d.名称 组合名称," & vbNewLine & _
                         "       e.收费状态,e.应收金额,e.实收金额,b.参考高值,b.参考低值,c.排列序号,b.检验结果 日志结果, " & vbNewLine & _
                         "       e.id 申请组合ID,b.结果标志, b.OD, b.CUTOFF, b.SCO,c.结果类型,c.计算公式,b.是否审核," & vbNewLine & _
                         "       c.指标代码,c.临床意义,c.项目类别,nvl(c.小数位数,2) 小数位数,b.上次标志,d.编码 组合编码,a.病人ID,a.核收时间,B.ID 排序 " & vbNewLine & _
                           "from 检验报告记录 a, 检验报告明细 b,检验指标 c,检验组合项目 d,检验申请组合 e" & vbNewLine & _
                           "where a.id = b.标本id and  b.项目id = c.id and  b.组合id = d.id(+) and" & vbNewLine & _
                         "      b.标本id = e.标本id and e.组合id is null and b.组合id is null and b.检验结果 is not null and a.id = [1] ) order by 排序 desc" & vbNewLine
7             End If
8             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", lngSampleID)
9         Else
10            strSQL = "   Select /*+ rule */" & vbNewLine & _
                     "  Distinct '' 序号,a.标本id, a.诊疗项目id, a.编码, a.排列序号, a.固定项目, a.Id, a.检验项目, a.临床意义, a.缩写 As 英文名, a.Cv," & vbNewLine & _
                     " 结果标志 , Decode(a.本次结果, '-', '阴性（-）', '+', '阳性（+）', '*', '*.**', a.本次结果) As 结果, Rownum As 序号, a.标志, a.仪器id, a.标本类别," & vbNewLine & _
                     "   a.核收时间, a.标本序号, a.标本号显示, a.检验备注, a.姓名, a.性别, a.年龄, a.门诊号, a.住院号, a.当前床号, a.主页id, a.结果范围, Nvl(g.小数位数, 2) As 小数," & vbNewLine & _
                     "    a.警戒上限, a.警戒下限, a.单位," & vbNewLine & _
                     "   a.结果参考 As 参考, a.Od, a.Cutoff, a.Cov, a.酶标板id, a.变异报警, a.变异警示, a.结果类型," & vbNewLine & _
                     "   A.结果参考,'' 是否审核,a.诊疗项目" & vbNewLine & _
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
                     " From 检验标本记录 A, 检验标本记录 E, 检验普通结果 B, 诊治所见项目 C, 检验项目 D, 诊疗项目目录 H" & vbNewLine & _
                     " Where a.Id = b.检验标本id And b.检验项目id = c.Id And c.Id = d.诊治项目id And b.诊疗项目id = h.Id(+) And b.记录类型 = a.报告结果 And" & vbNewLine & _
                     "       e.Id = a.合并id And e.id = [1]) A, 检验仪器项目 G" & vbNewLine & _
                     "  Where a.仪器id = g.仪器id(+) And a.Id = g.项目id(+)" & vbNewLine & _
                     "  Order By a.诊疗项目ID,a.排列序号,a.编码"
12            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验技师站", lngSampleID)

13        End If
          '    rsTmp.Sort = "排列序号"
          '    If Not vfgLoadFromRecord(vsfCenter, rsTmp, strErr, imgVsf) Then Exit Sub

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
59            .ColHidden(.ColIndex("是否审核")) = True
60            If vsfLeft.Row > 0 Then
61                If Me.vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("部分审核人")) <> "" Then
62                    For i = 1 To .Rows - 1
63                        If Val(.TextMatrix(i, .ColIndex("是否审核"))) = 1 Then
64                            .Cell(flexcpPicture, i, .ColIndex("结果")) = Me.imgVsf.ListImages("部分审核").ExtractIcon
65                            .Cell(flexcpPictureAlignment, i, .ColIndex("结果")) = flexAlignRightCenter
66                            .RowHidden(i) = False
67                        Else
68                            .RowHidden(i) = True
69                        End If
70                    Next
71                End If
72            End If

73            For i = 1 To .Rows - 1
74                If .Cell(flexcpFontBold, i, .ColIndex("序号")) = True Then
75                    .RowHidden(i) = IIf(Me.chkGroup.value = 1, False, True)
76                End If
77            Next
78        End With

79        CalcReferenceColour

80        Exit Sub
ReadSampleVal_Error:
81        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(ReadSampleVal)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
82        Err.Clear

End Sub

Private Sub ReadSampleBacteriology(lngSampleID As Long, Optional intVal As Integer = 25)
      '功能   读入结果信息
          Dim strErr As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo ReadSampleBacteriology_Error

2         If intVal = 25 Then
3             strSQL = "select b.id,b.中文名 || '(' || b.英文名 || ')' 细菌名,a.检验结果,a.培养描述 描述," & vbNewLine & _
                     "       a.耐药机制,a.组合id," & vbNewLine & _
                       "a.培养时间,a.正常菌,a.未检出,a.补充描述,a.无致病菌,a.无细菌,a.镜检设备,a.镜检检出," & _
                       "a.镜检未检出,a.阳性评语,a.阴性评语,a.结果标志,a.细菌ID,a.是否镜检结果, a.结果性质" & vbNewLine & _
                       "from 检验报告细菌 a,检验细菌记录 b" & vbNewLine & _
                       "where a.细菌id = b.id(+) and a.标本id = [1] "

4             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", lngSampleID)
5         Else
6             strSQL = "SELECT Distinct B.编码, B.ID 细菌id ,D.报告结果,B.中文名 AS 细菌名, " & _
                       "A.检验结果 AS 检验结果,A.培养描述 as 描述,A.耐药机制, d.检验备注,d.备注 " & _
                       "FROM 检验普通结果 A,检验细菌 B,检验标本记录 D  " & _
                       "WHERE A.细菌id = B.ID And D.审核人 is Not null  " & _
                       "AND A.记录类型 = [1]  " & _
                       "AND D.ID=A.检验标本ID  " & _
                       "AND D.ID= [2] Order by B.编码"
7             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "检验技师站", mlngValueC, lngSampleID)
8         End If
          '    rsTmp.Sort = "排列序号"
9         If Not vfgLoadFromRecord(VsfMicrobe, rsTmp, strErr, imgVsf) Then Exit Sub

10        With VsfMicrobe
11            If intVal = 25 Then
12                .ColWidth(.ColIndex("细菌名")) = 3000: .ColHidden(.ColIndex("细菌名")) = False
13                .ColWidth(.ColIndex("检验结果")) = 2000: .ColHidden(.ColIndex("检验结果")) = False
14                .ColWidth(.ColIndex("描述")) = 3000: .ColHidden(.ColIndex("描述")) = False
15                .ColWidth(.ColIndex("耐药机制")) = 3000: .ColHidden(.ColIndex("耐药机制")) = False
16                If rsTmp.RecordCount > 0 Then
17                    rsTmp.MoveFirst

18                    Me.txtNormalMicrobe = rsTmp("正常菌") & ""
19                    Me.txtNoFindMicrobe = rsTmp("未检出") & ""
20                    Me.txtNormalMicrobes = rsTmp("补充描述") & ""
21                    Me.chkPathopoiesiaGerm.value = IIf(rsTmp("无致病菌") = 1, 1, 0)
22                    Me.chkNoGerm.value = IIf(rsTmp("无细菌") = 1, 1, 0)
23                    Me.txtMicroscope = rsTmp("镜检设备") & ""
24                    Me.txtMicroscopeFinded = rsTmp("镜检检出") & ""
25                    Me.txtMicroscopeNOFind = rsTmp("镜检未检出") & ""
26                    Me.txtMicrobePositiveComment = rsTmp("阳性评语") & ""
27                    Me.txtGermComment = rsTmp("阴性评语") & ""
28                    If Val(rsTmp("是否镜检结果") & "") = 0 Then
29                        chkMicroscope.value = 0
30                    Else
31                        chkMicroscope.value = 1
32                    End If
33                    If Val(rsTmp("结果性质") & "") = 0 Then
34                        optReport(1).value = True
35                    Else
36                        optReport(0).value = True
37                    End If
38                    optReportShow

39                    ReadSampleAntibiotic mlngKey, Val(rsTmp("细菌ID") & "")
40                End If
41            Else
42                .ColWidth(.ColIndex("细菌名")) = 3000: .ColHidden(.ColIndex("细菌名")) = False
43                .ColWidth(.ColIndex("检验结果")) = 2000: .ColHidden(.ColIndex("检验结果")) = False
44                .ColWidth(.ColIndex("描述")) = 3000: .ColHidden(.ColIndex("描述")) = False
45                .ColWidth(.ColIndex("耐药机制")) = 3000: .ColHidden(.ColIndex("耐药机制")) = False
46                If rsTmp.RecordCount > 0 Then
47                    rsTmp.MoveFirst
48                    txtMicrobePositiveComment = rsTmp("备注") & ""
49                    txtComment = rsTmp("检验备注") & ""
50                    ReadSampleAntibiotic mlngKey, Val(rsTmp("细菌ID") & ""), 10
51                End If
52            End If
53        End With


54        Exit Sub
ReadSampleBacteriology_Error:
55        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(ReadSampleBacteriology)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
56        Err.Clear

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

Private Sub ReadSampleAntibiotic(lngSampleID As Long, lngBacteriologyID As Long, Optional intVal As Integer = 25)
          '功能           读入抗生素写入VSF
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strErr As String
          Dim i As Integer

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
14            .ColWidth(.ColIndex("药敏方法")) = 1500: .ColHidden(.ColIndex("药敏方法")) = False

15            If .Rows > 1 And Not mrsAntibioticValType Is Nothing And intVal = 25 Then
16                For i = 1 To .Rows - 1
17                    If InStr(.TextMatrix(i, .ColIndex("结果类型")), "-") > 0 Then
                          '药敏结果类型颜色标识
18                        mrsAntibioticValType.Filter = ""
19                        mrsAntibioticValType.Filter = "编码='" & Split(.TextMatrix(i, .ColIndex("结果类型")), "-")(0) & "'"
20                        .Cell(flexcpBackColor, i, .ColIndex("结果类型"), i, .ColIndex("结果类型")) = Val(mrsAntibioticValType("颜色") & "")
21                        mrsAntibioticValType.Filter = ""
22                    Else
23                        .Cell(flexcpBackColor, i, .ColIndex("结果类型"), i, .ColIndex("结果类型")) = 0
24                    End If
25                Next
26            End If
27        End With


28        Exit Sub
ReadSampleAntibiotic_Error:
29        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(ReadSampleAntibiotic)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
30        Err.Clear

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
3             blnTre = IsTre(lngSampleID)

4             If blnTre Then
5                 strSQL = "Select b.id, b.中文名, b.英文名, b.单位, a.id 次数, c.报告时间, a.检验结果, e.耐受时间, b.变异报警率, b.结果类型, a.结果标志" & vbNewLine & _
                           "   From 检验报告明细 A, 检验指标 B, 检验报告记录 C, 耐受试验标本 D, 检验耐受时间方案 E" & vbNewLine & _
                           "   Where A.项目ID = B.ID And A.标本ID = C.ID And A.ID = D.报告明细id And D.耐受方案id = e.ID And A.标本ID = [1]" & vbNewLine & _
                           "   Order By a.id Desc"
6                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入比对数据", lngSampleID)
7             Else
8                 strSQL = "Select " & vbNewLine & _
                           " B.Id, B.中文名, B.英文名, B.单位, A.次数, A.报告时间, A.检验结果, B.变异报警率, B.结果类型, A.结果标志" & vbNewLine & _
                           "From (Select B.项目id 检验项目id, B.次数, B.报告时间, B.检验结果, B.结果标志" & vbNewLine & _
                           "       From (Select A.Id 次数, A.病人id, A.标本类型, A.报告时间, B.项目id" & vbNewLine & _
                           "              From 检验报告记录 A, 检验报告明细 B" & vbNewLine & _
                           "              Where A.Id = B.标本id And A.Id = [1] and b.检验结果 is not null ) A," & vbNewLine & _
                           "            (Select A.Id 次数, A.病人id, A.标本类型, A.报告时间, B.项目id, B.检验结果, B.结果标志" & vbNewLine & _
                           "              From 检验报告记录 A, 检验报告明细 B" & vbNewLine & _
                           "              Where A.Id = B.标本id And A.病人id = [2] And 核收时间 Between [3] And [4] and a.id <= [1] and b.检验结果 is not null ) B" & vbNewLine & _
                           "       Where A.病人id = B.病人id And A.项目id + 0 = B.项目id And Nvl(A.标本类型, 0) = Nvl(B.标本类型, 0) ) A, 检验指标 B" & vbNewLine & _
                           "Where A.检验项目id = B.Id" & vbNewLine & _
                           "Order By LPad(B.排列序号, 10, '0'),b.id, A.次数 desc "

9                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入比对数据", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                         CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
10            End If
11        Else
12            strSQL = "    Select " & vbNewLine & _
                       "       i.Id, i.名称 As 中文名, v.缩写 As 英文名, i.计算单位 As 单位, a.次数, a.报告时间, a.检验结果, v.变异报警率, v.结果类型, a.结果标志" & vbNewLine & _
                       "       From (Select b.检验项目id, b.次数, b.报告时间, b.检验结果, b.结果标志" & vbNewLine & _
                       "              From (Select a.Id 次数, a.病人id, a.标本类型, a.审核时间 报告时间, b.检验项目id, b.检验结果, b.结果标志" & vbNewLine & _
                       "                     From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                       "                     Where a.Id = b.检验标本id And a.Id = [1] And 病人id = [2] And b.检验结果 Is Not Null) A," & vbNewLine & _
                       "                   (Select a.Id 次数, a.病人id, a.标本类型, a.审核时间 报告时间, b.检验项目id, b.检验结果, b.结果标志" & vbNewLine & _
                       "                     From 检验标本记录 A, 检验普通结果 B" & vbNewLine & _
                       "                     Where a.Id = b.检验标本id And a.Id < [1] And 病人id = [2]  And  审核时间 Between [3] And [4]  And b.检验结果 Is Not Null) B" & vbNewLine & _
                       "              Where a.病人id = b.病人id And a.检验项目id + 0 = b.检验项目id) A, 检验项目 V, 检验报告项目 R, 诊疗项目目录 I" & vbNewLine & _
                       "       Where A.检验项目id = v.诊治项目id And A.检验项目id = r.报告项目id And r.诊疗项目id = i.ID And i.组合项目 <> 1"
13            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入比对数据", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                     CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
14        End If
15        vfgSetting 0, VSFList
16        With VSFList
17            .AllowSelection = True
18            .SelectionMode = flexSelectionFree
19            .Rows = 1
20            .Cols = 1
21            .FixedRows = 1
              '        .FixedCols = 1
22            .TextMatrix(0, 0) = "检验项目": .ColWidth(0) = 2500: .RowHeight(0) = 800
23            Do Until rsTmp.EOF
24                If lngItemid <> rsTmp("ID") Then
25                    .Rows = .Rows + 1
26                    intCol = 0
27                    If .Cols - 1 < intCol Then
28                        .Cols = .Cols + 1
29                        .ColWidth(intCol) = 1500
30                    End If

31                    If intCol = 0 Then
                          '写入项目
32                        .TextMatrix(.Rows - 1, intCol) = rsTmp("中文名") & "(" & rsTmp("英文名") & ")"

33                    End If
34                    intCol = intCol + 1
35                    If .Cols - 1 < intCol Then
36                        .Cols = .Cols + 1
37                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter

38                        If blnTre Then
39                            .TextMatrix(0, intCol) = rsTmp("耐受时间") & ""
40                        Else
41                            .TextMatrix(0, intCol) = "本次(" & Mid(Mid(rsTmp("报告时间"), 3), 1, Len(Mid(rsTmp("报告时间"), 3)) - 3) & ")"
42                        End If


43                    End If
                      '写入内容
44                    .TextMatrix(.Rows - 1, intCol) = rsTmp("检验结果") & ""
45                    .Cell(flexcpBackColor, .Rows - 1, intCol) = GetValColour(Val(rsTmp("结果标志" & "")))

46                Else
47                    intCol = intCol + 1
48                    If .Cols - 1 < intCol Then
49                        .Cols = .Cols + 1
50                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
51                        If blnTre Then
52                            .TextMatrix(0, intCol) = rsTmp("耐受时间") & ""
53                        Else
54                            .TextMatrix(0, intCol) = "上" & intCol - 1 & "次(" & Mid(Mid(rsTmp("报告时间"), 3), 1, Len(Mid(rsTmp("报告时间"), 3)) - 3) & ")"
55                        End If
56                        dblTmp = Val(CalcVolatility(.TextMatrix(.Rows - 1, 1), .TextMatrix(.Rows - 1, intCol)))
57                        If dblTmp <> 0 And Val(rsTmp("变异报警率") & "") <> 0 Then
58                            If dblTmp > Val(rsTmp("变异报警率") & "") Then
59                                .Cell(flexcpBackColor, .Rows - 1, intCol) = RGB(248, 194, 169)
60                            End If
61                        End If
62                    End If
                      '写入内容
63                    .TextMatrix(.Rows - 1, intCol) = rsTmp("检验结果") & ""
64                    .Cell(flexcpBackColor, .Rows - 1, intCol) = GetValColour(Val(rsTmp("结果标志" & "")))
65                End If
66                lngItemid = rsTmp("ID")
67                rsTmp.MoveNext
68            Loop
69        End With

70        LoadContrastDBWriteVSF = True


71        Exit Function
LoadContrastDBWriteVSF_Error:
72        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(LoadContrastDBWriteVSF)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
73        Err.Clear

End Function

Private Sub ReadHistorData()
    '功能           读出历次的数据
    Dim strErr As String
    Call LoadContrastDBWriteVSF(VSFContrast, mlngKey, mlngPatientID, mReportDate, 60, strErr)
End Sub


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
56        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(LoadVSFContrastToCht)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
57        Err.Clear

End Function
Private Sub ReadContrastToVsf()
    '功能       读入历次比对到VSF
    Dim strErr As String

    Me.VSFContrast.Rows = 1: Me.VSFContrast.Rows = 2


    '没有病人ID时退出
    If mlngPatientID = 0 Then Exit Sub

    Call LoadContrastDBWriteVSF(Me.VSFContrast, mlngKey, mlngPatientID, mReportDate, Val(txtMaxDay), strErr)
    Call VSFContrast_SelChange
End Sub

Private Sub InitFace()
    '功能           初始化界面
    '========================================显示颜色设置============================================
    '显示颜色设置
    gSampleShowColour.正常 = &H80000005
    gSampleShowColour.偏高 = Val(ComGetPara(Sel_Lis_DB, "显示偏高颜色", 2500, 2500, "8438015"))
    gSampleShowColour.偏低 = Val(ComGetPara(Sel_Lis_DB, "显示偏低颜色", 2500, 2500, "8454143"))
    gSampleShowColour.警示偏高 = Val(ComGetPara(Sel_Lis_DB, "显示警示偏高颜色", 2500, 2500, "255"))
    gSampleShowColour.警示偏低 = Val(ComGetPara(Sel_Lis_DB, "显示警示偏低颜色", 2500, 2500, "255"))
    gSampleShowColour.复查偏高 = Val(ComGetPara(Sel_Lis_DB, "显示复查偏高颜色", 2500, 2500, "65280"))
    gSampleShowColour.复查偏低 = Val(ComGetPara(Sel_Lis_DB, "显示复查偏低颜色", 2500, 2500, "12648384"))
    gSampleShowColour.异常 = Val(ComGetPara(Sel_Lis_DB, "显示异常颜色", 2500, 2500, "16576"))

    picGeneral.Visible = True
    picMicrobePositive.Visible = False
    PicNegative.Visible = False
End Sub

Private Sub CalcReferenceColour()
          '功能           计算结果的颜色
          Dim intCol As Integer
          Dim intRow As Integer
          
1         On Error GoTo CalcReferenceColour_Error
          
2         With vsfCenter
3             For intRow = 1 To .Rows - 1
4                 If .ColIndex("id") <> -1 Then
5                     If Val(.TextMatrix(intRow, .ColIndex("id"))) <> 0 Then
6                         .SelectionMode = flexSelectionFree
7                         If intRow = .Row Then
8                             For intCol = 0 To .Cols - 1
9                                 .Cell(flexcpBackColor, intRow, intCol, intRow, intCol) = &HFFEBD7
10                            Next
11                        End If
                          
12                        .Cell(flexcpBackColor, intRow, .ColIndex("结果"), intRow, .ColIndex("结果")) = GetValColour(Val(.TextMatrix(intRow, .ColIndex("结果标志"))))
13                        If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("版本"))) = 25 Then
14                            .Cell(flexcpBackColor, intRow, .ColIndex("上次"), intRow, .ColIndex("上次")) = GetValColour(Val(.TextMatrix(intRow, .ColIndex("上次标志"))))
15                        End If
16                    End If
17                End If
18            Next
19        End With
          
          
20        Exit Sub
CalcReferenceColour_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(CalcReferenceColour)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
22        Err.Clear
End Sub

Private Sub vsfCenter_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
          Dim intCol As Integer

1         On Error GoTo vsfCenter_AfterRowColChange_Error

2         If OldRow <> NewRow Or OldCol <> NewCol Then
3             With vsfCenter
4                 If .ColIndex("id") <> -1 Then
5                     If Val(.TextMatrix(.Row, .ColIndex("id"))) <> 0 Then
6                         txtSignificance.Text = .TextMatrix(.Row, .ColIndex("临床意义"))
7                     End If
8                 End If
9             End With
10        End If

11        Exit Sub
vsfCenter_AfterRowColChange_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(vsfCenter_AfterRowColChange)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear
End Sub

Private Sub vsfLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim Point As POINTAPI
    Dim strTitle As String

    With Me.vsfLeft
        lngRow = .MouseRow: lngCol = .MouseCol
        If Button = 1 Then
            If lngRow > 0 Then
                If lngCol = .ColIndex("选择") Then
                    If .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 1 And .TextMatrix(lngRow, lngCol) = "" Then
                        .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 2
                    ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                        .Cell(flexcpChecked, lngRow, .ColIndex("选择"), lngRow, .ColIndex("选择")) = 1
                    End If
                End If
            End If
        End If

        If Button = 2 Then
            If lngRow = 0 Then
                Call GetCursorPos(Point)
                strTitle = SetVsfColHiden(Me, Me.vsfLeft, Point.X * 15, Point.Y * 15, "检验报告显示列", 2500, 2001, "结果次数,版本")
                If strTitle <> "" Then
                    SaveDBLog 18, 6, 0, "检验结果浏览", "设置表格列的显示和排序:" & strTitle, 2500, "临床实验室管理"
                    Call ReadPatientList
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
7                         mlngKey = Val(.TextMatrix(.Row, .ColIndex("ID")))
8                         mlngPatientID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
9                         mReportDate = .TextMatrix(.Row, .ColIndex("核收时间"))
10                        txtComment = .TextMatrix(.Row, .ColIndex("备注"))
11                        txtDiagnose = .TextMatrix(.Row, .ColIndex("诊断"))
12                        mlngValueC = .TextMatrix(.Row, .ColIndex("结果次数"))
13                        mintVer = .TextMatrix(.Row, .ColIndex("版本"))
14                        Me.txtResultComment.Text = .TextMatrix(.Row, .ColIndex("结果说明"))
15                        If Val(.TextMatrix(.Row, .ColIndex("微生物"))) = 1 Then
16                            If Val(.TextMatrix(.Row, .ColIndex("阳性报告"))) = 1 Then
17                                picGeneral.Visible = False
18                                picMicrobePositive.Visible = True
19                                PicNegative.Visible = False
20                                If Val(.TextMatrix(.Row, .ColIndex("版本"))) = 25 Then
21                                    ReadSampleBacteriology mlngKey, 25
22                                Else
23                                    ReadSampleBacteriology mlngKey, 10
24                                End If
25                            Else
26                                picGeneral.Visible = False
27                                picMicrobePositive.Visible = False
28                                PicNegative.Visible = True
29                                If Val(.TextMatrix(.Row, .ColIndex("版本"))) = 25 Then
30                                    ReadSampleBacteriology mlngKey
31                                Else

32                                End If
33                            End If
34                        Else
35                            picGeneral.Visible = True
36                            picMicrobePositive.Visible = False
37                            PicNegative.Visible = False
38                            If Val(.TextMatrix(.Row, .ColIndex("版本"))) = 25 Then
39                                ReadSampleVal mlngKey, 25

40                            Else
41                                ReadSampleVal mlngKey, 10
42                            End If
43                        End If

                          '补充报告
44                        If Val(.TextMatrix(.Row, .ColIndex("补充报告"))) = 3 Then
45                            picSupplement.Visible = True
46                            Call GetSupplementReport(mlngKey, vsfSupplement)    '获取补充报告
47                            Call EditSampleValueList(vsfCenter, vsfSupplement)
48                        Else
49                            picSupplement.Visible = False
50                        End If
51                        Call picGeneral_Resize


52                        If .TextMatrix(.Row, .ColIndex("查阅")) <> "1" Then
53                            Call funWriteAdvicesLookState(.TextMatrix(.Row, .ColIndex("申请ID")), 1)
54                            .TextMatrix(.Row, .ColIndex("查阅")) = 1
55                            .Cell(flexcpPicture, .Row, .ColIndex("查阅")) = imgVsf.ListImages("查阅").ExtractIcon
56                        End If
      '                    mlngKey = 0
57                        RefreshTab Me.TabPage.Selected.Index
58                    End If
59                End If
60            End If
61        End With


62        Exit Sub
vsfLeft_SelChange_Error:
63        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(vsfLeft_SelChange)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
64        Err.Clear
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
    On Error GoTo ReadImages_Error

    Call ImageTypeSet(9, True)
    '读入图像数据
    If ReadSampleImage(lngSampleID, strChart, strErr, intVal) = False Then Exit Sub
    For intloop = 0 To 8
        If strChart(intloop) <> "" Then
            chtPic(intloop).Load (strChart(intloop))
        End If
    Next
    '读入完成再排版
    Call ImageTypeSet(9)


    Exit Sub
ReadImages_Error:
    Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(ReadImages)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear

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
16            End If
17        End If
18        PtintOldReport = True


19        Exit Function
PtintOldReport_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(PtintOldReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
21        Err.Clear
End Function




Public Function PrintReport(objFrm As Object, lngSampleID As Long, Optional byRunMode As Byte = 2, Optional intRow As Integer, Optional lngPrintCount As Long, Optional strErr As String) As Boolean
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

2         strSQL = "select b.id 仪器id ,b.名称 仪器名称,b.仪器类别,Nvl(a.病人来源,1) 病人来源,a.报告时间,a.阳性报告,a.标本序号,a.医生站打印 from 检验报告记录 a,检验仪器记录 b where a.仪器id = b.id and a.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "报告打印", lngSampleID)

4         If rsTmp.RecordCount = 0 Then Exit Function

          '对比打印次数和参数
5         If lngPrintCount > 0 Then
6             If Val(rsTmp("医生站打印") & "") >= lngPrintCount And Val(rsTmp("病人来源") & "") = 2 Then
7                 With Me.vsfLeft
8                     .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
9                     .Cell(flexcpPicture, intRow, .ColIndex("打印")) = imgVsf.ListImages("禁止打印").ExtractIcon
10                End With
11                PrintReport = False
12                Exit Function
13            End If
14        End If

15        strSQL = "select id,编码,名称,门诊单据,住院单据,体检单据,院外单据,门诊格式,住院格式,体检格式,院外格式,格式数量," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊单据, '00000')) || '-2' 门诊单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院单据, '00000')) || '-2' 住院单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检单据, '00000')) || '-2' 体检单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外单据, '00000')) || '-2' 院外单据号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(门诊格式, '00000')) || '-2' 门诊格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(住院格式, '00000')) || '-2' 住院格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(体检格式, '00000')) || '-2' 体检格式号," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(院外格式, '00000')) || '-2' 院外格式号" & vbNewLine & _
                      "from 检验仪器记录 where id = [1] "

16        Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "检验技师站", Val(rsTmp("仪器ID") & ""))


17        rsReportFormat.Filter = "id=" & Val(rsTmp("仪器ID") & "")
18        If Val(rsTmp("仪器类别")) = 1 Then
19            If Val(rsTmp("阳性报告") & "") = 1 Then
                  '阳性
20                intSel = 0
21            Else
                  '阴性
22                intSel = 1
23            End If
24        Else
25            intCount = GetSampleValCount(lngSampleID)
              '没有结果时提示
26            If intCount = 0 Then
27                Exit Function
28            End If
29            If rsReportFormat.RecordCount > 0 Then
30                If Val(rsReportFormat("格式数量") & "") > 0 Then
31                    If intCount > Val(rsReportFormat("格式数量") & "") Then
32                        intSel = 0
33                    Else
34                        intSel = 1
35                    End If
36                End If
37            Else
38                intSel = 0
39            End If

40        End If
41        Select Case Val(rsTmp("病人来源"))
              Case 1
42                If intSel = 0 Then
43                    strNO = rsReportFormat("门诊单据号")
44                Else
45                    strNO = rsReportFormat("门诊格式号")
46                End If
47            Case 2
48                If intSel = 0 Then
49                    strNO = rsReportFormat("住院单据号")
50                Else
51                    strNO = rsReportFormat("住院格式号")
52                End If
53            Case 3
54                If intSel = 0 Then
55                    strNO = rsReportFormat("住院单据号")
56                Else
57                    strNO = rsReportFormat("住院格式号")
58                End If
59            Case 4
60                If intSel = 0 Then
61                    strNO = rsReportFormat("院外单据号")
62                Else
63                    strNO = rsReportFormat("院外格式号")
64                End If
65            Case Else
66                If intSel = 0 Then
67                    strNO = rsReportFormat("门诊单据号")
68                Else
69                    strNO = rsReportFormat("门诊格式号")
70                End If
71        End Select
72        If byRunMode = 3 Then
73            If strNO <> "" Then
74                FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
75            End If
76        Else
             '读图像
77            strTmp = "开始读入图像:" & Now & vbCrLf
78            If ReadSampleImage(lngSampleID, strChart, strErr, 25) = False Then
79                Exit Function
80            End If
81            strTmp = strTmp & "读入图像完成:" & Now & vbCrLf

82            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "标本ID=" & lngSampleID, "图形1=" & strChart(0), "图形2=" & strChart(1), "图形3=" & strChart(2), _
                      "图形4=" & strChart(3), "图形5=" & strChart(4), "图形6=" & strChart(5), "图形7=" & strChart(6), "图形8=" & strChart(7), _
                      "图形9=" & strChart(8), byRunMode
83            strTmp = strTmp & "打印完成:" & Now & vbCrLf

              '对于审核过的标本标识
84            strSQL = "Zl_检验报告打印_Edit(1," & lngSampleID & ",1)"
85            Call ComExecuteProc(Sel_Lis_DB, strSQL, "打印标本")
86            strTmp = strTmp & "完成打印:" & Now

87            SaveDBLog 18, 6, lngSampleID, "打印", "报告打印", 2500, "临床实验室管理"
88        End If

89        PrintReport = True

          '发送刷新科内概况已打印标签申请
90        Call SendMessage("RefreshDeptSurvey7")


91        Exit Function
PrintReport_Error:
92        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(PrintReport)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
93        Err.Clear
End Function


Private Sub GetDept(Optional intType As Integer)
          '功能               读入科室或病区
          '参数               intType 0=科室 1=病区
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
1         On Error GoTo GetDept_Error

2         If intType = 0 Then
3             strSQL = "Select b.id,C.编码, C.名称" & vbNewLine & _
                      "From 部门人员 A, 人员表 B, 部门表 C" & vbNewLine & _
                      "Where A.人员id = B.Id And A.部门id = C.Id And (C.撤档时间 Is Null Or C.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) and a.人员id = [1] "
4         Else
5             strSQL = ""
6         End If
7         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入科室", gUserInfo.ID)
      '    With cboPatient
      '        .Clear
      '        .AddItem "所有科室"
      '        Do Until rsTmp.EOF
      '            .AddItem Trim(rsTmp("编码")) & "-" & Trim(rsTmp("名称")) & ""
      '            .ItemData(.NewIndex) = rsTmp("id")
      '            rsTmp.MoveNext
      '        Loop
      '        .ListIndex = 0
      '    End With


8         Exit Sub
GetDept_Error:
9         Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetDept)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
10        Err.Clear
End Sub
Private Function InitDepts(intDeptView As Integer, Optional strErr As String) As Boolean
      '功能：初始化住院临床科室
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          Dim strDeptIDs As String, lngPreDept As Long

1         On Error GoTo InitDepts_Error

2         If cboDept.ListIndex <> -1 Then
3             lngPreDept = cboDept.ItemData(cboDept.ListIndex)
4         End If

5         If intDeptView = 0 Then
              '按科室读取显示
              '包含门急诊观察室的病人还没有上床，不加只显床上有病人的科室的限制
6             If InStr(mstrPrivs, "全院病人") > 0 Then
7                 strDeptIDs = GetUser科室IDs
8                 strSQL = _
                      " Select Distinct A.ID,A.编码,A.名称" & _
                      " From 部门表 A,部门性质说明 B" & _
                      " Where B.部门ID=A.ID And B.工作性质='临床'" & _
                      " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                      " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                      " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)" & _
                      " Order by A.编码"
9             Else
                  '求有权限的科室：本身所在科室+所属病区包含的科室
10                strSQL = _
                      " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
                      " From 部门表 A,部门性质说明 B,部门人员 C" & _
                      " Where B.部门ID=A.ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                      " And (B.服务对象 IN(2,3) Or (B.服务对象=1 And Exists(Select 1 From 床位状况记录 C Where B.部门ID = C.科室ID)))" & _
                      " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                      " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)" & _
                      " And B.工作性质='临床'"
11                strSQL = strSQL & " Union " & _
                      " Select C.ID,C.编码,C.名称,Nvl(A.缺省,0) As 缺省" & _
                      " From 部门人员 A,病区科室对应 B,部门表 C" & _
                      " Where A.部门ID=B.病区ID And B.科室ID=C.ID And A.人员ID=[1]" & _
                      " And Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.病区ID)" & _
                      " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.病区ID)" & _
                      " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL)" & _
                      " And (C.站点='" & gUserInfo.NodeNo & "' Or C.站点 is Null)"
12                If InStr(mstrPrivs, "ICU病人") > 0 Then
13                    strSQL = strSQL & " Union " & _
                          " Select A.ID,A.编码,A.名称,0 As 缺省" & _
                          " From 部门表 A" & _
                          " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                          " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='临床')" & _
                          " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                          " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)"
14                End If
15                strSQL = "Select ID,编码,名称,Max(缺省) As 缺省 From (" & strSQL & ") Group By ID,编码,名称 Order by 编码"
16            End If
17        Else
              '按病区读取显示
18            If InStr(mstrPrivs, "全院病人") > 0 Then
19                strDeptIDs = GetUser病区IDs
20                strSQL = _
                      " Select Distinct A.ID,A.编码,A.名称" & _
                      " From 部门表 A,部门性质说明 B " & _
                      " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                      " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                      " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)" & _
                      " Order by A.编码"
21            Else
                  '求有权病区：直接所在病区+所在科室所属病区
22                strSQL = _
                      " Select A.ID,A.编码,A.名称,Nvl(C.缺省,0) as 缺省" & _
                      " From 部门表 A,部门性质说明 B,部门人员 C" & _
                      " Where A.ID=B.部门ID And A.ID=C.部门ID And C.人员ID=[1]" & _
                      " And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
                      " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                      " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)"
23                strSQL = strSQL & " Union " & _
                      " Select C.ID,C.编码,C.名称,Nvl(A.缺省,0) as 缺省" & _
                      " From 部门人员 A,病区科室对应 B,部门表 C" & _
                      " Where A.部门ID=B.科室ID And B.病区ID=C.ID And A.人员ID=[1]" & _
                      " And Exists(Select 1 From 部门性质说明 Where 工作性质='临床' And 部门ID=B.科室ID)" & _
                      " And Not Exists(Select 1 From 部门性质说明 Where 工作性质='护理' And 部门ID=B.科室ID)" & _
                      " And (C.撤档时间 is NULL or Trunc(C.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                      " And (C.站点='" & gUserInfo.NodeNo & "' Or C.站点 is Null)"
24                If InStr(mstrPrivs, "ICU病人") > 0 Then
25                    strSQL = strSQL & " Union " & _
                          " Select A.ID,A.编码,A.名称,0 As 缺省" & _
                          " From 部门表 A" & _
                          " Where Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='ICU')" & _
                          " And Exists(Select 1 From 部门性质说明 B Where A.ID=B.部门ID And B.工作性质='护理')" & _
                          " And (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
                          " And (A.站点='" & gUserInfo.NodeNo & "' Or A.站点 is Null)"
26                End If
27                strSQL = "Select ID,编码,名称,Max(缺省) as 缺省 From (" & strSQL & ") Group by ID,编码,名称 Order by 编码"
28            End If
29        End If

30        cboDept.Clear
31        If InStr(mstrPrivs, "所有科室") > 0 Then cboDept.AddItem "00-所有科室"
32        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, gUserInfo.ID)

33        For i = 1 To rsTmp.RecordCount
34            cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
35            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
36            rsTmp.MoveNext
37        Next
38        If rsTmp.RecordCount > 0 Then
39            cboDept.ListIndex = 0
40        End If
41        InitDepts = True


42        Exit Function
InitDepts_Error:
43        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(InitDepts)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
44        Err.Clear

End Function
Public Function GetUser科室IDs(Optional ByVal bln病区 As Boolean, Optional strErr As String) As String
      '功能：获取操作员所属的科室(本身所在科室+所属病区包含的科室),可能有多个
      '参数：是否取所属病区下的科室
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser科室IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
          '没有强制限制临床,可能医技科室用
7         If blnNew Then
8             strSQL = "Select 1 as 类别,部门ID From 部门人员 Where 人员ID=[1] Union" & _
                      " Select Distinct 2 as 类别,B.科室ID From 部门人员 A,病区科室对应 B" & _
                      " Where A.部门ID=B.病区ID And A.人员ID=[1]"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", gUserInfo.ID)
10        End If
11        If bln病区 = False Then
12            rsTmp.Filter = "类别 = 1"
13        Else
14            rsTmp.Filter = ""
15        End If

16        For i = 1 To rsTmp.RecordCount
17            If InStr("," & GetUser科室IDs & ",", "," & rsTmp!部门ID & ",") = 0 Then
18                GetUser科室IDs = GetUser科室IDs & "," & rsTmp!部门ID
19            End If
20            rsTmp.MoveNext
21        Next
22        GetUser科室IDs = Mid(GetUser科室IDs, 2)


23        Exit Function
GetUser科室IDs_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetUser科室IDs)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
25        Err.Clear

End Function
Public Function GetUser病区IDs(Optional strErr As String) As String
      '功能：获取操作员所属的病区(直接属于病区或所在科室所属的病区),可能有多个
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser病区IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
7         If blnNew Then
8             strSQL = _
                  "Select Distinct 病区ID From (" & _
                  " Select A.部门ID as 病区ID" & _
                  " From 部门性质说明 A,部门人员 B" & _
                  " Where A.部门ID=B.部门ID And B.人员ID=[1]" & _
                  " And A.服务对象 in(1,2,3) And A.工作性质='护理'" & _
                  " Union" & _
                  " Select A.病区ID From 病区科室对应 A,部门人员 B" & _
                  " Where A.科室ID=B.部门ID And B.人员ID=[1])"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", gUserInfo.ID)
10        ElseIf rsTmp.RecordCount > 0 Then
11            rsTmp.MoveFirst
12        End If
13        For i = 1 To rsTmp.RecordCount
14            GetUser病区IDs = GetUser病区IDs & "," & rsTmp!病区ID
15            rsTmp.MoveNext
16        Next

17        GetUser病区IDs = Mid(GetUser病区IDs, 2)


18        Exit Function
GetUser病区IDs_Error:
19        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetUser病区IDs)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
20        Err.Clear

End Function

Private Function GetDepts(lngID As Long) As String
          '功能           通过病区ID取得病区下所有科室的科室编码，用","分隔
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
1         On Error GoTo GetDepts_Error

2         strSQL = "select b.编码 from 病区科室对应 a,部门表 b where  a.科室id = b.id and a.病区id = [1] "
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngID)
4         Do Until rsTmp.EOF
5             GetDepts = GetDepts & "," & rsTmp("编码")
6             rsTmp.MoveNext
7         Loop
8         If GetDepts <> "" Then
9             GetDepts = Mid(GetDepts, 2)
10        End If


11        Exit Function
GetDepts_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetDepts)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
13        Err.Clear
End Function

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
Public Function ShowMe(objFrm As Object, lngPatientID As Long, strPrivs As String, lngDept As Long, lngDeptDistrict As Long, _
                intPatientType As Integer, Optional lngPatientPage As Long, Optional strErr As String, Optional ByVal blnShowBorder As Boolean, _
                Optional ByRef objOutFrm As Object) As Boolean

    On Error GoTo errH

    mlngGetPatientID = lngPatientID
    mlngPatientPage = lngPatientPage
    mintPatientType = intPatientType
    mblnShowBorder = blnShowBorder

    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 2001)

    mstrPrivs = strPrivs & ";" & mstrPrivs

    Call ShowType(mintPatientType)

    If lngDeptDistrict > 0 Then
        lblDept.Caption = "申请病区↓"
        InitDepts 1
        If cboDept.ListCount > 0 Then
            CboFind cboDept, lngDeptDistrict
        End If
    Else
        lblDept.Caption = "申请科室↓"
        InitDepts 0
        If cboDept.ListCount > 0 Then
            CboFind cboDept, lngDept
        End If
    End If
    If mintPatientType <> 2 Then
        lblPatient.Visible = False
        cboPatients.Visible = False
        getPartDor
    Else
        lblPatient.Visible = True
        cboPatients.Visible = True
        Call GetPatientsList
    End If
    If blnShowBorder Then
        Me.Show , objFrm  '如果不显示窗体的边框，则表示该窗体为嵌入式调用，不是调用show方法
    Else
        Call YSystemMenu(Me.hWnd)
    End If
    Set objOutFrm = Me

    Exit Function
errH:
    strErr = "出错函数(ShowMe),出错信息:" & Err.Number & " " & Err.Description
End Function

Public Function getPartDor()

          Dim rsDeptDor As ADODB.Recordset

1         On Error GoTo getPartDor_Error

2         cbodor.Clear
3         Set rsDeptDor = GetDeptDor(cboDept.ItemData(cboDept.ListIndex))
4         With Me.cbodor
5             .AddItem "所有"
6             .ItemData(.NewIndex) = 0
7             Do Until rsDeptDor.EOF
8                 .AddItem rsDeptDor("姓名") & ""
          '            .ItemData(.NewIndex) = rsTmp("HIS病人ID")
9                 rsDeptDor.MoveNext
10            Loop
11            If .ListCount > 0 Then .ListIndex = 0
12        End With


13        Exit Function
getPartDor_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(getPartDor)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear

End Function


Private Sub BatchPrint(Optional byRunMode As Byte = 2)
          '功能   批量打印
          Dim intRow As Integer
          Dim lngPrintCount As Long   '医生工作站允许打印报告的次数
          Dim blnPrint As Boolean     '是否打印成功
          Dim blnContinue As Boolean  '是否继续打印其他可打印的报告

          '获取参数中的医生站打印次数
1         On Error GoTo BatchPrint_Error

2         lngPrintCount = Val(ComGetPara(Sel_Lis_DB, "医生工作站报告打印次数", 2500, 2500, 1))

3         With vsfLeft
4             For intRow = 1 To .Rows - 1
5                 If .Cell(flexcpChecked, intRow, .ColIndex("选择"), intRow, .ColIndex("选择")) = 1 Then
6                     If Val(.TextMatrix(intRow, .ColIndex("id"))) > 0 Then
7                         If Val(.TextMatrix(intRow, .ColIndex("版本"))) = 25 Then
8                             If .TextMatrix(intRow, .ColIndex("部分审核人")) <> "" And .TextMatrix(intRow, .ColIndex("审核时间")) = "" Then
9                                 If Not blnContinue Then
10                                    If MsgBox("<" & .TextMatrix(intRow, .ColIndex("申请项目")) & ">只审核了部分指标,无法打印,是否继续打印其他可打印的报告?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
11                                        Exit Sub
12                                    Else
13                                        blnContinue = True
14                                    End If
15                                End If
16                            Else
17                                blnPrint = PrintReport(Me, Val(.TextMatrix(intRow, .ColIndex("id"))), byRunMode, intRow, lngPrintCount)
18                            End If
19                        Else
20                            blnPrint = PtintOldReport(Me, Val(.TextMatrix(intRow, .ColIndex("id"))), Val(.TextMatrix(intRow, .ColIndex("病人id"))), byRunMode)
21                        End If
22                        If byRunMode = 2 And blnPrint = True Then
23                            .TextMatrix(intRow, .ColIndex("打印")) = 1
24                            .Cell(flexcpPicture, intRow, .ColIndex("打印")) = imgVsf.ListImages("打印").ExtractIcon
25                        End If
26                    End If
27                End If
28            Next
29        End With


30        Exit Sub
BatchPrint_Error:
31        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(BatchPrint)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
32        Err.Clear
End Sub

Private Sub VsfMicrobe_SelChange()
    With VsfMicrobe
        If .ColIndex("细菌ID") <> -1 Then
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

Private Sub ShowType(intType As Integer)
    '功能       如果inttype <> 2 调整来门诊显示方式
    If intType = 2 Then Exit Sub
    lblDept.Visible = False
    cboDept.Visible = False
    lblPages.Visible = False
    cboPages.Visible = False
    Label1.Left = lblDept.Left
    dtpS.Left = Label1.Left + Label1.Width + 100
    dtpE.Left = dtpS.Left + dtpS.Width + 100
    chkApplyDate.Left = dtpE.Left + dtpE.Width + 50
    Label5.Left = chkApplyDate.Left + chkApplyDate.Width + 100
    dtpVS.Left = Label5.Left + Label5.Width + 100
    dtpVE.Left = dtpVS.Left + dtpVS.Width + 100
    chkVerifyDate.Left = dtpVE.Left + dtpVE.Width + 50

    cbodor.Left = dtpS.Left
    cbodor.Width = dtpS.Width + dtpE.Width + 100
    lblNo.Caption = "门诊号"
    lblNo.Left = Label5.Left
    txtPatiNo.Left = dtpVS.Left
    lblTimeOut.Left = dtpVE.Left
    txtDay.Left = lblTimeOut.Left + 1140
    Line1.X1 = Line1.X1 - 2050
    Line1.X2 = Line1.X2 - 2050
End Sub

Private Function GetDeptDor(lngDeptID) As ADODB.Recordset
          '功能           传入科室或病区返回对应的医生记录集
          '参数
          '               lngDeptID 科室ID或病区ID
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
1         On Error GoTo GetDeptDor_Error

2         strSQL = "Select b.姓名" & vbNewLine & _
                   "From 部门人员 A, 人员表 B, 部门表 C" & vbNewLine & _
                   "Where A.人员id = B.Id And A.部门id = C.Id And (C.撤档时间 Is Null Or C.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) and c.id = [1] "

3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngDeptID)
4         Set GetDeptDor = rsTmp


5         Exit Function
GetDeptDor_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetDeptDor)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear
End Function

Private Function GetDeptPatients(intDeptType As Integer, lngDeptID) As String
          '功能           传入科室或病区返回对应的病人ID串
          '参数           intDeptType = 1 科室 =2 病区
          '               lngDeptID 科室ID或病区ID
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim strPatients As String
1         On Error GoTo GetDeptPatients_Error

2         If txtDay > 99 Then
3             If MsgBox("录入的出院时间已对于99天，查看相关出院病人，可能会耗时过长，" & vbCrLf & "请问是否继续加载?", vbYesNo + vbQuestion + vbDefaultButton2, "中联软件") = vbNo Then
4                 Exit Function
5             End If
6         End If
7         If intDeptType = 1 Then
8             strSQL = "select 病人id from 在院病人 where 科室ID = [1] union all select 病人id from 病案主页 where 入院科室ID = [1] and 出院日期 between sysdate - [3] and sysdate "
9         ElseIf intDeptType = 2 Then
10            strSQL = "select 病人id from 在院病人 where 病区ID = [1]  union all select 病人id from 病案主页 where 入院病区ID = [1] and 出院日期 between sysdate - [3] and sysdate "
11        Else
12            strSQL = "select distinct a.病人id from 在院病人 a,病案主页 b where a.病人id = b.病人id and b.出院日期 is null and  a.病区ID = [1] and b.住院医师 = [2] "
13        End If
14        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngDeptID, gUserInfo.Name, Val(txtDay.Text))
15        Do While Not rsTmp.EOF
16            If strPatients = "" Then
17                strPatients = rsTmp("病人ID")
18            Else
19                strPatients = strPatients & "," & rsTmp("病人ID")
20            End If
21            rsTmp.MoveNext
22        Loop
23        GetDeptPatients = strPatients


24        Exit Function
GetDeptPatients_Error:
25        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetDeptPatients)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
26        Err.Clear
End Function

Private Function GetPatientsList()
          '功能               把病人信息读入到指定的下拉框中
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim intDeptType As Integer
          Dim strPatients As String
          Dim rsDeptDor As ADODB.Recordset

          Dim strPatientsCut As String
          Dim strLast As String

1         On Error GoTo GetPatientsList_Error

2         strSQL = "Select a.id,c.分类,0 选择,A.姓名, Decode(A.性别, 1, '男', 2, '女', 9, '未知', '') 性别, A.年龄, C.名称 申请项目, " & _
                   " A.住院号, A.床号,B.申请时间,a.病人ID,a.核收时间,a.审核时间,a.备注,a.诊断,a.微生物,a.阳性报告,a.医生站打印 " & vbNewLine & _
                  " From 检验报告记录 A, 检验申请组合 B, 检验组合项目 C" & vbNewLine & _
                  " Where A.Id = B.标本id And B.组合id = C.Id(+) and (a.审核人 is not null or a.部分审核人 is not null) "


3         strSQL = "select /*+ rule */ distinct a.HIS病人ID,a.姓名 from 检验报告记录 a where "

4         If mintPatientType = 2 Then

5             If InStr(mstrPrivs, "本科病人") > 0 Then
6                 If cboDept <> "" Then
7                     If lblDept.Caption = "申请病区↓" Then
8                         intDeptType = 2
9                     Else
10                        intDeptType = 1
11                    End If
12                    strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
13                    strSQL = strSQL & "  a.his病人id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
14                Else
15                    If lblDept.Caption = "申请病区↓" Then
16                        intDeptType = 2
17                    Else
18                        intDeptType = 1
19                    End If
20                    strPatients = GetDeptPatients(intDeptType, gUserInfo.DeptID)
21                    strSQL = strSQL & "  a.his病人id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
22                End If
23            Else
24                intDeptType = 3
25                If cboDept <> "" Then
26                    strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
27                    strSQL = strSQL & "  a.his病人id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
28                Else
29                    strPatients = GetDeptPatients(intDeptType, gUserInfo.DeptID)
30                    strSQL = strSQL & "  a.his病人id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
31                End If
32            End If

33            If strPatients <> "" Then
34                With Me.cboPatients
35                    .Clear
36                    .AddItem "所有"
37                    .ItemData(.NewIndex) = 0
                      '当strPatients长度大于3500时,需要分解
38                    Do While Len(strPatients) >= 3500
39                        strPatientsCut = Mid(strPatients, 1, 3500)
40                        strPatients = Mid(strPatients, 3501)
41                        strLast = Mid(strPatientsCut, InStrRev(strPatientsCut, ","))
42                        strPatientsCut = Mid(strPatientsCut, 1, InStrRev(strPatientsCut, ",") - 1)
43                        strPatients = strLast & strPatients

44                        If Mid(strPatientsCut, 1, 1) = "," Then
45                            strPatientsCut = Mid(strPatientsCut, 2)
46                        End If
47                        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入病人列表", strPatientsCut)
48                        Do Until rsTmp.EOF
49                            .AddItem rsTmp("姓名") & ""
50                            .ItemData(.NewIndex) = rsTmp("HIS病人ID")
51                            rsTmp.MoveNext
52                        Loop
53                    Loop
54                    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入病人列表", strPatients)
55                    Do Until rsTmp.EOF
56                        .AddItem rsTmp("姓名") & ""
57                        .ItemData(.NewIndex) = rsTmp("HIS病人ID")
58                        rsTmp.MoveNext
59                    Loop
60                End With
61            End If
62        End If
63        If cboDept <> "" Then
64            Set rsDeptDor = GetDeptDor(cboDept.ItemData(cboDept.ListIndex))
65        End If
66        cbodor.Clear
67        If Not rsDeptDor Is Nothing Then
68            With Me.cbodor
69                .AddItem "所有"
70                .ItemData(.NewIndex) = 0
71                Do Until rsDeptDor.EOF
72                    .AddItem rsDeptDor("姓名") & ""
          '            .ItemData(.NewIndex) = rsTmp("HIS病人ID")
73                    rsDeptDor.MoveNext
74                Loop
75                If .ListCount > 0 Then .ListIndex = 0
76            End With
77        End If
78        strSQL = "select 姓名 from 病人信息 where 病人ID = [1]  "
79        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "读入病人列表", mlngGetPatientID)
80        If rsTmp.EOF = False Then
81            If CheckCboID(Me.cboPatients, mlngGetPatientID) = False Then
82                Me.cboPatients.AddItem rsTmp("姓名") & ""
83                Me.cboPatients.ItemData(Me.cboPatients.NewIndex) = mlngGetPatientID
84                If cboPatients.ListCount = 1 Then
85                    cboPatients.ListIndex = 0
86                End If
87            End If
      '        Me.cboPatients.Text = rsTmp("姓名") & ""
88        End If


89        Exit Function
GetPatientsList_Error:
90        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetPatientsList)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
91        Err.Clear

End Function

Private Function CheckCboID(cboobj As ComboBox, lngID As Long) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能                           检查CBO控件中ID是否重复
    '参数
    '                               cboobj = cbo对象
    '                               lngID = 需要检查的ID
    '返回                           true = 重复 false= 不重复
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intloop As Long

    For intloop = 0 To cboobj.ListCount - 1
        If cboobj.ItemData(intloop) = lngID Then
            cboobj.ListIndex = intloop
            CheckCboID = True
            Exit Function
        End If
    Next
    CheckCboID = False
End Function

Private Function GetDictType(strType As String) As ADODB.Recordset
          '功能   从字典表提取指定的分类
          Dim strSQL As String

1     On Error GoTo GetDictType_Error

2         strSQL = "Select 小组, 编码, 名称, 简码, 内容, 备注, 颜色 From 检验字典表 Where 分类 = [1]"
3         Set GetDictType = ComOpenSQL(Sel_Lis_DB, strSQL, "检验字典", strType)


4         Exit Function
GetDictType_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(GetDictType)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
6         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2018/5/25
'功    能:调用API动态设置窗体的border
'入    参:
'           new_Hwnd    窗体的句柄
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2019-04-19
'功    能:  显示诊疗参考
'入    参:
'出    参:
'返    回:
'调整影响:
'---------------------------------------------------------------------------------------
Private Sub ShowClincHelp()
          Dim lngSampleID As Long
          Dim lngVer As Long

1         On Error GoTo ShowClincHelp_Error

2         With Me.vsfLeft
3             If .Row < 1 Then
4                 MsgBox "请选中一份报告", vbInformation, gSysInfo.AppName
5                 Exit Sub
6             End If
7             If Val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
8                 MsgBox "请选中一份报告", vbInformation, gSysInfo.AppName
9                 Exit Sub
10            End If
11            lngSampleID = Val(.TextMatrix(.Row, .ColIndex("ID")))
12            lngVer = Val(.TextMatrix(.Row, .ColIndex("版本")))
13        End With

14        Call funShowClincHelp(Me, lngSampleID, lngVer)

15        Exit Sub
ShowClincHelp_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "执行(ShowClincHelp)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
17        Err.Clear
End Sub












