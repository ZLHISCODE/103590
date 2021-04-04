VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人信息"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   Icon            =   "frmDegreeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin XtremeSuiteControls.TabControl TabSub 
      Height          =   5985
      Left            =   75
      TabIndex        =   140
      Top             =   45
      Width           =   10680
      _Version        =   589884
      _ExtentX        =   18838
      _ExtentY        =   10557
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   5130
      Left            =   6960
      TabIndex        =   85
      Top             =   60
      Width           =   10275
      _cx             =   18124
      _cy             =   9049
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
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
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
      AutoSizeMode    =   1
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
   Begin VB.PictureBox PicInInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Left            =   2640
      ScaleHeight     =   5475
      ScaleWidth      =   10425
      TabIndex        =   86
      Top             =   915
      Width           =   10425
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   45
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1965
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   37
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   750
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   36
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   750
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   40
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   750
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   58
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   4800
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   35
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   330
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   44
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   43
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   42
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   1560
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   49
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   2325
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   48
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   103
         Top             =   2325
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   55
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   102
         Top             =   4800
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   47
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   2310
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   56
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   100
         Top             =   4800
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   39
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   1155
         Width           =   3690
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   2955
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   98
         Top             =   330
         Width           =   1935
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   57
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1950
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   61
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   1950
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   62
         Left            =   8910
         Locked          =   -1  'True
         TabIndex        =   95
         Top             =   2310
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   38
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   94
         Top             =   1155
         Width           =   3675
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   51
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   3180
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   50
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   92
         Top             =   2790
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   54
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   91
         Top             =   4350
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   52
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   3570
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   46
         Left            =   8910
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   1965
         Width           =   1245
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   53
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   3960
         Width           =   8955
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   66
         Left            =   8910
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   750
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "护理等级"
         Height          =   180
         Left            =   8085
         TabIndex        =   139
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床位等级"
         Height          =   180
         Left            =   5670
         TabIndex        =   138
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院号"
         Height          =   180
         Left            =   615
         TabIndex        =   137
         Top             =   810
         Width           =   540
      End
      Begin VB.Label lbl费别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   6030
         TabIndex        =   136
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院次数"
         Height          =   180
         Left            =   435
         TabIndex        =   135
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院天数"
         Height          =   180
         Left            =   5670
         TabIndex        =   134
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "转科信息"
         Height          =   180
         Left            =   435
         TabIndex        =   133
         Top             =   3630
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院病况"
         Height          =   180
         Left            =   5670
         TabIndex        =   132
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院科室"
         Height          =   180
         Left            =   2835
         TabIndex        =   131
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
         Height          =   180
         Left            =   5850
         TabIndex        =   130
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "责任护士"
         Height          =   180
         Left            =   2835
         TabIndex        =   129
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院医师"
         Height          =   180
         Left            =   435
         TabIndex        =   128
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院时间"
         Height          =   180
         Left            =   435
         TabIndex        =   127
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院时间"
         Height          =   180
         Left            =   435
         TabIndex        =   126
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院科室"
         Height          =   180
         Left            =   2835
         TabIndex        =   125
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "床号"
         Height          =   180
         Left            =   3195
         TabIndex        =   124
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "住院目的"
         Height          =   180
         Left            =   435
         TabIndex        =   123
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院中医诊断"
         Height          =   180
         Left            =   75
         TabIndex        =   122
         Top             =   4410
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院诊断"
         Height          =   180
         Left            =   435
         TabIndex        =   121
         Top             =   2850
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入院中医诊断"
         Height          =   180
         Left            =   75
         TabIndex        =   120
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Left            =   6030
         TabIndex        =   119
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主任医师"
         Height          =   180
         Left            =   2835
         TabIndex        =   118
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "主治医师"
         Height          =   180
         Left            =   435
         TabIndex        =   117
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前病况"
         Height          =   180
         Left            =   8115
         TabIndex        =   116
         Top             =   2370
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出院诊断"
         Height          =   180
         Left            =   435
         TabIndex        =   115
         Top             =   4020
         Width           =   720
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人类型"
         Height          =   180
         Left            =   8085
         TabIndex        =   114
         Top             =   810
         Width           =   720
      End
   End
   Begin VB.PictureBox PicBaseInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5475
      Left            =   1050
      ScaleHeight     =   5475
      ScaleWidth      =   10425
      TabIndex        =   0
      Top             =   120
      Width           =   10425
      Begin VB.CheckBox chk担保 
         BackColor       =   &H8000000A&
         Caption         =   "临时"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9480
         MaskColor       =   &H00000000&
         TabIndex        =   43
         Top             =   4905
         Width           =   735
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         Left            =   4335
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   480
         Width           =   580
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1185
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   480
         Width           =   680
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   135
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   135
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   840
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   23
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   4365
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   15
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1890
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   14
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   13
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   1530
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1185
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1185
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   19
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2595
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   20
         Left            =   3645
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   25
         Top             =   2955
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   22
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3300
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   25
         Left            =   975
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   23
         Top             =   4005
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   27
         Left            =   6180
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   22
         Top             =   2580
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   28
         Left            =   8850
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2580
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   21
         Left            =   975
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         Top             =   2955
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   17
         Left            =   6180
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   19
         Top             =   1530
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   32
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   4935
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   31
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   4935
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   34
         Left            =   8160
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   16
         Top             =   4935
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   33
         Left            =   6180
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   15
         Top             =   4935
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   59
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4005
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   41
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3660
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   64
         Left            =   8850
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   63
         Left            =   8850
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   135
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   60
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4365
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   30
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   3300
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   29
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2955
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   26
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2235
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   24
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3660
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   18
         Left            =   975
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2235
         Width           =   3945
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   12
         Left            =   8850
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1185
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   16
         Left            =   3645
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1890
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   8850
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   495
         Width           =   1275
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   65
         Left            =   6180
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1890
         Width           =   3945
      End
      Begin VB.Label lbl区域 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "区域"
         Height          =   180
         Left            =   540
         TabIndex        =   84
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "保险类别"
         Height          =   180
         Left            =   5415
         TabIndex        =   83
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   360
         TabIndex        =   82
         Top             =   195
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊费别"
         Height          =   180
         Left            =   5415
         TabIndex        =   81
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         Height          =   180
         Left            =   3240
         TabIndex        =   80
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         Height          =   180
         Left            =   540
         TabIndex        =   79
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   3240
         TabIndex        =   78
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         Height          =   180
         Left            =   5595
         TabIndex        =   77
         Top             =   195
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         Height          =   180
         Left            =   360
         TabIndex        =   76
         Top             =   900
         Width           =   540
      End
      Begin VB.Label lbl医疗付款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "医疗付款方式"
         Height          =   180
         Left            =   7725
         TabIndex        =   75
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         Height          =   180
         Left            =   2880
         TabIndex        =   74
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lbl出生地点 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生地点"
         Height          =   180
         Left            =   180
         TabIndex        =   73
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   5415
         TabIndex        =   72
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label lbl身份 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份"
         Height          =   180
         Left            =   540
         TabIndex        =   71
         Top             =   1950
         Width           =   360
      End
      Begin VB.Label lbl职业 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "职业"
         Height          =   180
         Left            =   3240
         TabIndex        =   70
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         Height          =   180
         Left            =   5775
         TabIndex        =   69
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label lbl国籍 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "国籍"
         Height          =   180
         Left            =   3240
         TabIndex        =   68
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label lbl学历 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "学历"
         Height          =   180
         Left            =   8445
         TabIndex        =   67
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label lvl婚姻状况 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婚姻状况"
         Height          =   180
         Left            =   180
         TabIndex        =   66
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label lbl家庭地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭地址"
         Height          =   180
         Left            =   180
         TabIndex        =   65
         Top             =   2655
         Width           =   720
      End
      Begin VB.Label lbl家庭电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "家庭电话"
         Height          =   180
         Left            =   180
         TabIndex        =   64
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lbl户口邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "户口邮编"
         Height          =   180
         Left            =   2880
         TabIndex        =   63
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lbl联系人姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人姓名"
         Height          =   180
         Left            =   0
         TabIndex        =   62
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label lbl联系人关系 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人关系"
         Height          =   180
         Left            =   0
         TabIndex        =   61
         Top             =   4425
         Width           =   900
      End
      Begin VB.Label lbl联系人地址 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人地址"
         Height          =   180
         Left            =   0
         TabIndex        =   60
         Top             =   3720
         Width           =   900
      End
      Begin VB.Label lbl联系人电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "联系人电话"
         Height          =   180
         Left            =   0
         TabIndex        =   59
         Top             =   4065
         Width           =   900
      End
      Begin VB.Label lbl工作单位 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "工作单位"
         Height          =   180
         Left            =   5415
         TabIndex        =   58
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lbl单位电话 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位电话"
         Height          =   180
         Left            =   5415
         TabIndex        =   57
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label lbl单位邮编 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位邮编"
         Height          =   180
         Left            =   8085
         TabIndex        =   56
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label lbl单位开户行 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位开户行"
         Height          =   180
         Left            =   5235
         TabIndex        =   55
         Top             =   3015
         Width           =   900
      End
      Begin VB.Label lbl单位帐号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单位帐号"
         Height          =   180
         Left            =   5415
         TabIndex        =   54
         Top             =   3360
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未结费用"
         Height          =   180
         Left            =   2880
         TabIndex        =   53
         Top             =   4995
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "预交余额"
         Height          =   180
         Left            =   180
         TabIndex        =   52
         Top             =   4995
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保人"
         Height          =   180
         Index           =   1
         Left            =   5595
         TabIndex        =   51
         Top             =   4995
         Width           =   540
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊诊断"
         Height          =   180
         Left            =   5415
         TabIndex        =   50
         Top             =   4065
         Width           =   720
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊中医诊断"
         Height          =   180
         Left            =   5055
         TabIndex        =   49
         Top             =   4425
         Width           =   1080
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊医师"
         Height          =   180
         Left            =   5415
         TabIndex        =   48
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "担保额"
         Height          =   180
         Left            =   7545
         TabIndex        =   47
         Top             =   4995
         Width           =   540
      End
      Begin VB.Label lbl登记时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记时间"
         Height          =   180
         Left            =   8085
         TabIndex        =   46
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "就诊卡号"
         Height          =   180
         Left            =   8085
         TabIndex        =   45
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lbl其他证件 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "其他证件"
         Height          =   180
         Left            =   5400
         TabIndex        =   44
         Top             =   1950
         Width           =   720
      End
   End
   Begin VB.PictureBox PicNoRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   105
      ScaleHeight     =   5565
      ScaleWidth      =   10530
      TabIndex        =   141
      Top             =   105
      Width           =   10530
      Begin VB.Label Label38 
         Caption         =   "不能正确读取病人信息  请与系统管理员联系！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   48
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2070
         Left            =   45
         TabIndex        =   142
         Top             =   1710
         Width           =   10680
      End
   End
End
Attribute VB_Name = "frmDegreeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mlng病人ID As Long '要查看的病人ID
Private mlng主页ID As Long '住院病人时传入主页ID

Private Enum txtName
    '要求和SQL的字段对应
    病人ID = 0
    姓名 = 1
    性别 = 2
    年龄 = 3
    门诊号 = 4
    费别 = 5
    医疗付款方式 = 6
    医保号 = 7
    险类 = 8
    区域 = 9
    国籍 = 10
    民族 = 11
    学历 = 12
    婚姻状况 = 13
    职业 = 14
    身份 = 15
    出生日期 = 16
    身份证号 = 17
    出生地点 = 18
    家庭地址 = 19
    家庭地址邮编 = 20
    家庭电话 = 21
    联系人姓名 = 22
    联系人关系 = 23
    联系人地址 = 24
    联系人电话 = 25
    工作单位 = 26
    单位电话 = 27
    单位邮编 = 28
    单位开户行 = 29
    单位帐号 = 30
    预交余额 = 31
    费用余额 = 32
    担保人 = 33
    担保额 = 34
    
    住院次数 = 35
    住院号 = 36
    出院病床 = 37
    备注 = 38
    住院目的 = 39
    住院费别 = 40
    门诊医师 = 41
    住院医师 = 42
    主治医师 = 57
    主任医师 = 61
    责任护士 = 43
    登记人 = 44
    床位等级 = 45
    护理等级 = 46
    入院日期 = 47
    入院科室 = 48
    入院病况 = 49
    当前病况 = 62
    转科信息 = 52
    出院日期 = 55
    出院科室 = 56
    住院天数 = 58
    
    门诊诊断 = 59
    门诊中医诊断 = 60
    入院诊断 = 50
    入院中医诊断 = 51
    出院诊断 = 53
    出院中医诊断 = 54
    
End Enum

Private Enum cboName
    主页ID = 0
    年龄单位 = 1
End Enum
Public Sub ShowMe(ByVal lng病人ID As Long, ByVal lng主页ID As Long)
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    Me.Show 1
End Sub
Private Function ReadCard(ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal bln查看某次住院 As Boolean) As Boolean
'功能：读取指定病人信息,并显示在界面上
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTxt As String, strTmp As String, strHead As String
    Dim i As Integer, j As Integer, arrTxt As Variant
    
    On Error GoTo errH
    
    strSQL = "Select a.病人id, a.姓名, a.性别, a.年龄, a.门诊号, a.费别, a.医疗付款方式, a.险类, a.区域, a.国籍, a.民族, a.学历," & vbNewLine & _
            "            a.婚姻状况, a.职业, a.身份, Decode(To_Date(To_Char(出生日期, 'YYYY-MM-DD HH24:MI'), 'YYYY-MM-DD HH24:MI') - Trunc(出生日期), 0, To_Char(出生日期, 'YYYY-MM-DD'),To_char(出生日期,'YYYY-MM-DD HH24:MI')) 出生日期, " & _
            "            a.身份证号, a.出生地点, a.家庭地址, a.家庭电话, a.家庭地址邮编, a.联系人姓名," & vbNewLine & _
            "            a.联系人关系, a.联系人地址, a.联系人电话, a.工作单位, a.单位电话, a.单位邮编, a.单位开户行, a.单位帐号, a.担保人," & vbNewLine & _
            "            a.担保额, a.担保性质, a.住院次数, a.住院号, To_char(a.登记时间,'yyyy-mm-dd hh24:mi:ss') As 登记时间, a.就诊卡号, b.出院病床, b.备注, b.住院目的, b.门诊医师, b.住院医师," & vbNewLine & _
            "            b.责任护士, b.登记人, b.入院病况, b.当前病况, b.住院天数, b.费别 As 住院费别, c.预交余额, c.费用余额," & vbNewLine & _
            "            Nvl(A.医保号,d.信息值) 医保号, e.名称 As 护理等级, g.名称 As 床位等级, m.名称 As 入院科室, n.名称 As 出院科室," & vbNewLine & _
            "            To_char(b.入院日期,'yyyy-mm-dd hh24:mi:ss') 入院日期, To_char(b.出院日期,'yyyy-mm-dd hh24:mi:ss') 出院日期,A.其他证件,Nvl(B.病人类型,Decode(B.险类,Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
            "From 病人信息 a, 病案主页 b, 病人余额 c, 病案主页从表 d, 收费项目目录 e, 床位状况记录 f, 收费项目目录 g, 部门表 m, 部门表 n" & vbNewLine & _
            "Where a.病人id = b.病人id(+) And " & IIf(lng主页ID = 0, "Nvl(a.住院次数,0)", "[2]") & "=b.主页ID(+) And a.病人ID=[1] And a.病人id = c.病人id(+) And" & vbNewLine & _
            "           c.性质(+) = 1 And b.病人id = d.病人id(+) And" & vbNewLine & _
            "           b.主页id = d.主页id(+) And d.信息名(+) = '医保号' And b.入院科室id = m.Id(+) And b.出院科室id = n.Id(+) And" & vbNewLine & _
            "           b.护理等级id = e.ID(+) And b.当前病区id = f.病区id(+) And b.出院病床 = f.床号(+) And f.等级id = g.ID(+)"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
    If rsTmp.EOF Then Exit Function
        
    If bln查看某次住院 Then
       strTxt = "出院病床=37,备注=38,住院目的=39,住院费别=40,门诊医师=41,住院医师=42,责任护士=43,登记人=44,床位等级=45,护理等级=46,入院日期=47,入院科室=48," & _
                " 入院病况=49,当前病况=62,转科信息=52,出院日期=55,出院科室=56,住院天数=58,门诊诊断=59,门诊中医诊断=60,入院诊断=50,入院中医诊断=51,出院诊断=53,出院中医诊断=54,病人类型=66"
    Else
        strTxt = "病人ID=0,姓名=1,性别=2,年龄=3,门诊号=4,费别=5,医疗付款方式=6,医保号=7,险类=8,区域=9,国籍=10,民族=11,学历=12,婚姻状况=13,职业=14," & _
                " 身份=15,出生日期=16,身份证号=17,出生地点=18,家庭地址=19,家庭地址邮编=20,家庭电话=21,联系人姓名=22,联系人关系=23,联系人地址=24,联系人电话=25," & _
                " 工作单位=26,单位电话=27,单位邮编=28,单位开户行=29,单位帐号=30,预交余额=31,费用余额=32,担保人=33,担保额=34,住院次数=35,住院号=36," & _
                " 出院病床=37,备注=38,住院目的=39,住院费别=40,门诊医师=41,住院医师=42,责任护士=43,登记人=44,床位等级=45,护理等级=46,入院日期=47,入院科室=48," & _
                " 入院病况=49,当前病况=62,转科信息=52,出院日期=55,出院科室=56,住院天数=58,门诊诊断=59,门诊中医诊断=60,入院诊断=50,入院中医诊断=51,出院诊断=53," & _
                " 出院中医诊断=54,登记时间=63,就诊卡号=64,其他证件=65,病人类型=66"
    End If
    
    arrTxt = Split(strTxt, ",")
    
    For i = 0 To UBound(arrTxt)
        strTmp = Trim(arrTxt(i))
        
        If strTmp <> "" Then
            '排开暂不处理的字段
            If InStr(1, ",门诊诊断,门诊中医诊断,入院诊断,入院中医诊断,出院诊断,出院中医诊断,转科信息,", "," & Trim(Split(strTmp, "=")(0)) & ",") = 0 Then
                If InStr(1, ",费用余额,预交余额,", "," & Trim(Split(strTmp, "=")(0)) & ",") > 0 Then
                    txt(Trim(Split(strTmp, "=")(1))).Text = Format(Val("" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))), "0.00")
                Else
                    txt(Trim(Split(strTmp, "=")(1))).Text = "" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))
                End If
            End If
        End If
    Next
    
    '其它专门处理
    '----------------------------------------------
    Call LoadOldData("" & rsTmp!年龄, txt(txtName.年龄), cbo(cboName.年龄单位))
    If cbo(cboName.年龄单位).ListIndex = -1 Then txt(txtName.年龄).Width = txt(txtName.年龄).Width + cbo(cboName.年龄单位).Width
    chk担保.Value = Val("" & rsTmp!担保性质)
    
    
    '住院信息
    '----------------------------------------------
    If Not bln查看某次住院 Then lng主页ID = Val(txt(txtName.住院次数).Text)
    '住院病人的诊断情况
    If lng主页ID > 0 Then
        strSQL = "Select 诊断类型,疾病ID,诊断描述 From 病人诊断记录 Where 诊断次序=1 And 记录来源=2 And 病人ID=[1] And 主页ID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Select Case rsTmp!诊断类型
                    Case 1
                        j = txtName.门诊诊断
                    Case 11
                        j = txtName.门诊中医诊断
                    Case 2
                        j = txtName.入院诊断
                    Case 12
                        j = txtName.入院中医诊断
                    Case 3
                        j = txtName.出院诊断
                    Case 13
                        j = txtName.出院中医诊断
                    Case Else
                        j = 0
                End Select
                If j <> 0 Then txt(j).Text = IIf(IsNull(rsTmp!疾病id), "", "(" & rsTmp!疾病id & ")") & rsTmp!诊断描述
                
                rsTmp.MoveNext
            Next
        Else
            txt(txtName.门诊诊断).Text = ""
            txt(txtName.门诊中医诊断).Text = ""
            txt(txtName.入院诊断).Text = ""
            txt(txtName.入院中医诊断).Text = ""
            txt(txtName.出院诊断).Text = ""
            txt(txtName.出院中医诊断).Text = ""
        End If
        
        '转科信息
        txt(txtName.转科信息).Text = ""
        strSQL = _
            " Select Distinct 1 as 开始原因,To_Date('1900-01-01','YYYY-MM-DD') as 开始时间,B.名称" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.科室ID=B.ID And A.开始时间 is Not NULL And A.开始原因 IN(1,2)" & _
            " And A.病人ID=[1] And 主页ID=[2]" & _
            " Union ALL " & _
            " Select A.开始原因,A.开始时间,B.名称" & _
            " From 病人变动记录 A,部门表 B" & _
            " Where A.科室ID=B.ID And A.开始时间 is Not NULL And A.开始原因=3" & _
            " And A.病人ID=[1] And 主页ID=[2]" & _
            " Order by 开始时间"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        rsTmp.Filter = "开始原因=3"
        If Not rsTmp.EOF Then
            rsTmp.Filter = 0
            Do While Not rsTmp.EOF
                txt(txtName.转科信息).Text = txt(txtName.转科信息).Text & " ─→ " & rsTmp!名称
                rsTmp.MoveNext
            Loop
            txt(txtName.转科信息).Text = Mid(txt(txtName.转科信息).Text, 5)
        End If
        
        '病案主页从表
        txt(txtName.主治医师).Text = ""
        txt(txtName.主任医师).Text = ""
        strSQL = " Select 信息名,信息值 From 病案主页从表 Where (信息名='主治医师' Or 信息名='主任医师') And 病人ID=[1] And 主页ID=[2]"
        rsTmp.Filter = ""
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID)
        rsTmp.Filter = "信息名='主治医师'"
        If Not rsTmp.EOF Then txt(txtName.主治医师).Text = "" & rsTmp!信息值
        rsTmp.Filter = "信息名='主任医师'"
        If Not rsTmp.EOF Then txt(txtName.主任医师).Text = "" & rsTmp!信息值
    End If
    
    
    '3.病人合并信息
    If Not bln查看某次住院 Then
        
        strSQL = "Select 原信息,合并原因,操作员姓名,合并时间 From 病人合并记录 Where 病人ID=[1] Order by 合并时间 Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID)
                
        strHead = "合并时间,1,1800|操作员,4,800|合并原因,1,1800|" & _
                "病人ID,1,800|门诊号,1,900|住院号,1,800|就诊卡号,1,900|姓名,4,800|" & _
                "性别,4,500|年龄,4,800|出生日期,1,1000|身份证号,1,1800|婚姻状况,4,900|职业,1,1000|家庭地址,1,4200"
        With vsList
            .Redraw = False
            .Rows = rsTmp.RecordCount + 1
            
            .Cols = UBound(Split(strHead, "|")) + 1
            For i = 0 To UBound(Split(strHead, "|"))
                .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
                .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
                .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            Next
            
            .RowHeight(0) = 320
            .Col = 0: .ColSel = .Cols - 1
            
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = Format(rsTmp!合并时间, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 1) = "" & rsTmp!操作员姓名
                .TextMatrix(i, 2) = "" & rsTmp!合并原因
                
                'v_原信息:=r_InfoA.病人Id || ',' || r_InfoA.门诊号 || ',' ||  r_InfoA.住院号 || ',' ||  r_InfoA.就诊卡号 || ',' ||  r_InfoA.姓名 ||  ',' ||  r_InfoA.性别 ||  ',' ||
                '   r_InfoA.年龄 ||  ',' || to_char(r_InfoA.出生日期,'yyyy-mm-dd') ||  ',' || r_InfoA.身份证号 ||  ',' || r_InfoA.婚姻状况 ||  ',' || r_InfoA.职业 ||  ',' || r_InfoA.家庭地址;
                arrTxt = Split(rsTmp!原信息, ",")
                If UBound(arrTxt) >= 11 Then
                    .TextMatrix(i, 3) = arrTxt(0)
                    .TextMatrix(i, 4) = arrTxt(1)
                    .TextMatrix(i, 5) = arrTxt(2)
                    .TextMatrix(i, 6) = arrTxt(3)
                    .TextMatrix(i, 7) = arrTxt(4)
                    .TextMatrix(i, 8) = arrTxt(5)
                    .TextMatrix(i, 9) = arrTxt(6)
                    .TextMatrix(i, 10) = arrTxt(7)
                    .TextMatrix(i, 11) = arrTxt(8)
                    .TextMatrix(i, 12) = arrTxt(9)
                    .TextMatrix(i, 13) = arrTxt(10)
                    .TextMatrix(i, 14) = arrTxt(11)
                End If
                rsTmp.MoveNext
            Next
            
            If rsTmp.RecordCount = 0 Then .Rows = 2: .Row = 1: .FixedRows = 1
            .Redraw = True
        End With
        
    End If
    
    ReadCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo_Click(Index As Integer)
    If Index = cboName.主页ID Then      '启动加载住院次数时不调用
        If cbo(cboName.主页ID).Visible Then Call ReadCard(mlng病人ID, cbo(cboName.主页ID).ListIndex + 1, True)
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    '固定信息处理
    With cbo(cboName.年龄单位)
        .AddItem "岁"
        .AddItem "月"
        .AddItem "天"
        .ListIndex = 0
    End With
    
    If Not ReadCard(mlng病人ID, mlng主页ID) Then
        Call InitFaceScheme(True)
    Else
        Call InitFaceScheme
    End If
    
    '住院信息
    If Val(txt(txtName.住院次数)) > 0 Then
        With cbo(cboName.主页ID)
            For i = 1 To Val(txt(txtName.住院次数))
                .AddItem "第" & CStr(i) & "次住院"
                If i = mlng主页ID Then .ListIndex = .NewIndex '不会再调用readcard
            Next
        End With
        cbo(cboName.主页ID).Enabled = True
        cbo(cboName.主页ID).Locked = False
    Else
        cbo(cboName.主页ID).Enabled = False
        cbo(cboName.主页ID).Locked = True
    End If
    
End Sub
Private Sub InitFaceScheme(Optional blnNoRecord As Boolean)
    Dim Item As TabControlItem
    
    With TabSub
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True
        If blnNoRecord Then
            .InsertItem 1, "基本信息", PicNoRecord.Hwnd, 0
            PicBaseInfo.Visible = False: PicInInfo.Visible = False: vsList.Visible = False
        Else
            PicNoRecord.Visible = False
            .InsertItem 1, "基本信息", PicBaseInfo.Hwnd, 0
            .InsertItem 2, "住院信息", PicInInfo.Hwnd, 0
            .InsertItem 3, "合并记录", vsList.Hwnd, 0
        End If
        .Item(0).Selected = True
    End With
    PicBaseInfo.Width = Me.Width
    PicInInfo.Width = Me.Width
    vsList.Width = Me.Width
    PicNoRecord.Width = Me.Width
End Sub

Private Sub LoadOldData(ByVal strOld As String, ByRef txt年龄 As TextBox, ByRef cbo年龄单位 As ComboBox)
'功能:将数据库中保存的年龄按规范的格式加载到界面,不规范的原样显示
    Dim strTmp As String, lngIdx As Long
    
    If Trim(strOld) = "" Then Exit Sub
    
    lngIdx = -1
    strTmp = strOld
    If InStr(strOld, "岁") > 0 Then
        If InStr(strOld, "岁") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "岁") - 1)
            lngIdx = 0
        End If
    ElseIf InStr(strOld, "月") > 0 Then
        If InStr(strOld, "月") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "月") - 1)
            lngIdx = 1
        End If
    ElseIf InStr(strOld, "天") > 0 Then
        If InStr(strOld, "天") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "天") - 1)
            lngIdx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        lngIdx = 0
    End If
    
    txt年龄.Text = strTmp
    If cbo年龄单位.ListCount > 0 Then Call zlControl.CboSetIndex(cbo年龄单位.Hwnd, lngIdx)
    If lngIdx = -1 Then
        cbo年龄单位.Visible = False
    Else
        If cbo年龄单位.Visible = False Then cbo年龄单位.Visible = True
    End If
End Sub
