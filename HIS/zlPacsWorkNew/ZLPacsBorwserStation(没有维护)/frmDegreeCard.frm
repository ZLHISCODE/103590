VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ"
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
   StartUpPosition =   1  '����������
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
         Name            =   "����"
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
         Caption         =   "����ȼ�"
         Height          =   180
         Left            =   8085
         TabIndex        =   139
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ȼ�"
         Height          =   180
         Left            =   5670
         TabIndex        =   138
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   615
         TabIndex        =   137
         Top             =   810
         Width           =   540
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   6030
         TabIndex        =   136
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         Height          =   180
         Left            =   435
         TabIndex        =   135
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ����"
         Height          =   180
         Left            =   5670
         TabIndex        =   134
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ת����Ϣ"
         Height          =   180
         Left            =   435
         TabIndex        =   133
         Top             =   3630
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   5670
         TabIndex        =   132
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   2835
         TabIndex        =   131
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ���"
         Height          =   180
         Left            =   5850
         TabIndex        =   130
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���λ�ʿ"
         Height          =   180
         Left            =   2835
         TabIndex        =   129
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺҽʦ"
         Height          =   180
         Left            =   435
         TabIndex        =   128
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   435
         TabIndex        =   127
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   435
         TabIndex        =   126
         Top             =   2385
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ����"
         Height          =   180
         Left            =   2835
         TabIndex        =   125
         Top             =   4860
         Width           =   720
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3195
         TabIndex        =   124
         Top             =   810
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺĿ��"
         Height          =   180
         Left            =   435
         TabIndex        =   123
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ҽ���"
         Height          =   180
         Left            =   75
         TabIndex        =   122
         Top             =   4410
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   435
         TabIndex        =   121
         Top             =   2850
         Width           =   720
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ��ҽ���"
         Height          =   180
         Left            =   75
         TabIndex        =   120
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע"
         Height          =   180
         Left            =   6030
         TabIndex        =   119
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   2835
         TabIndex        =   118
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   435
         TabIndex        =   117
         Top             =   2010
         Width           =   720
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   8115
         TabIndex        =   116
         Top             =   2370
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժ���"
         Height          =   180
         Left            =   435
         TabIndex        =   115
         Top             =   4020
         Width           =   720
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
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
      Begin VB.CheckBox chk���� 
         BackColor       =   &H8000000A&
         Caption         =   "��ʱ"
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
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   540
         TabIndex        =   84
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   5415
         TabIndex        =   83
         Top             =   900
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
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
         Caption         =   "����ѱ�"
         Height          =   180
         Left            =   5415
         TabIndex        =   81
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3240
         TabIndex        =   80
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   540
         TabIndex        =   79
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3240
         TabIndex        =   78
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   5595
         TabIndex        =   77
         Top             =   195
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Height          =   180
         Left            =   360
         TabIndex        =   76
         Top             =   900
         Width           =   540
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ��ʽ"
         Height          =   180
         Left            =   7725
         TabIndex        =   75
         Top             =   555
         Width           =   1080
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2880
         TabIndex        =   74
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Left            =   180
         TabIndex        =   73
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   5415
         TabIndex        =   72
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   540
         TabIndex        =   71
         Top             =   1950
         Width           =   360
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   3240
         TabIndex        =   70
         Top             =   1590
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   5775
         TabIndex        =   69
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   3240
         TabIndex        =   68
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��"
         Height          =   180
         Left            =   8445
         TabIndex        =   67
         Top             =   1245
         Width           =   360
      End
      Begin VB.Label lvl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Left            =   180
         TabIndex        =   66
         Top             =   1590
         Width           =   720
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   180
         Left            =   180
         TabIndex        =   65
         Top             =   2655
         Width           =   720
      End
      Begin VB.Label lbl��ͥ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Left            =   180
         TabIndex        =   64
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lbl�����ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʱ�"
         Height          =   180
         Left            =   2880
         TabIndex        =   63
         Top             =   3015
         Width           =   720
      End
      Begin VB.Label lbl��ϵ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ������"
         Height          =   180
         Left            =   0
         TabIndex        =   62
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˹�ϵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˹�ϵ"
         Height          =   180
         Left            =   0
         TabIndex        =   61
         Top             =   4425
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ"
         Height          =   180
         Left            =   0
         TabIndex        =   60
         Top             =   3720
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵绰"
         Height          =   180
         Left            =   0
         TabIndex        =   59
         Top             =   4065
         Width           =   900
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         Height          =   180
         Left            =   5415
         TabIndex        =   58
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lbl��λ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   180
         Left            =   5415
         TabIndex        =   57
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label lbl��λ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Left            =   8085
         TabIndex        =   56
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label lbl��λ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ������"
         Height          =   180
         Left            =   5235
         TabIndex        =   55
         Top             =   3015
         Width           =   900
      End
      Begin VB.Label lbl��λ�ʺ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʺ�"
         Height          =   180
         Left            =   5415
         TabIndex        =   54
         Top             =   3360
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
         Height          =   180
         Left            =   2880
         TabIndex        =   53
         Top             =   4995
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
         Height          =   180
         Left            =   180
         TabIndex        =   52
         Top             =   4995
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
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
         Caption         =   "�������"
         Height          =   180
         Left            =   5415
         TabIndex        =   50
         Top             =   4065
         Width           =   720
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������ҽ���"
         Height          =   180
         Left            =   5055
         TabIndex        =   49
         Top             =   4425
         Width           =   1080
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽʦ"
         Height          =   180
         Left            =   5415
         TabIndex        =   48
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   7545
         TabIndex        =   47
         Top             =   4995
         Width           =   540
      End
      Begin VB.Label lbl�Ǽ�ʱ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ǽ�ʱ��"
         Height          =   180
         Left            =   8085
         TabIndex        =   46
         Top             =   195
         Width           =   720
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���￨��"
         Height          =   180
         Left            =   8085
         TabIndex        =   45
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lbl����֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����֤��"
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
         Caption         =   "������ȷ��ȡ������Ϣ  ����ϵͳ����Ա��ϵ��"
         BeginProperty Font 
            Name            =   "����"
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
Option Explicit 'Ҫ���������
Private mlng����ID As Long 'Ҫ�鿴�Ĳ���ID
Private mlng��ҳID As Long 'סԺ����ʱ������ҳID

Private Enum txtName
    'Ҫ���SQL���ֶζ�Ӧ
    ����ID = 0
    ���� = 1
    �Ա� = 2
    ���� = 3
    ����� = 4
    �ѱ� = 5
    ҽ�Ƹ��ʽ = 6
    ҽ���� = 7
    ���� = 8
    ���� = 9
    ���� = 10
    ���� = 11
    ѧ�� = 12
    ����״�� = 13
    ְҵ = 14
    ��� = 15
    �������� = 16
    ���֤�� = 17
    �����ص� = 18
    ��ͥ��ַ = 19
    ��ͥ��ַ�ʱ� = 20
    ��ͥ�绰 = 21
    ��ϵ������ = 22
    ��ϵ�˹�ϵ = 23
    ��ϵ�˵�ַ = 24
    ��ϵ�˵绰 = 25
    ������λ = 26
    ��λ�绰 = 27
    ��λ�ʱ� = 28
    ��λ������ = 29
    ��λ�ʺ� = 30
    Ԥ����� = 31
    ������� = 32
    ������ = 33
    ������ = 34
    
    סԺ���� = 35
    סԺ�� = 36
    ��Ժ���� = 37
    ��ע = 38
    סԺĿ�� = 39
    סԺ�ѱ� = 40
    ����ҽʦ = 41
    סԺҽʦ = 42
    ����ҽʦ = 57
    ����ҽʦ = 61
    ���λ�ʿ = 43
    �Ǽ��� = 44
    ��λ�ȼ� = 45
    ����ȼ� = 46
    ��Ժ���� = 47
    ��Ժ���� = 48
    ��Ժ���� = 49
    ��ǰ���� = 62
    ת����Ϣ = 52
    ��Ժ���� = 55
    ��Ժ���� = 56
    סԺ���� = 58
    
    ������� = 59
    ������ҽ��� = 60
    ��Ժ��� = 50
    ��Ժ��ҽ��� = 51
    ��Ժ��� = 53
    ��Ժ��ҽ��� = 54
    
End Enum

Private Enum cboName
    ��ҳID = 0
    ���䵥λ = 1
End Enum
Public Sub ShowMe(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    Me.Show 1
End Sub
Private Function ReadCard(ByVal lng����ID As Long, ByVal lng��ҳID As Long, Optional ByVal bln�鿴ĳ��סԺ As Boolean) As Boolean
'���ܣ���ȡָ��������Ϣ,����ʾ�ڽ�����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTxt As String, strTmp As String, strHead As String
    Dim i As Integer, j As Integer, arrTxt As Variant
    
    On Error GoTo errH
    
    strSQL = "Select a.����id, a.����, a.�Ա�, a.����, a.�����, a.�ѱ�, a.ҽ�Ƹ��ʽ, a.����, a.����, a.����, a.����, a.ѧ��," & vbNewLine & _
            "            a.����״��, a.ְҵ, a.���, Decode(To_Date(To_Char(��������, 'YYYY-MM-DD HH24:MI'), 'YYYY-MM-DD HH24:MI') - Trunc(��������), 0, To_Char(��������, 'YYYY-MM-DD'),To_char(��������,'YYYY-MM-DD HH24:MI')) ��������, " & _
            "            a.���֤��, a.�����ص�, a.��ͥ��ַ, a.��ͥ�绰, a.��ͥ��ַ�ʱ�, a.��ϵ������," & vbNewLine & _
            "            a.��ϵ�˹�ϵ, a.��ϵ�˵�ַ, a.��ϵ�˵绰, a.������λ, a.��λ�绰, a.��λ�ʱ�, a.��λ������, a.��λ�ʺ�, a.������," & vbNewLine & _
            "            a.������, a.��������, a.סԺ����, a.סԺ��, To_char(a.�Ǽ�ʱ��,'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, a.���￨��, b.��Ժ����, b.��ע, b.סԺĿ��, b.����ҽʦ, b.סԺҽʦ," & vbNewLine & _
            "            b.���λ�ʿ, b.�Ǽ���, b.��Ժ����, b.��ǰ����, b.סԺ����, b.�ѱ� As סԺ�ѱ�, c.Ԥ�����, c.�������," & vbNewLine & _
            "            Nvl(A.ҽ����,d.��Ϣֵ) ҽ����, e.���� As ����ȼ�, g.���� As ��λ�ȼ�, m.���� As ��Ժ����, n.���� As ��Ժ����," & vbNewLine & _
            "            To_char(b.��Ժ����,'yyyy-mm-dd hh24:mi:ss') ��Ժ����, To_char(b.��Ժ����,'yyyy-mm-dd hh24:mi:ss') ��Ժ����,A.����֤��,Nvl(B.��������,Decode(B.����,Null,'��ͨ����','ҽ������')) �������� " & vbNewLine & _
            "From ������Ϣ a, ������ҳ b, ������� c, ������ҳ�ӱ� d, �շ���ĿĿ¼ e, ��λ״����¼ f, �շ���ĿĿ¼ g, ���ű� m, ���ű� n" & vbNewLine & _
            "Where a.����id = b.����id(+) And " & IIf(lng��ҳID = 0, "Nvl(a.סԺ����,0)", "[2]") & "=b.��ҳID(+) And a.����ID=[1] And a.����id = c.����id(+) And" & vbNewLine & _
            "           c.����(+) = 1 And b.����id = d.����id(+) And" & vbNewLine & _
            "           b.��ҳid = d.��ҳid(+) And d.��Ϣ��(+) = 'ҽ����' And b.��Ժ����id = m.Id(+) And b.��Ժ����id = n.Id(+) And" & vbNewLine & _
            "           b.����ȼ�id = e.ID(+) And b.��ǰ����id = f.����id(+) And b.��Ժ���� = f.����(+) And f.�ȼ�id = g.ID(+)"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    If rsTmp.EOF Then Exit Function
        
    If bln�鿴ĳ��סԺ Then
       strTxt = "��Ժ����=37,��ע=38,סԺĿ��=39,סԺ�ѱ�=40,����ҽʦ=41,סԺҽʦ=42,���λ�ʿ=43,�Ǽ���=44,��λ�ȼ�=45,����ȼ�=46,��Ժ����=47,��Ժ����=48," & _
                " ��Ժ����=49,��ǰ����=62,ת����Ϣ=52,��Ժ����=55,��Ժ����=56,סԺ����=58,�������=59,������ҽ���=60,��Ժ���=50,��Ժ��ҽ���=51,��Ժ���=53,��Ժ��ҽ���=54,��������=66"
    Else
        strTxt = "����ID=0,����=1,�Ա�=2,����=3,�����=4,�ѱ�=5,ҽ�Ƹ��ʽ=6,ҽ����=7,����=8,����=9,����=10,����=11,ѧ��=12,����״��=13,ְҵ=14," & _
                " ���=15,��������=16,���֤��=17,�����ص�=18,��ͥ��ַ=19,��ͥ��ַ�ʱ�=20,��ͥ�绰=21,��ϵ������=22,��ϵ�˹�ϵ=23,��ϵ�˵�ַ=24,��ϵ�˵绰=25," & _
                " ������λ=26,��λ�绰=27,��λ�ʱ�=28,��λ������=29,��λ�ʺ�=30,Ԥ�����=31,�������=32,������=33,������=34,סԺ����=35,סԺ��=36," & _
                " ��Ժ����=37,��ע=38,סԺĿ��=39,סԺ�ѱ�=40,����ҽʦ=41,סԺҽʦ=42,���λ�ʿ=43,�Ǽ���=44,��λ�ȼ�=45,����ȼ�=46,��Ժ����=47,��Ժ����=48," & _
                " ��Ժ����=49,��ǰ����=62,ת����Ϣ=52,��Ժ����=55,��Ժ����=56,סԺ����=58,�������=59,������ҽ���=60,��Ժ���=50,��Ժ��ҽ���=51,��Ժ���=53," & _
                " ��Ժ��ҽ���=54,�Ǽ�ʱ��=63,���￨��=64,����֤��=65,��������=66"
    End If
    
    arrTxt = Split(strTxt, ",")
    
    For i = 0 To UBound(arrTxt)
        strTmp = Trim(arrTxt(i))
        
        If strTmp <> "" Then
            '�ſ��ݲ�������ֶ�
            If InStr(1, ",�������,������ҽ���,��Ժ���,��Ժ��ҽ���,��Ժ���,��Ժ��ҽ���,ת����Ϣ,", "," & Trim(Split(strTmp, "=")(0)) & ",") = 0 Then
                If InStr(1, ",�������,Ԥ�����,", "," & Trim(Split(strTmp, "=")(0)) & ",") > 0 Then
                    txt(Trim(Split(strTmp, "=")(1))).Text = Format(Val("" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))), "0.00")
                Else
                    txt(Trim(Split(strTmp, "=")(1))).Text = "" & rsTmp.Fields(Trim(Split(strTmp, "=")(0)))
                End If
            End If
        End If
    Next
    
    '����ר�Ŵ���
    '----------------------------------------------
    Call LoadOldData("" & rsTmp!����, txt(txtName.����), cbo(cboName.���䵥λ))
    If cbo(cboName.���䵥λ).ListIndex = -1 Then txt(txtName.����).Width = txt(txtName.����).Width + cbo(cboName.���䵥λ).Width
    chk����.Value = Val("" & rsTmp!��������)
    
    
    'סԺ��Ϣ
    '----------------------------------------------
    If Not bln�鿴ĳ��סԺ Then lng��ҳID = Val(txt(txtName.סԺ����).Text)
    'סԺ���˵�������
    If lng��ҳID > 0 Then
        strSQL = "Select �������,����ID,������� From ������ϼ�¼ Where ��ϴ���=1 And ��¼��Դ=2 And ����ID=[1] And ��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Select Case rsTmp!�������
                    Case 1
                        j = txtName.�������
                    Case 11
                        j = txtName.������ҽ���
                    Case 2
                        j = txtName.��Ժ���
                    Case 12
                        j = txtName.��Ժ��ҽ���
                    Case 3
                        j = txtName.��Ժ���
                    Case 13
                        j = txtName.��Ժ��ҽ���
                    Case Else
                        j = 0
                End Select
                If j <> 0 Then txt(j).Text = IIf(IsNull(rsTmp!����id), "", "(" & rsTmp!����id & ")") & rsTmp!�������
                
                rsTmp.MoveNext
            Next
        Else
            txt(txtName.�������).Text = ""
            txt(txtName.������ҽ���).Text = ""
            txt(txtName.��Ժ���).Text = ""
            txt(txtName.��Ժ��ҽ���).Text = ""
            txt(txtName.��Ժ���).Text = ""
            txt(txtName.��Ժ��ҽ���).Text = ""
        End If
        
        'ת����Ϣ
        txt(txtName.ת����Ϣ).Text = ""
        strSQL = _
            " Select Distinct 1 as ��ʼԭ��,To_Date('1900-01-01','YYYY-MM-DD') as ��ʼʱ��,B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=B.ID And A.��ʼʱ�� is Not NULL And A.��ʼԭ�� IN(1,2)" & _
            " And A.����ID=[1] And ��ҳID=[2]" & _
            " Union ALL " & _
            " Select A.��ʼԭ��,A.��ʼʱ��,B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=B.ID And A.��ʼʱ�� is Not NULL And A.��ʼԭ��=3" & _
            " And A.����ID=[1] And ��ҳID=[2]" & _
            " Order by ��ʼʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        rsTmp.Filter = "��ʼԭ��=3"
        If Not rsTmp.EOF Then
            rsTmp.Filter = 0
            Do While Not rsTmp.EOF
                txt(txtName.ת����Ϣ).Text = txt(txtName.ת����Ϣ).Text & " ���� " & rsTmp!����
                rsTmp.MoveNext
            Loop
            txt(txtName.ת����Ϣ).Text = Mid(txt(txtName.ת����Ϣ).Text, 5)
        End If
        
        '������ҳ�ӱ�
        txt(txtName.����ҽʦ).Text = ""
        txt(txtName.����ҽʦ).Text = ""
        strSQL = " Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where (��Ϣ��='����ҽʦ' Or ��Ϣ��='����ҽʦ') And ����ID=[1] And ��ҳID=[2]"
        rsTmp.Filter = ""
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
        rsTmp.Filter = "��Ϣ��='����ҽʦ'"
        If Not rsTmp.EOF Then txt(txtName.����ҽʦ).Text = "" & rsTmp!��Ϣֵ
        rsTmp.Filter = "��Ϣ��='����ҽʦ'"
        If Not rsTmp.EOF Then txt(txtName.����ҽʦ).Text = "" & rsTmp!��Ϣֵ
    End If
    
    
    '3.���˺ϲ���Ϣ
    If Not bln�鿴ĳ��סԺ Then
        
        strSQL = "Select ԭ��Ϣ,�ϲ�ԭ��,����Ա����,�ϲ�ʱ�� From ���˺ϲ���¼ Where ����ID=[1] Order by �ϲ�ʱ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
                
        strHead = "�ϲ�ʱ��,1,1800|����Ա,4,800|�ϲ�ԭ��,1,1800|" & _
                "����ID,1,800|�����,1,900|סԺ��,1,800|���￨��,1,900|����,4,800|" & _
                "�Ա�,4,500|����,4,800|��������,1,1000|���֤��,1,1800|����״��,4,900|ְҵ,1,1000|��ͥ��ַ,1,4200"
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
                .TextMatrix(i, 0) = Format(rsTmp!�ϲ�ʱ��, "yyyy-MM-dd HH:mm:ss")
                .TextMatrix(i, 1) = "" & rsTmp!����Ա����
                .TextMatrix(i, 2) = "" & rsTmp!�ϲ�ԭ��
                
                'v_ԭ��Ϣ:=r_InfoA.����Id || ',' || r_InfoA.����� || ',' ||  r_InfoA.סԺ�� || ',' ||  r_InfoA.���￨�� || ',' ||  r_InfoA.���� ||  ',' ||  r_InfoA.�Ա� ||  ',' ||
                '   r_InfoA.���� ||  ',' || to_char(r_InfoA.��������,'yyyy-mm-dd') ||  ',' || r_InfoA.���֤�� ||  ',' || r_InfoA.����״�� ||  ',' || r_InfoA.ְҵ ||  ',' || r_InfoA.��ͥ��ַ;
                arrTxt = Split(rsTmp!ԭ��Ϣ, ",")
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
    If Index = cboName.��ҳID Then      '��������סԺ����ʱ������
        If cbo(cboName.��ҳID).Visible Then Call ReadCard(mlng����ID, cbo(cboName.��ҳID).ListIndex + 1, True)
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
    
    '�̶���Ϣ����
    With cbo(cboName.���䵥λ)
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
        .ListIndex = 0
    End With
    
    If Not ReadCard(mlng����ID, mlng��ҳID) Then
        Call InitFaceScheme(True)
    Else
        Call InitFaceScheme
    End If
    
    'סԺ��Ϣ
    If Val(txt(txtName.סԺ����)) > 0 Then
        With cbo(cboName.��ҳID)
            For i = 1 To Val(txt(txtName.סԺ����))
                .AddItem "��" & CStr(i) & "��סԺ"
                If i = mlng��ҳID Then .ListIndex = .NewIndex '�����ٵ���readcard
            Next
        End With
        cbo(cboName.��ҳID).Enabled = True
        cbo(cboName.��ҳID).Locked = False
    Else
        cbo(cboName.��ҳID).Enabled = False
        cbo(cboName.��ҳID).Locked = True
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
            .InsertItem 1, "������Ϣ", PicNoRecord.Hwnd, 0
            PicBaseInfo.Visible = False: PicInInfo.Visible = False: vsList.Visible = False
        Else
            PicNoRecord.Visible = False
            .InsertItem 1, "������Ϣ", PicBaseInfo.Hwnd, 0
            .InsertItem 2, "סԺ��Ϣ", PicInInfo.Hwnd, 0
            .InsertItem 3, "�ϲ���¼", vsList.Hwnd, 0
        End If
        .Item(0).Selected = True
    End With
    PicBaseInfo.Width = Me.Width
    PicInInfo.Width = Me.Width
    vsList.Width = Me.Width
    PicNoRecord.Width = Me.Width
End Sub

Private Sub LoadOldData(ByVal strOld As String, ByRef txt���� As TextBox, ByRef cbo���䵥λ As ComboBox)
'����:�����ݿ��б�������䰴�淶�ĸ�ʽ���ص�����,���淶��ԭ����ʾ
    Dim strTmp As String, lngIdx As Long
    
    If Trim(strOld) = "" Then Exit Sub
    
    lngIdx = -1
    strTmp = strOld
    If InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 0
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            lngIdx = 2
        End If
    ElseIf IsNumeric(strOld) Then
        lngIdx = 0
    End If
    
    txt����.Text = strTmp
    If cbo���䵥λ.ListCount > 0 Then Call zlControl.CboSetIndex(cbo���䵥λ.Hwnd, lngIdx)
    If lngIdx = -1 Then
        cbo���䵥λ.Visible = False
    Else
        If cbo���䵥λ.Visible = False Then cbo���䵥λ.Visible = True
    End If
End Sub
