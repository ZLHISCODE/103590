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
   Caption         =   "���鱨���ѯ"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15435
   Icon            =   "frmPatientReprotFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   15435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
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
               Caption         =   "�� Ⱦ ��"
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
            Caption         =   "�� �� ��"
            Height          =   180
            Left            =   180
            TabIndex        =   58
            Top             =   615
            Width           =   720
         End
         Begin VB.Label Label5 
            Caption         =   "��ӡ״̬"
            Height          =   195
            Left            =   3030
            TabIndex        =   49
            Top             =   255
            Width           =   735
         End
         Begin VB.Label lblDept 
            Caption         =   "�������"
            Height          =   225
            Left            =   180
            TabIndex        =   47
            Top             =   255
            Width           =   930
         End
         Begin VB.Label lblName 
            Caption         =   "��    ��"
            Height          =   225
            Left            =   5400
            TabIndex        =   44
            Top             =   255
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "����ʱ��"
            Height          =   225
            Left            =   5400
            TabIndex        =   11
            Top             =   615
            Width           =   720
         End
         Begin VB.Label lblNo 
            Caption         =   "����š�"
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
         Caption         =   "�ٴ�����:"
         Height          =   180
         Left            =   30
         TabIndex        =   89
         Top             =   3090
         Width           =   810
      End
      Begin VB.Label lblDiagnose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���:"
         Height          =   180
         Left            =   30
         TabIndex        =   31
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע:"
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
               Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "����������:"
            Height          =   180
            Left            =   90
            TabIndex        =   27
            Top             =   90
            Width           =   1170
         End
         Begin VB.Label lblContrast 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ˢ��"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "������(&1)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "���ֵ(&2)"
            BeginProperty Font 
               Name            =   "����"
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
            Caption         =   "ͼ������"
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
            Caption         =   "���ѡ��"
            BeginProperty Font 
               Name            =   "����"
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
               Caption         =   "������"
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
               Caption         =   "��ϸ������"
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
               Caption         =   "���²�������"
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
               Caption         =   "����"
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
               Caption         =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "��������"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "δ �� ��"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "����"
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
            Caption         =   "������"
            BeginProperty Font 
               Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Text            =   "��΢�����"
               Top             =   270
               Width           =   3915
            End
            Begin VB.Label lblMicroscopeFinded 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "δ �� ��"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "ͨ���豸"
               BeginProperty Font 
                  Name            =   "����"
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
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "����:"
               Height          =   180
               Left            =   30
               TabIndex        =   42
               Top             =   2820
               Width           =   450
            End
            Begin VB.Label lblAntibiotic 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������:"
               Height          =   180
               Left            =   30
               TabIndex        =   40
               Top             =   1560
               Width           =   630
            End
            Begin VB.Label lblMicrobe 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ϸ��:"
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
               Caption         =   "��ʾ�����Ŀ"
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
                  Name            =   "����"
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
            Caption         =   "δ��"
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
            Caption         =   "δ֪"
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
            Caption         =   "����"
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
            Caption         =   "סԺ"
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
            Caption         =   "Ժ��"
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
            Caption         =   "���"
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
               Name            =   "����"
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
            Caption         =   "�ѳ�"
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
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":DC0B
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":1446D
            Key             =   "�ϰ�"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":1ACCF
            Key             =   "�°�"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotFind.frx":21531
            Key             =   "���"
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
Private mblnShow As Boolean                                         '�����Ƿ���ʾ
Private mlngKey As Long                                                 '��ǰѡ��ı걾ID
Private mlngPatientID As Long                                              '����ID
Private mdteReportDate As Date                                             '����ʱ��
Private mlngGetPatientID As Long                                        '�ϼ������Ĳ���ID
Private mSampleShowColour As SampleValShowColour                    '�����ʾ��ɫ
Private mstrPrivs As String                                         '������ϼ���Ȩ��
Private mblnLoad  As Boolean                                        '�����Ƿ��һ����ʾ
Private mlngValueC As Long                                              '΢����������
Private mintVer As Integer                                              '�汾��25-�°� 10-�ϰ�
Private mlngSelRow As Long                                              '֮ǰѡ����
Private mintIn As Integer                                               'ʵ���Ҳ鿴
Private mlngPicLeftWidth As Long         '��಼�ֿ��
Private mlngPicCenterWidth As Long       '�м䲼���


Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long
Private mstrTag As String           '���������ӡ

Private Type SampleValShowColour                                    '�����ɫ��ʾ

    ���� As Double
    ƫ�� As Double
    ƫ�� As Double
    �쳣 As Double
    ��ʾƫ�� As Double
    ��ʾƫ�� As Double
    ����ƫ�� As Double
    ����ƫ�� As Double
End Type
Private mobjFSO As New Scripting.FileSystemObject    'FSO����
Private mObjImg As Object

Private mObjIco As IcoObject                                            'ͼ�����

'����ͼ�����
'�����frmWorkBaseReprot,frmWorkBaseReprotFind,frmWorkBaseAuditingSample
'���������е�����һ�������ͼ��ؼ��������ͼƬ,�����������嶼��Ҫͬ������
Private Enum mIcoIndex
    ѡ�� = 1
    ��ӡ
    �°�
    �ϰ�
    ���
End Enum

Private Type IcoObject
    Obj���  As Object
    Obj��ӡ As Object
    Objѡ�� As Object
    Obj�°� As Object
    Obj�ϰ� As Object
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
    '��ʼ��ͼ������
    If mobjImgList Is Nothing Or Not mObjIco.Obj��ӡ Is Nothing Then
        Exit Sub
    End If
    Set mObjIco.Obj��� = mobjImgList.ListImages(mIcoIndex.���).ExtractIcon
    Set mObjIco.Obj��ӡ = mobjImgList.ListImages(mIcoIndex.��ӡ).ExtractIcon
    Set mObjIco.Objѡ�� = mobjImgList.ListImages(mIcoIndex.ѡ��).ExtractIcon
    Set mObjIco.Obj�°� = mobjImgList.ListImages(mIcoIndex.�°�).ExtractIcon
    Set mObjIco.Obj�ϰ� = mobjImgList.ListImages(mIcoIndex.�ϰ�).ExtractIcon
End Sub

Private Sub setIcoFree()
    '�ͷ���Դ
    Set mObjIco.Obj��� = Nothing
    Set mObjIco.Obj��ӡ = Nothing
    Set mObjIco.Objѡ�� = Nothing
    Set mObjIco.Obj�°� = Nothing
    Set mObjIco.Obj�ϰ� = Nothing
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
        '������ʾ��ѡ���ӡ״̬ˢ�²����б�
       Call ReadPatientList(1)
    Else
        mblnLoad = True
    End If
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_SelAll
            VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("ѡ��"), 1
        Case ConMenu_Browse_ClsAll
            VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("ѡ��"), 2
        Case conFun_Sample_Auditing     '����
            Call AuditingSample(1)
        Case conFun_Sample_unAuditing     'ȡ������
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
                If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�汾"))) = 25 Then
                    PrintReport Me, mlngKey, 3
                Else
                    PtintOldReport Me, mlngKey, , 3
                End If
            End If
        Case ConMenu_Browse_PrintView   'Ԥ��
            If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�汾"))) = 25 Then
                PrintReport Me, mlngKey, 1
            Else
                PtintOldReport Me, mlngKey, , 1
            End If
        Case ConMenu_Browse_Exit
            Unload Me
        Case ConMenu_pop_SampleCode
            lblNo.Caption = "����š�"
        Case ConMenu_pop_Out
            lblNo.Caption = "����š�"
        Case ConMenu_pop_In
            lblNo.Caption = "סԺ�š�"
        Case ConMenu_pop_bed
            lblNo.Caption = "  ���š�"
        Case ConMenu_pop_PatiCard
            lblNo.Caption = "���￨��"
        Case ConMenu_Browse_unPrint     '���ô�ӡ
            Call ResetPrintType
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
            Call ExePlugIn(Control.Parameter, mlngKey)
    End Select
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2017/6/13
'��    ��:���������������ӡ����
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub ResetPrintType()
          Dim strSQL As String

1         On Error GoTo ResetPrintType_Error

2         If mlngKey <= 0 Then Exit Sub
3         strSQL = "Zl_���鱨���ӡ_Edit(2," & mlngKey & ",2)"
4         Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
5         SaveDBLog 18, 6, mlngKey, "��ӡ", "������������ӡ����", 2500, "�ٴ�ʵ���ҹ���"

6         MsgBox "��������ӡ״̬������", vbInformation, Me.Caption


7         Exit Sub
ResetPrintType_Error:
8         Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(ResetPrintType)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
9         Err.Clear

End Sub

Private Sub AuditingSample(ByVal intType As Integer)
          '����/ȡ������
          'intType    1=����,2=ȡ������

          Dim strSQL As String
          Dim lngSampleKey As String  '�걾id

1         On Error GoTo AuditingSample_Error

2         With Me.vsfLeft
3             If .Row > 0 Then
4                 lngSampleKey = Val(.TextMatrix(.Row, .ColIndex("ID")))
5             Else
6                 MsgBox "��ѡ�д�Ⱦ����¼", vbInformation, Me.Caption
7                 Exit Sub
8             End If
9         End With

10        strSQL = "Zl_���鴫Ⱦ������_Edit(" & intType & "," & lngSampleKey & ",'" & UserInfo.Name & "')"
11        Call ComExecuteProc(Sel_Lis_DB, strSQL, "��Ⱦ�����渴��")

12        If vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�汾")) = 25 Then
13            SaveDBLog 18, 6, Val(lngSampleKey), IIf(intType = 1, "����", "ȡ������"), IIf(intType = 1, "����", "ȡ������"), 2500, "�ٴ�ʵ���ҹ���"
14        End If

          'ˢ���б�
15        mlngPatientID = 0
16        Call ReadPatientList(1)
17        Call vsfLeft_Click


18        Exit Sub
AuditingSample_Error:
19        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(AuditingSample)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
      Select Case Control.ID
        Case ConMenu_Browse_PrintView   'Ԥ��
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
        If .ColIndex("���") < 0 Then Exit Sub
        For i = 1 To .Rows - 1
            If .Cell(flexcpFontBold, i, .ColIndex("���")) = True Then
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
    Call frmPaitReport.ShowMe(Me, 705845, "ICU����;�����һ���չʾ����;���Ʋ���;��������ύ;��������;��ӡ��ҳ;����ʡ������ҳ;����ʡ��ҽ������ҳ;����໤;���ﲡ��;����;����ҩ��Խ��ʹ�û��ܱ�;����ҩ��Խ��ʹ����ϸ��;�ٴ��Թ�ҩ;ȫԺ����;��鷴������;��ҳ������Ϣ;��ҳ����;�Ĵ�ʡ��ҽ��ҳ;�Ĵ�ʡ��ҽ��ҳ;Σ��ֵ����;�޸������ȼ�;�޸�ҽ�Ƹ��ʽ;ҩռ�Ȳ�ѯ;ԤԼ�Һ�;ԤԼ�Һŵ�;����ʡ��ҽ��ҳ;����ʡ��ҽ��ҳ;��ҽ������ҳ;סԺһ��", 57, 57, 2, 1, , True, , True)
End Sub

Private Sub Form_Activate()
    If mblnShow = False Then
        InitFace
        If mlngPatientID = 0 Then Call ReadPatientList(-1, True)
        mblnShow = True
    End If
End Sub

Private Sub Form_Load()
    '���ܴ���������
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
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_SelAll, "ȫѡ")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_ClsAll, "ȫ��")

        Set cbrControl = .Add(xtpControlButton, conFun_Sample_Auditing, "����")
        cbrControl.Visible = False
        cbrControl.Enabled = False: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conFun_Sample_unAuditing, "ȡ������")
        cbrControl.Visible = False
        cbrControl.Enabled = False

        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "��ӡ")
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "��ӡ����  ")
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_unPrint, "���ô�ӡ  ")
            cbrControl.Visible = InStr(mstrPrivs, "���������������ӡ����") > 0
        End With
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "Ԥ��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "�˳�"): cbrControl.BeginGroup = True
    End With

    '���������ť
    Call CreatePlugInButton(cbrToolBar)

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next



    '�б�
    With Me.TabPage
        .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True

        '��ע
        .InsertItem 1, "��ע", picComment.hWnd, ConTab_Sample_Comment
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        '�б�
        .InsertItem 2, "����", PICContrast.hWnd, ConTab_Sample_History
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        .InsertItem 3, "ͼ��", PicPic.hWnd, ConTab_Sample_Comment
        .PaintManager.Position = xtpTabPositionTop
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        .Item(0).Selected = True
    End With

    With cboPrint
        .AddItem "����"
        .AddItem "�Ѵ�ӡ"
        .AddItem "δ��ӡ"
        .ListIndex = 0
    End With

    dtpS = Now - 7
    dtpE = Now
    picLeft.Width = GetSetWith(1)

    Me.chkGroup.value = Val(ComGetPara(Sel_Lis_DB, "�Ƿ���ʾ�����Ŀ", gSysInfo.SysNo, gSysInfo.ModlNo, 1))
    strPicWidth = GetSetting("ZLSOFT", "˽��ģ��\" & UserInfo.DBUser & "\" & App.EXEName & "\" & Me.Name, "PICWIDTH", "")
    If strPicWidth <> "" Then
        mlngPicLeftWidth = Val(Split(strPicWidth, ";")(0))
        mlngPicCenterWidth = Val(Split(strPicWidth, ";")(1))
    End If
    
    ReadSampleBacteriology 0
    ReadSampleBacteriology 0
    ReadSampleVal 0

End Sub

Private Function GetSetWith(ByVal intType As Integer) As Long
    '��ȡ/���ô�����߲��ֵĿ��
    '1-��ȡ,2-����
    If intType = 1 Then
        GetSetWith = ComGetPara(Sel_Lis_DB, "���鱨����Ϣ��", 2500, 2500, "5000")
    ElseIf intType = 2 Then
        Call ComSetPara(Sel_Lis_DB, "���鱨����Ϣ��", picLeft.Width, 2500, 2500)
    End If
End Function

Public Sub PicDrowBorder(Picobj As PictureBox, Optional lngLineColour As Long = -1)
    '����       ��ͼƬ�߿�
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

    Call ComSetPara(Sel_Lis_DB, "�Ƿ���ʾ�����Ŀ", Me.chkGroup.value, gSysInfo.SysNo, gSysInfo.ModlNo)
    Call setIcoFree
    '���������
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & UserInfo.DBUser & "\" & App.EXEName & "\" & Me.Name, "PICWIDTH", picLeft.Width & ";" & picCenter.Width)
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
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_SampleCode, "�����")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Out, "�����")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_In, "סԺ��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_bed, "����")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_PatiCard, "���￨")
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
      '����           ���������������б�
      '               blnLoadFrm=True  ��ʼ������ʱ���������ݣ�ֻ����VSF
          Dim rsTmp As ADODB.Recordset, rsOldLisData As ADODB.Recordset
          Dim strSQL As String
          Dim lngKey As Long
          Dim strDepts As String
          Dim strDept As String
          Dim lngPatiID As Long
          Dim strTemp As String
          Dim strWhere As String
          Dim lngLoop As Long
          Dim strTitle As String   '�б�
          Dim var_tmp As Variant
          Dim var_SubTmp As Variant
          Dim blnReadData As Boolean
          Dim strTiredFind As String
          Dim strFindSQL As String

          '�����������
1         On Error GoTo ReadPatientList_Error

          '��ȡ����ֵ
2         strTitle = ComGetPara(Sel_Lis_DB, "����������ʾ��", 2500, 1013)

3         If Trim(txtPatiNo <> "") Then
4             If lblNo.Caption = "���￨��" Then
5                 strSQL = "select ����id from ������Ϣ where ���￨��  = [1] "
6                 Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���Ҳ��˾��￨", txtPatiNo)
7                 If rsTmp.RecordCount > 0 Then
8                     lngPatiID = rsTmp("����id")
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

          '�߷�ʱ�����Ʋ�ѯ
19        blnReadData = True
20        If blnLoadFrm = False Then
21            If mintIn = 0 Then
22                If Not funCheckRushHours(2500, 1013, "���鱨���ѯ����", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59")) Then
23                    blnReadData = False
24                Else
25                    blnReadData = True
26                End If
27            End If
28        End If

29        If blnReadData = True Then


30            strSQL = "Select a.Id, a.ѡ��, a.����, a.�Ա�, a.����, a.��ӡ, a.����, a.������Ŀ, a.סԺ��, a.����, a.����ʱ��, a.����id, a.����ʱ��, a.���ʱ��, a.��ע, a.���, a.΢����," & vbNewLine & _
                     "          a.���Ա���, a.������, a.�����, a.������Դ, a.������,�������, a.�汾, a.�������, b.����ʱ�� ��ӡʱ��, b.����Ա ��ӡ��,a.�Ƿ�Ⱦ��,a.������,a.����ʱ��,a.��������,a.������,a.����ʱ��,a.������" & vbNewLine & _
                     "   From (Select a.Id, 0 ѡ��, a.����, Decode(a.�Ա�, 1, '��', 2, 'Ů', 9, 'δ֪', '') �Ա�, Decode(a.�����, Null, 'δ��', '�ѳ�') ����," & vbNewLine & _
                     "                 Decode(a.��ӡ����, Null, (Decode(a.ҽ��վ��ӡ, Null, Decode(a.������ӡ����, Null, 0, 1), 1)), 1) ��ӡ, a.����, c.���� ������Ŀ, a.סԺ��," & vbNewLine & _
                     "                 a.����, b.����ʱ��, a.����id, a.����ʱ��, a.���ʱ��, a.��ע, a.���, a.΢����, a.���Ա���, a.������, a.�����, Nvl(a.������Դ, 0) ������Դ, a.������,a.�������," & vbNewLine & _
                     "                 25 �汾, 0 �������, (Select Max(ID) From ���������־ D Where d.�걾id = a.Id And d.�������� = '�걾��ӡ') ������־id,a.�Ƿ�Ⱦ��,a.������,a.����ʱ��,a.��������,a.������,a.����ʱ��,a.������" & vbNewLine & _
                     "          From ���鱨���¼ A, ����������� B, ���������Ŀ C" & vbNewLine & _
                     "          Where a.Id(+) = b.�걾id And b.���id = c.Id(+) And a.����ʱ�� Between [1] And [2] [����]) A, ���������־ B" & vbNewLine & _
                     "   Where a.������־id = b.Id(+)"

31            Select Case cboPrint.ListIndex

              Case 1
                  '�Ѵ�ӡ
32                strWhere = strWhere & "and ( a.��ӡ���� is not null or a.ҽ��վ��ӡ is not  null or a.������ӡ���� is not null ) "
33            Case 2
                  'δ��ӡ
34                strWhere = strWhere & " and a.��ӡ���� is null   and  a.ҽ��վ��ӡ  is  null and a.������ӡ���� is  null "
35            End Select

36            If chkAudit(0).value = 1 And chkAudit(1).value = 0 Then
37                strWhere = strWhere & "And a.����� Is Not Null And a.���ʱ�� Is Not Null"
38            ElseIf chkAudit(0).value = 0 And chkAudit(1).value = 1 Then
39                strWhere = strWhere & "And a.����� Is  Null And a.���ʱ�� Is  Null"
40            ElseIf chkAudit(0).value = 1 And chkAudit(1).value = 1 Then

41            End If

              '��Ⱦ��
42            Select Case Me.cboDiseases.Text
              Case "��Ⱦ��"
43                strWhere = strWhere & " and a.�Ƿ�Ⱦ��=1"
44            Case "�Ǵ�Ⱦ��"
45                strWhere = strWhere & " and (a.�Ƿ�Ⱦ��<>1 or �Ƿ�Ⱦ�� is null)"
46            End Select

              '��������Դ
47            strTemp = checkboxSource()
48            If strTemp <> "" Then
49                strWhere = strWhere & " and nvl(a.������Դ,0) in (" & strTemp & ")"
50            Else
                  'Ϊѡ������Դʱ������Դ����Ϊ-1
51                strWhere = strWhere & " and nvl(a.������Դ,0) in (-1)"
52            End If

53            If lngID = -1 Then
54                strWhere = strWhere & " and a.id = -1 "
55                strFindSQL = strFindSQL & " and a.id = -1 "
56            End If

57            If txtName <> "" Then
58                strWhere = strWhere & " and a.���� like '" & txtName & "%' "
59                strFindSQL = strFindSQL & " and a.���� like '" & txtName & "%' "
60            End If

61            If txtDoctor.Text <> "" Then
62                strWhere = strWhere & " and a.������ like '" & txtDoctor.Text & "%'"
63                strFindSQL = strFindSQL & " and a.������ like '" & txtDoctor.Text & "%'"
64            End If

65            If Trim(txtPatiNo <> "") Then
66                If lblNo.Caption = "סԺ�š�" Then
67                    strWhere = strWhere & " and a.סԺ�� = [3] "
68                    strFindSQL = strFindSQL & " and a.סԺ�� = [3] "
69                ElseIf lblNo.Caption = "����š�" Then
70                    strWhere = strWhere & " and a.����� = [3] "
71                    strFindSQL = strFindSQL & " and a.����� = [3] "
72                ElseIf lblNo.Caption = "���š�" Then
73                    strWhere = strWhere & " and a.���� = [3] "
74                    strFindSQL = strFindSQL & " and a.���� = [3] "
75                ElseIf lblNo.Caption = "���￨��" Then
76                    strWhere = strWhere & " and a.HIS����ID = [7] "
77                    strFindSQL = strFindSQL & " and a.HIS����ID = [7] "
78                ElseIf lblNo.Caption = "����š�" Then
79                    strWhere = strWhere & " and a.�������� = [3] "
80                    strFindSQL = strFindSQL & " and a.�������� = [3] "
81                End If
82            End If
83            If cboDept <> "" Then
84                strDept = Mid(cboDept.Text, InStr(cboDept.Text, "-") + 1)
85                If InStr(strDept, "���п���") > 0 Then

86                Else
87                    strWhere = strWhere & " and a.������� =[6] "
88                    strFindSQL = strFindSQL & " and a.������� =[6] "
89                End If
90            End If
91            DoEvents
92            If lngPatiID = 0 And mlngPatientID <> 0 Then
93                lngPatiID = mlngPatientID
94                strWhere = strWhere & " and a.HIS����ID = [7] "
95                strFindSQL = strFindSQL & " and a.HIS����ID = [7] "
96            End If

97            strSQL = Replace(strSQL, "[����]", strWhere)


98            strTiredFind = "   union all    Select a.����id  Id, 0 ѡ��, a.����, a.�Ա�, '�ѳ�' ����, 0 ��ӡ, a.����, c.���� ������Ŀ, a.סԺ��, a.����, a.����ʱ��, a.����id, sysdate ����ʱ��, sysdate ���ʱ��, '' ��ע, a.���," & vbNewLine & _
                           " 1 ΢����, 3 ���Ա���, a.�ͼ���, '' �����, a.������Դ, a.������, �������, 25 �汾, 1 ������� , null ��ӡʱ��, '' ��ӡ��, 0 �Ƿ�Ⱦ��, '' ������, null ����ʱ��," & vbNewLine & _
                           " a.��������, '' ������, null ����ʱ��,null ������" & vbNewLine & _
                             "From ����������� A ,���������Ŀ C " & vbNewLine & _
                           " Where a.����ʱ�� Between [1] And [2] And a.����״̬ = 4 and a.���id =c.id(+) "
99            strTiredFind = strTiredFind & strFindSQL
100           strSQL = strSQL & strTiredFind
101           strSQL = " select * from (" & strSQL & " ) order by ���ʱ��,����id,id"
102           If mintIn = 1 Then
103               strSQL = Replace(strSQL, "And a.����ʱ�� Between [1] And [2]", "")

104           End If
105           Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���벡���б�", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
                                     CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59"), txtPatiNo, mlngGetPatientID, strDepts, strDept, lngPatiID)
106       End If

107       With vsfLeft
108           If strTitle = "" Then
109               .Rows = 1
110               .Cols = 35
111               .FixedRows = 1
112               .ColKey(0) = "���": .ColWidth(.ColIndex("���")) = 500: .ColAlignment(.ColIndex("���")) = flexAlignCenterCenter
113               .ColKey(1) = "id": .ColWidth(.ColIndex("id")) = 2000: .ColAlignment(.ColIndex("id")) = flexAlignCenterCenter: .ColHidden(.ColIndex("id")) = True
114               .ColKey(2) = "ѡ��": .ColWidth(.ColIndex("ѡ��")) = 250: .ColAlignment(.ColIndex("ѡ��")) = flexAlignCenterCenter: .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
115               .ColKey(3) = "��ӡ": .ColWidth(.ColIndex("��ӡ")) = 300: .ColAlignment(.ColIndex("��ӡ")) = flexAlignCenterCenter
116               .ColKey(4) = "�汾": .ColWidth(.ColIndex("�汾")) = 300: .ColAlignment(.ColIndex("�汾")) = flexAlignCenterCenter
117               .ColKey(5) = "����": .ColWidth(.ColIndex("����")) = 500: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter

118               .ColKey(6) = "������Դ": .ColWidth(.ColIndex("������Դ")) = 420: .ColAlignment(.ColIndex("������Դ")) = flexAlignCenterCenter
119               .ColKey(7) = "����": .ColWidth(.ColIndex("����")) = 750: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
120               .ColKey(8) = "�Ա�": .ColWidth(.ColIndex("�Ա�")) = 500: .ColAlignment(.ColIndex("�Ա�")) = flexAlignCenterCenter
121               .ColKey(9) = "����": .ColWidth(.ColIndex("����")) = 500: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
122               .ColKey(10) = "������Ŀ": .ColWidth(.ColIndex("������Ŀ")) = 2200: .ColAlignment(.ColIndex("������Ŀ")) = flexAlignCenterCenter
123               .ColKey(11) = "��������": .ColWidth(.ColIndex("��������")) = 1300: .ColAlignment(.ColIndex("��������")) = flexAlignCenterCenter
124               .ColKey(12) = "���ʱ��": .ColWidth(.ColIndex("���ʱ��")) = 2000: .ColAlignment(.ColIndex("���ʱ��")) = flexAlignCenterCenter
125               .ColKey(13) = "סԺ��": .ColWidth(.ColIndex("סԺ��")) = 750: .ColAlignment(.ColIndex("סԺ��")) = flexAlignCenterCenter
126               .ColKey(14) = "����": .ColWidth(.ColIndex("����")) = 500: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
127               .ColKey(15) = "����ʱ��": .ColWidth(.ColIndex("����ʱ��")) = 2000: .ColAlignment(.ColIndex("����ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ʱ��")) = True
128               .ColKey(16) = "����ID": .ColWidth(.ColIndex("����ID")) = 2000: .ColAlignment(.ColIndex("����ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ID")) = True
129               .ColKey(17) = "����ʱ��": .ColWidth(.ColIndex("����ʱ��")) = 2000: .ColAlignment(.ColIndex("����ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ʱ��")) = True
130               .ColKey(18) = "��ע": .ColWidth(.ColIndex("��ע")) = 2000: .ColAlignment(.ColIndex("��ע")) = flexAlignCenterCenter: .ColHidden(.ColIndex("��ע")) = True
131               .ColKey(19) = "���": .ColWidth(.ColIndex("���")) = 2000: .ColAlignment(.ColIndex("���")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���")) = True
132               .ColKey(20) = "΢����": .ColWidth(.ColIndex("΢����")) = 2000: .ColAlignment(.ColIndex("΢����")) = flexAlignCenterCenter: .ColHidden(.ColIndex("΢����")) = True
133               .ColKey(21) = "���Ա���": .ColWidth(.ColIndex("���Ա���")) = 2000: .ColAlignment(.ColIndex("���Ա���")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���Ա���")) = True
134               .ColKey(22) = "������": .ColWidth(.ColIndex("������")) = 750: .ColAlignment(.ColIndex("������")) = flexAlignCenterCenter
135               .ColKey(23) = "�����": .ColWidth(.ColIndex("�����")) = 750: .ColAlignment(.ColIndex("�����")) = flexAlignCenterCenter
136               .ColKey(24) = "������": .ColWidth(.ColIndex("������")) = 750: .ColAlignment(.ColIndex("������")) = flexAlignCenterCenter
137               .ColKey(25) = "�������": .ColWidth(.ColIndex("�������")) = 750: .ColAlignment(.ColIndex("�������")) = flexAlignCenterCenter

138               .ColKey(26) = "��ӡ��": .ColWidth(.ColIndex("��ӡ��")) = 750: .ColAlignment(.ColIndex("��ӡ��")) = flexAlignCenterCenter
139               .ColKey(27) = "��ӡʱ��": .ColWidth(.ColIndex("��ӡʱ��")) = 2000: .ColAlignment(.ColIndex("��ӡʱ��")) = flexAlignCenterCenter
140               .ColKey(28) = "�������": .ColWidth(.ColIndex("�������")) = 2000: .ColAlignment(.ColIndex("�������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("�������")) = True
141               .ColKey(29) = "��Ⱦ��": .ColWidth(.ColIndex("��Ⱦ��")) = 750: .ColAlignment(.ColIndex("��Ⱦ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("��Ⱦ��")) = True
142               .ColKey(30) = "������": .ColWidth(.ColIndex("������")) = 750: .ColAlignment(.ColIndex("������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("������")) = Not InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
143               .ColKey(31) = "����ʱ��": .ColWidth(.ColIndex("����ʱ��")) = 750: .ColAlignment(.ColIndex("����ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ʱ��")) = Not InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
144               .ColKey(32) = "������": .ColWidth(.ColIndex("������")) = 750: .ColAlignment(.ColIndex("������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("������")) = True
145               .ColKey(33) = "����ʱ��": .ColWidth(.ColIndex("����ʱ��")) = 750: .ColAlignment(.ColIndex("����ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ʱ��")) = True
                  .ColKey(34) = "������": .ColWidth(.ColIndex("������")) = 750: .ColAlignment(.ColIndex("������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("������")) = True

146           Else
147               If InStr(strTitle, "������") <= 0 Then
148                   strTitle = strTitle & ";������,750,1;����ʱ��,750,1"
149               End If
150               var_tmp = Split(strTitle, ";")
151               .Rows = 1
152               .FixedRows = 1
153               .Cols = UBound(var_tmp) + 1
154               For lngLoop = LBound(var_tmp) To UBound(var_tmp)
155                   var_SubTmp = Split(var_tmp(lngLoop), ",")
156                   .ColKey(lngLoop) = var_SubTmp(0): .ColWidth(.ColIndex(var_SubTmp(0))) = var_SubTmp(1): .ColAlignment(.ColIndex(var_SubTmp(0))) = flexAlignCenterCenter: .ColHidden(.ColIndex(var_SubTmp(0))) = Not (Val(var_SubTmp(2)) = 1)
157                   If var_SubTmp(0) = "������" Or var_SubTmp(0) = "����ʱ��" Then
158                       .ColHidden(.ColIndex(var_SubTmp(0))) = Not InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
159                   End If
160                   .ColDataType(.ColIndex("ѡ��")) = flexDTNull
161               Next
162               .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
163           End If
164           .Cell(flexcpPicture, 0, .ColIndex("���")) = mObjIco.Obj���
165           .TextMatrix(0, .ColIndex("ѡ��")) = ""
166           .Cell(flexcpPicture, 0, .ColIndex("��ӡ")) = mObjIco.Obj��ӡ
167           .TextMatrix(0, .ColIndex("����")) = "����"
168           .TextMatrix(0, .ColIndex("����")) = "����"
169           .TextMatrix(0, .ColIndex("�Ա�")) = "�Ա�"
170           .TextMatrix(0, .ColIndex("����")) = "����"
171           .TextMatrix(0, .ColIndex("������Ŀ")) = "������Ŀ"
172           .TextMatrix(0, .ColIndex("��������")) = "��������"
173           .TextMatrix(0, .ColIndex("���ʱ��")) = "���ʱ��"
174           .TextMatrix(0, .ColIndex("סԺ��")) = "סԺ��"
175           .TextMatrix(0, .ColIndex("����")) = "����"
176           .TextMatrix(0, .ColIndex("������")) = "������"
177           .TextMatrix(0, .ColIndex("�����")) = "�����"
178           .TextMatrix(0, .ColIndex("������Դ")) = "��Դ"
179           .TextMatrix(0, .ColIndex("������")) = "������"
180           .TextMatrix(0, .ColIndex("�������")) = "�������"

181           .TextMatrix(0, .ColIndex("��ӡ��")) = "��ӡ��"
182           .TextMatrix(0, .ColIndex("��ӡʱ��")) = "��ӡʱ��"
183           .TextMatrix(0, .ColIndex("�汾")) = "�汾"
184           .TextMatrix(0, .ColIndex("�������")) = "�������"
185           .TextMatrix(0, .ColIndex("������")) = "������"
186           .TextMatrix(0, .ColIndex("����ʱ��")) = "����ʱ��"
187           .TextMatrix(0, .ColIndex("������")) = "������"
188           .TextMatrix(0, .ColIndex("����ʱ��")) = "����ʱ��"
              .TextMatrix(0, .ColIndex("������")) = "������"



189           .Row = 0: .Col = .ColIndex("ѡ��"): .CellPicture = mObjIco.Objѡ��
190           .ExplorerBar = flexExSortShow

191           If blnReadData Then
192               Do Until rsTmp.EOF
193                   If lngKey <> Val(rsTmp("id") & "") Then
194                       .Rows = .Rows + 1

195                       .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""
196                       If rsTmp("��ӡ") & "" = 0 And rsTmp("����") = "�ѳ�" Then
197                           .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 1
198                       Else
199                           .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 0
200                       End If
201                       If Val(rsTmp("��ӡ") & "") <> 0 Then
                              '��Ϊ0,���Ѿ���ӡ,��ʾ��ӡͼ��
202                           .Cell(flexcpPicture, .Rows - 1, .ColIndex("��ӡ")) = mObjIco.Obj��ӡ
203                       End If
204                       .Cell(flexcpPicture, .Rows - 1, .ColIndex("�汾")) = mObjIco.Obj�°�

205                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
206                       If mintIn = 1 Then
207                           If txtName = "" Then
208                               txtName = rsTmp("����") & ""
209                           End If
210                       End If

211                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
212                       .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsTmp("�Ա�") & ""
213                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
214                       .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsTmp("������Ŀ") & ""
215                       .TextMatrix(.Rows - 1, .ColIndex("��������")) = rsTmp("��������") & ""
216                       .TextMatrix(.Rows - 1, .ColIndex("���ʱ��")) = rsTmp("���ʱ��") & ""
217                       .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = rsTmp("סԺ��") & ""
218                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
219                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsTmp("����ʱ��") & ""
220                       .TextMatrix(.Rows - 1, .ColIndex("����ID")) = rsTmp("����ID") & ""
221                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Format(rsTmp("����ʱ��") & "", "yyyy-mm-dd HH:mm:ss")
222                       .TextMatrix(.Rows - 1, .ColIndex("��ע")) = rsTmp("��ע") & ""
223                       .TextMatrix(.Rows - 1, .ColIndex("���")) = rsTmp("���") & ""
224                       .TextMatrix(.Rows - 1, .ColIndex("΢����")) = rsTmp("΢����") & ""
225                       .TextMatrix(.Rows - 1, .ColIndex("���Ա���")) = rsTmp("���Ա���") & ""
226                       .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
227                       .TextMatrix(.Rows - 1, .ColIndex("�����")) = rsTmp("�����") & ""
228                       .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
229                       .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsTmp("�������") & ""

230                       .TextMatrix(.Rows - 1, .ColIndex("��ӡ��")) = rsTmp("��ӡ��") & ""
231                       .TextMatrix(.Rows - 1, .ColIndex("��ӡʱ��")) = rsTmp("��ӡʱ��") & ""
232                       .TextMatrix(.Rows - 1, .ColIndex("�汾")) = rsTmp("�汾") & ""
233                       .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsTmp("�������") & ""
234                       .TextMatrix(.Rows - 1, .ColIndex("��Ⱦ��")) = rsTmp("�Ƿ�Ⱦ��") & ""
235                       .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
236                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsTmp("����ʱ��") & ""
237                       .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
238                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsTmp("����ʱ��") & ""
                          .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""



239                       .TextMatrix(.Rows - 1, .ColIndex("������Դ")) = chkSource(rsTmp("������Դ") & "").Caption
240                       If mlngGetPatientID > 0 Then
241                           txtPatiNo = rsTmp("סԺ��") & ""
242                       End If
243                   Else
244                       .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) & "," & rsTmp("������Ŀ") & ""
245                   End If
246                   lngKey = Val(rsTmp("id") & "")
247                   rsTmp.MoveNext
248               Loop


249               Set rsOldLisData = GetOldLisData(lngID, lngPatiID)
250               Do Until rsOldLisData.EOF
251                   If lngKey <> Val(rsOldLisData("id") & "") Then
252                       .Rows = .Rows + 1
253                       .TextMatrix(.Rows - 1, .ColIndex("id")) = rsOldLisData("id") & ""
254                       If rsOldLisData("��ӡ") & "" = "" And rsOldLisData("����") & "" = "�ѳ�" Then
255                           .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 1
256                       Else
257                           .TextMatrix(.Rows - 1, .ColIndex("ѡ��")) = 0
258                       End If
259                       If Val(rsOldLisData("��ӡ") & "") <> 0 Then
                              '��Ϊ0,���Ѿ���ӡ,��ʾ��ӡͼ��
260                           .Cell(flexcpPicture, .Rows - 1, .ColIndex("��ӡ")) = mObjIco.Obj��ӡ
261                       End If

262                       .Cell(flexcpPicture, .Rows - 1, .ColIndex("�汾")) = mObjIco.Obj�ϰ�
263                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
264                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
265                       .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsOldLisData("�Ա�") & ""
266                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
267                       .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsOldLisData("������Ŀ") & ""
268                       .TextMatrix(.Rows - 1, .ColIndex("��������")) = rsOldLisData("��������") & ""
269                       .TextMatrix(.Rows - 1, .ColIndex("���ʱ��")) = rsOldLisData("���ʱ��") & ""
270                       .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = rsOldLisData("סԺ��") & ""
271                       .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
272                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsOldLisData("����ʱ��") & ""
273                       .TextMatrix(.Rows - 1, .ColIndex("����ID")) = rsOldLisData("����ID") & ""
274                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = Format(rsOldLisData("����ʱ��") & "", "yyyy-mm-dd HH:mm:ss")
275                       .TextMatrix(.Rows - 1, .ColIndex("��ע")) = rsOldLisData("��ע") & ""
276                       .TextMatrix(.Rows - 1, .ColIndex("���")) = rsOldLisData("���") & ""
277                       .TextMatrix(.Rows - 1, .ColIndex("΢����")) = rsOldLisData("΢����") & ""
278                       .TextMatrix(.Rows - 1, .ColIndex("���Ա���")) = rsOldLisData("���Ա���") & ""
279                       .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""
280                       .TextMatrix(.Rows - 1, .ColIndex("�����")) = rsOldLisData("�����") & ""
281                       .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""
282                       .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsOldLisData("�������") & ""

283                       .TextMatrix(.Rows - 1, .ColIndex("��ӡ��")) = rsOldLisData("��ӡ��") & ""
284                       .TextMatrix(.Rows - 1, .ColIndex("��ӡʱ��")) = rsOldLisData("��ӡʱ��") & ""
285                       .TextMatrix(.Rows - 1, .ColIndex("�汾")) = rsOldLisData("�汾") & ""
286                       .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsOldLisData("�������") & ""
287                       .TextMatrix(.Rows - 1, .ColIndex("��Ⱦ��")) = ""
288                       .TextMatrix(.Rows - 1, .ColIndex("������")) = ""
289                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = ""
290                       .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""
291                       .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsOldLisData("����ʱ��") & ""
                          .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""


292                       .TextMatrix(.Rows - 1, .ColIndex("������Դ")) = chkSource(rsOldLisData("������Դ") & "").Caption
293                       If mlngGetPatientID > 0 Then
294                           txtPatiNo = rsOldLisData("סԺ��") & ""
295                       End If
296                   Else
297                       .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) & "," & rsOldLisData("������Ŀ") & ""
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
309               .Cell(flexcpSort, .FixedRows, .ColIndex("����ʱ��"), .Rows - 1, .ColIndex("����ʱ��")) = 1
310           End If


311           If mlngSelRow <> 0 And Not mlngSelRow > .Rows - 1 Then
312               .Select mlngSelRow, .ColIndex("����")
313               .ShowCell mlngSelRow, .ColIndex("����")
314           End If

              '��ȡ���
315           For lngLoop = 1 To .Rows - 1
316               .TextMatrix(lngLoop, .ColIndex("���")) = lngLoop
317               .Cell(flexcpBackColor, lngLoop, .ColIndex("���")) = &HFFEBD7
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
327       Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(ReadPatientList)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)

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
      '    strSQL = "      Select a.Id, 0 ѡ��,  Decode(a.�����, Null, 'δ��', '�ѳ�') ����, a.����, a.�Ա�, a.����, c.ҽ������ ������Ŀ, a.סԺ��, a.�����, a.����, a.����ʱ��, a.����id, a.����ʱ��, a.���ʱ��, a.��ע," & vbNewLine & _
      '           "    b.��Ŀ || ':' || b.���� ���, a.΢����걾 ΢����, 1 ���Ա���, 0 ����, a.ҽ��id ����id, a.��ӡ���� ��ӡ, a.������, a.������, a.�����, Nvl(a.������Դ, 0) ������Դ, '' ��ӡ��, '' ��ӡʱ��, 10 �汾, a.������ �������" & vbNewLine & _
      '           "     From ����걾��¼ A, ����ҽ������ B, ����ҽ����¼ C,���ű� D" & vbNewLine & _
      '           "     Where a.ҽ��id = b.ҽ��id(+) And a.ҽ��id = c.Id(+) and a.�������id = d.id  And a.����ʱ�� Between [1] And [2]   "
2         On Error GoTo GetOldLisData_Error

3         strSQL = "Select Distinct a.Id, 0 ѡ��, Decode(a.�����, Null, 'δ��', '�ѳ�') ����, a.����, a.�Ա�, a.����, c.ҽ������ ������Ŀ, a.סԺ��, a.�����, a.����," & vbNewLine & _
                  "                a.����ʱ��, a.����id, a.����ʱ��, a.���ʱ��, a.���鱸ע ��ע, f.���, a.΢����걾 ΢����, 1 ���Ա���, 0 ����, a.ҽ��id ����id, a.��ӡ���� ��ӡ, a.������,a.���˿��� �������," & vbNewLine & _
                  "                a.������, a.�����, Nvl(a.������Դ, 0) ������Դ, '' ��ӡ��, '' ��ӡʱ��, 10 �汾, a.������ �������,a.��������,a.������,a.����ʱ��,a.��ʶ�� ������" & vbNewLine & _
                  "From ����걾��¼ A, ����ҽ����¼ C, ���ű� D," & vbNewLine & _
                  "     (Select b.ҽ��id ҽ��id, f_List2str(Cast(Collect(b.��Ŀ || ':' || b.����) As t_Strlist)) ���" & vbNewLine & _
                  "       From ����걾��¼ A, ����ҽ������ B" & vbNewLine & _
                  "       Where a.ҽ��id = b.ҽ��id and a.����ʱ�� between [1] and [2] [����]" & vbNewLine & _
                  "       Group By b.ҽ��id) F" & vbNewLine & _
                  "Where a.ҽ��id = c.Id(+) And a.ҽ��id = f.ҽ��id(+) And a.�������id = d.Id and a.����ʱ�� between [1] and [2] "

4         Select Case cboPrint.ListIndex

              Case 1
                  '�Ѵ�ӡ
5                 strWhere = strWhere & "and  a.��ӡ���� is not null  "
6                 strSQL = strSQL & "and  a.��ӡ���� is not null  "

7             Case 2
                  'δ��ӡ
8                 strWhere = strWhere & " and a.��ӡ���� is null  "
9                 strSQL = strSQL & "and  a.��ӡ���� is  null  "

10        End Select

11        If chkAudit(0).value = 1 And chkAudit(1).value = 0 Then
12            strSQL = strSQL & "And a.����� Is Not Null And a.���ʱ�� Is Not Null"
13        ElseIf chkAudit(0).value = 0 And chkAudit(1).value = 1 Then
14            strSQL = strSQL & "And a.����� Is  Null And a.���ʱ�� Is  Null"
15        ElseIf chkAudit(0).value = 1 And chkAudit(1).value = 1 Then

16        End If


          '��Ⱦ��
17        Select Case Me.cboDiseases.Text
              Case "��Ⱦ��"
18                strSQL = strSQL & " and a.id = -1"
19        End Select

          '��������Դ
20        strTemp = checkboxSource()
21        If strTemp <> "" Then
22            strWhere = strWhere & " and nvl(a.������Դ,0) in (" & strTemp & ")"
23            strSQL = strSQL & " and nvl(a.������Դ,0) in (" & strTemp & ")"

24        Else
              'Ϊѡ������Դʱ������Դ����Ϊ-1
25            strWhere = strWhere & " and nvl(a.������Դ,0) in (-1)"
26            strSQL = strSQL & " and nvl(a.������Դ,0) in (-1)"

27        End If

28        If lngID = -1 Then
29            strSQL = strSQL & " and a.id = -1 "
30        End If

31        If txtName <> "" Then
32            strWhere = strWhere & " and a.���� like '" & txtName & "%' "
33            strSQL = strSQL & " and a.���� like '" & txtName & "%' "

34        End If

35        If txtDoctor.Text <> "" Then
36            strWhere = strWhere & " and a.������ like '" & txtDoctor.Text & "%'"
37            strSQL = strSQL & " and a.������ like '" & txtDoctor.Text & "%'"

38        End If

39        If Trim(txtPatiNo <> "") Then
40            If lblNo.Caption = "סԺ�š�" Then
41                strWhere = strWhere & " and a.סԺ�� = [3] "
42                strSQL = strSQL & " and a.סԺ�� = [3] "

43            ElseIf lblNo.Caption = "����š�" Then
44                strWhere = strWhere & " and a.����� = [3] "
45                strSQL = strSQL & " and a.����� = [3] "

46            ElseIf lblNo.Caption = "���š�" Then
47                strWhere = strWhere & " and a.���� = [3] "
48                strSQL = strSQL & " and a.���� = [3] "
49            ElseIf lblNo.Caption = "���￨��" Then
50                strWhere = strWhere & " and a.����ID = [7] "
51                strSQL = strSQL & " and a.����ID = [7] "
52            ElseIf lblNo.Caption = "����š�" Then
53                strWhere = strWhere & " and a.�������� = [3] "
54                strSQL = strSQL & " and a.�������� = [3] "
55            End If
56        End If
57        If cboDept <> "" Then
58            strDept = Mid(cboDept.Text, InStr(cboDept.Text, "-") + 1)
59            If InStr(strDept, "���п���") > 0 Then

60            Else
61                strSQL = strSQL & " and d.���� =[6] "
62            End If
63        End If
64        If lngPatiID = 0 And mlngPatientID <> 0 Then
65            lngPatiID = mlngPatientID
66            strWhere = strWhere & " and a.����ID = [7] "
67            strSQL = strSQL & " and a.����ID = [7] "
68        Else
69            If lngPatiID <> 0 Then
70                strWhere = strWhere & " and a.����ID = [7] "
71                strSQL = strSQL & " and a.����ID = [7] "
72            End If
73        End If
74        strSQL = Replace(strSQL, "[����]", strWhere)
75        strSQL = strSQL & " order by a.���ʱ��,a.����id,a.id"
76        If mintIn = 1 Then
77            strSQL = Replace(strSQL, "and a.����ʱ�� between [1] and [2]", "")
78        End If
79        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���벡���б�", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
                                  CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59"), txtPatiNo, mlngGetPatientID, strDepts, strDept, lngPatiID)
80        Set GetOldLisData = rsTmp


81        Exit Function
GetOldLisData_Error:
82        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(GetOldLisData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
83        Err.Clear

End Function



'----------����΢���ﱨ�洦��
Private Sub beginPrint()
    Dim strFileSource As String
    Dim lng����ID As String
    strFileSource = GetLisRptFile(mstrTag)
    lng����ID = Split(mstrTag, ";")(0)
    Call FunFastPrint(strFileSource, lng����ID)

End Sub

Private Sub picRpt_Resize()
    On Error Resume Next
    webSub.Move 0, 0, picRpt.Width, picRpt.Height
End Sub


Private Function GetLisRptFile(ByVal strTag As String) As String
'���ܣ���LIS�����ļ��鿴����ȡ��ʱ�ļ�·��
    Dim strFile As String
    Dim objFile As New FileSystemObject
    Dim strTmp As String
    Dim lng����ID As String
    Dim str������ As String
    Dim lng���� As String
    Dim varTmp As Variant
    Dim strSuffix As String '�ļ���׺��

    Screen.MousePointer = 11

    varTmp = Split(strTag, ";")
    lng����ID = varTmp(0)
    strTmp = Replace(strTag, varTmp(0) & ";" & varTmp(1) & ";", "")
    varTmp = Split(strTmp, "<sTab>")
    lng���� = varTmp(0)
    If lng���� = 0 Then
        strSuffix = "pdf"
    ElseIf lng���� = 1 Then
        strSuffix = "html"
    Else
        strSuffix = "xps"
    End If
    str������ = varTmp(1)

    strFile = objFile.GetSpecialFolder(TemporaryFolder) & "\tmpReport_" & lng����ID & "." & strSuffix
    If Not objFile.FileExists(strFile) Then
        strFile = ReadLob(100, 22, lng����ID, strFile)
        If Not objFile.FileExists(strFile) Then
            MsgBox "�ļ����ݶ�ȡʧ�ܣ�", vbInformation, "������Ϣ":
            Screen.MousePointer = 0: Exit Function
        End If
    End If
    GetLisRptFile = strFile
    Screen.MousePointer = 0
End Function


Private Sub FunFastPrint(ByVal strFile As String, ByVal lngRptID As Long)
'���ܣ�API���ÿ��ٴ�ӡPDF�ļ�
'������strFile �ļ�·��
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
'    strSQL = "Zl_ҽ����������_Print(" & lngRptID & ",0)"
'    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
   Exit Sub
errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

Private Sub WebShow(ByVal strKey As String)
'���ܣ�Web�ؼ�չʾ�ļ�
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
      '����   ��������Ϣ
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
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')' || decode(g.����ʱ��,null,'', '(' || g.����ʱ�� || ')') ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ��,B.ID ���� " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e,��������걾 F,��������ʱ�䷽�� G" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and d.id =e.���id and b.ID=F.������ϸid(+) and F.���ܷ���id=G.id(+) AND b.���id is not null and e.���id is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')' || decode(g.����ʱ��,null,'', '(' || g.����ʱ�� || ')') ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ��,B.ID ����  " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e,��������걾 F,��������ʱ�䷽�� G" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and b.ID=F.������ϸid(+) and F.���ܷ���id=G.id(+) AND e.���id is null and b.���id is null and a.id = [1] ) order by ���� desc" & vbNewLine
5             Else
6                 strSQL = "select * from ( " & vbNewLine & _
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')' ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ�� " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and d.id =e.���id and  b.���id is not null and e.���id is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')' ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ�� " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and e.���id is null and b.���id is null and a.id = [1] ) order by ���id,�������" & vbNewLine

7             End If
8             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", lngSampleID)

9         Else
10            strSQL = "   Select /*+ rule */" & vbNewLine & _
                     "  Distinct '' ���,a.�걾id, a.������Ŀid, a.����, a.�������, a.�̶���Ŀ, a.Id, a.������Ŀ, a.�ٴ�����, a.��д As Ӣ����, a.Cv," & vbNewLine & _
                     " �����־ , Decode(a.���ν��, '-', '���ԣ�-��', '+', '���ԣ�+��', '*', '*.**', a.���ν��) As ���, Rownum As ���, a.��־, a.����id, a.�걾���," & vbNewLine & _
                     "   a.����ʱ��, a.�걾���, a.�걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�����, a.סԺ��, a.��ǰ����, a.��ҳid, a.�����Χ, Nvl(g.С��λ��, 2) As С��," & vbNewLine & _
                     "    a.��������, a.��������, a.��λ," & vbNewLine & _
                     "   Trim(Replace(Replace(' ' ||" & vbNewLine & _
                     "                         Zlgetreference(a.Id, a.�걾����, Decode(a.�Ա�, '��', 1, 'Ů', 2, 0), a.��������, a.����id, a.����), ' .'," & vbNewLine & _
                     "                         '0.'), '��.', '��0.')) As �ο�, a.Od, a.Cutoff, a.Cov, a.ø���id, a.���챨��, a.���쾯ʾ, a.�������," & vbNewLine & _
                     "   A.����ο�,a.������Ŀ" & vbNewLine & _
                     "  From (Select a.Id As �걾id, b.������Ŀid, LPad(Decode(d.�������, Null, Nvl(h.����, c.����), d.�������), 4, '0') As ����," & vbNewLine & _
                     "        Nvl(b.�������, 9999) As �������, Decode(b.������Ŀid, Null, 0, 1) As �̶���Ŀ, b.������Ŀid As ID," & vbNewLine & _
                     "       c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ, d.�ٴ�����, d.��д, b.ԭʼ���, '' As �ϴν��, '' As �ϴ�ʱ��, '' As Cv," & vbNewLine & _
                     "       b.�����־, b.������ As ���ν��, d.���㹫ʽ, d.�������, Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                     "        Nvl(a.����id, -1) As ����id, Nvl(a.�걾���, 0) As �걾���, a.����ʱ��, a.�걾���," & vbNewLine & _
                     "        Decode(a.����id, Null," & vbNewLine & _
                     "                To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000'), a.�걾���) As �걾����ʾ," & vbNewLine & _
                     "        a.���鱸ע, a.����, a.�Ա�, a.����, a.�걾����, a.��������, a.�����, a.סԺ��, a.���� As ��ǰ����, a.��ҳid, d.�����Χ, d.��������, d.��������, d.��λ," & vbNewLine & _
                     "        b.Od, b.Cutoff, b.Sco As Cov, b.ø���id, d.���챨���� As ���챨��, d.���쾯ʾ�� As ���쾯ʾ, b.����ο�,h.���� ������Ŀ" & vbNewLine & _
                     " From ����걾��¼ A, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ������ĿĿ¼ H" & vbNewLine & _
                     " Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And b.������Ŀid = h.Id(+) And b.��¼���� = a.������ And" & vbNewLine & _
                     "       A.ID = [1]" & vbNewLine & _
                     " Union All" & vbNewLine & _
                     " Select a.Id As �걾id, b.������Ŀid, LPad(Decode(d.�������, Null, Nvl(h.����, c.����), d.�������), 4, '0') As ����," & vbNewLine & _
                  "       Nvl(b.�������, 9999) As �������, Decode(b.������Ŀid, Null, 0, 1) As �̶���Ŀ, b.������Ŀid As ID,"
11            strSQL = strSQL & "        c.������ || Decode(d.��д, Null, '', '(' || d.��д || ')') As ������Ŀ, d.�ٴ�����, d.��д, b.ԭʼ���, '' As �ϴν��, '' As �ϴ�ʱ��, '' As Cv," & vbNewLine & _
                     "     b.�����־,   b.������ As ���ν��, d.���㹫ʽ, d.�������, Decode(b.�����־, 3, '��', 2, '��', 1, '', 4, '�쳣', 5, '����', 6, '����', '') As ��־," & vbNewLine & _
                     "        Nvl(a.����id, -1) As ����id, Nvl(a.�걾���, 0) As �걾���, a.����ʱ��, a.�걾���," & vbNewLine & _
                     "        Decode(a.����id, Null," & vbNewLine & _
                     "                To_Char(Trunc(a.�걾��� / 10000) + 1, '0000') || '-' || To_Char(Mod(a.�걾���, 10000), '0000'), a.�걾���) As �걾����ʾ," & vbNewLine & _
                     "        a.���鱸ע, a.����, a.�Ա�, a.����, a.�걾����, a.��������, a.�����, a.סԺ��, a.���� As ��ǰ����, a.��ҳid, d.�����Χ, d.��������, d.��������, d.��λ," & vbNewLine & _
                     "        b.Od, b.Cutoff, b.Sco As Cov, b.ø���id, d.���챨���� As ���챨��, d.���쾯ʾ�� As ���쾯ʾ, b.����ο�,h.���� ������Ŀ" & vbNewLine & _
                     " From ����걾��¼ A, ����걾��¼ E, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ����������Ŀ G, ������ĿĿ¼ H" & vbNewLine & _
                     " Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And b.������Ŀid = h.Id(+) And b.��¼���� = a.������ And" & vbNewLine & _
                     "       e.Id = a.�ϲ�id And e.id = [1]) A, ����������Ŀ G" & vbNewLine & _
                     "  Where a.����id = g.����id(+) And a.Id = g.��Ŀid(+)" & vbNewLine & _
                     "  Order By a.������ĿID, a.�������"
12            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���鼼ʦվ", lngSampleID)

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
24                    lngGroupId = Val(rsTmp("���ID") & "")
25                    strGroup = rsTmp("�������") & ""
26                Else
27                    lngGroupId = Val(rsTmp("������ĿID") & "")
28                    strGroup = rsTmp("������Ŀ") & ""
29                End If
30                If lngGroupMer <> lngGroupId Then
                      '                If lngGroupRow <> 0 Then .Cell(flexcpText, lngGroupRow, 0, lngGroupRow, .Cols - 1) = strGroup & "(��" & lngNo & "��)"
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
43                    .TextMatrix(.Rows - 1, .ColIndex("���")) = lngNo
44                Next
45                lngGroupMer = lngGroupId
46                If lngGroupRow <> 0 Then .Cell(flexcpText, lngGroupRow, 0, lngGroupRow, .Cols - 1) = strGroup & "(��" & lngNo & "��)"
47                rsTmp.MoveNext
48            Loop

49            .ColWidth(.ColIndex("���")) = 500: .ColHidden(.ColIndex("���")) = False
50            .ColWidth(.ColIndex("������Ŀ")) = 1800: .ColHidden(.ColIndex("������Ŀ")) = False
51            .ColWidth(.ColIndex("���")) = 800: .ColHidden(.ColIndex("���")) = False
52            If intVal = 25 Then
53                .ColWidth(.ColIndex("�ϴ�")) = 800: .ColHidden(.ColIndex("�ϴ�")) = False
54            Else
55                .ColWidth(.ColIndex("��־")) = 800: .ColHidden(.ColIndex("��־")) = False
56            End If
57            .ColWidth(.ColIndex("��λ")) = 900: .ColHidden(.ColIndex("��λ")) = False
58            .ColWidth(.ColIndex("�ο�")) = 1000: .ColHidden(.ColIndex("�ο�")) = False

59            For i = 1 To .Rows - 1
60                If .Cell(flexcpFontBold, i, .ColIndex("���")) = True Then
61                    .RowHidden(i) = IIf(Me.chkGroup.value = 1, False, True)
62                End If
63            Next
64        End With

65        CalcReferenceColour


66        Exit Sub
ReadSampleVal_Error:
67        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(ReadSampleVal)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
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
    '����   ��������Ϣ
    Dim strErr As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo ReadSampleBacteriology_Error

    If intVal = 25 Then
        strSQL = "select b.id,b.������ || '(' || b.Ӣ���� || ')' ϸ����,a.�������� ����," & vbNewLine & _
                "       a.��ҩ����,a.���id," & vbNewLine & _
                "a.����ʱ��,a.������,a.δ���,a.��������,a.���²���,a.��ϸ��,a.�����豸,a.������," & _
                "a.����δ���,a.��������,a.��������,a.�����־,a.ϸ��ID,a.�Ƿ񾵼���,a.������� " & vbNewLine & _
                "from ���鱨��ϸ�� a,����ϸ����¼ b" & vbNewLine & _
                "where a.ϸ��id = b.id(+) and a.�걾id = [1] "

        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", lngSampleID)
    Else
        strSQL = "SELECT Distinct B.����, B.ID ϸ��id ,D.������,B.������ AS ϸ����, " & _
                    "A.������ AS ������,A.�������� as ����, d.���鱸ע,d.��ע " & _
                    "FROM ������ͨ��� A,����ϸ�� B,����걾��¼ D  " & _
                    "WHERE A.ϸ��id = B.ID And D.����� is Not null  " & _
                        "AND A.��¼���� = [1]  " & _
                        "AND D.ID=A.����걾ID  " & _
                        "AND D.ID= [2] Order by B.����"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���鼼ʦվ", mlngValueC, lngSampleID)
    End If
'    rsTmp.Sort = "�������"
    If Not vfgLoadFromRecord(VsfMicrobe, rsTmp, strErr, imgVsf) Then Exit Sub

    With VsfMicrobe
       If intVal = 25 Then
            .ColWidth(.ColIndex("ϸ����")) = 1800: .ColHidden(.ColIndex("ϸ����")) = False
            .ColWidth(.ColIndex("����")) = 1500: .ColHidden(.ColIndex("����")) = False
            .ColWidth(.ColIndex("��ҩ����")) = 1500: .ColHidden(.ColIndex("��ҩ����")) = False
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst


                Me.txtNormalMicrobe = rsTmp("������") & ""
                Me.txtNoFindMicrobe = rsTmp("δ���") & ""
                Me.txtNormalMicrobes = rsTmp("��������") & ""
                Me.chkPathopoiesiaGerm.value = IIf(rsTmp("���²���") = 1, 1, 0)
                Me.chkNoGerm.value = IIf(rsTmp("��ϸ��") = 1, 1, 0)
                Me.txtMicroscope = rsTmp("�����豸") & ""
                Me.txtMicroscopeFinded = rsTmp("������") & ""
                Me.txtMicroscopeNOFind = rsTmp("����δ���") & ""
                Me.txtMicrobePositiveComment = rsTmp("��������") & ""
                Me.txtGermComment = rsTmp("��������") & ""
                If Val(rsTmp("�Ƿ񾵼���") & "") = 0 Then
                    chkMicroscope.value = 0
                Else
                    chkMicroscope.value = 1
                End If
                If Val(rsTmp("�������") & "") = 0 Then
                    optReport(1).value = True
                Else
                    optReport(0).value = True
                End If
                optReportShow

                ReadSampleAntibiotic mlngKey, Val(rsTmp("ϸ��ID") & "")
            End If
        Else
            .ColWidth(.ColIndex("ϸ����")) = 1800: .ColHidden(.ColIndex("ϸ����")) = False
            .ColWidth(.ColIndex("������")) = 1500: .ColHidden(.ColIndex("������")) = False
            .ColWidth(.ColIndex("����")) = 1500: .ColHidden(.ColIndex("����")) = False
            If rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                txtMicrobePositiveComment = rsTmp("��ע") & ""
                txtComment = rsTmp("���鱸ע") & ""
                ReadSampleAntibiotic mlngKey, Val(rsTmp("ϸ��ID") & ""), 10
            End If
        End If
    End With


    Exit Sub
ReadSampleBacteriology_Error:
    Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(ReadSampleBacteriology)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear

End Sub
Private Sub ReadSampleAntibiotic(lngSampleID As Long, lngBacteriologyID As Long, Optional intVal As Integer = 25)
          '����           ���뿹����д��VSF
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strErr As String

1         On Error GoTo ReadSampleAntibiotic_Error

2         If intVal = 25 Then
3             strSQL = "select c.id,a.ϸ��id,c.������ || '(' || c.Ӣ���� || ')' ��������,b.���," & vbNewLine & _
                      "b.�������,b.ҩ������,b.�������,b.�ο�����,b.ҩ����ID,a.ϸ��ID " & vbNewLine & _
                      "from ���鱨��ϸ�� a,���鱨��ҩ�� b,����ҩ�� c" & vbNewLine & _
                      "where a.id = b.���id and b.ҩ��id = c.id and a.�걾ID = [1] and a.ϸ��id = [2]  "
4             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", lngSampleID, lngBacteriologyID)
5         Else
6             strSQL = "SELECT C.ϸ��ID AS Key,B.ID,B.������ AS ��������, A.��� AS ���,  " & _
                  "DECODE(A.�������,'R','R-��ҩ','I','I-�н�','S','S-����','') AS �������, " & _
                  "DECODE(A.ҩ������,1,'1-MIC',2,'2-DISK',3,'3-K-B','') As ҩ������  " & _
                   "FROM ����ҩ����� A, �����ÿ����� B,������ͨ��� C  " & _
                  "Where A.������ID = B.ID And C.ID=A.ϸ�����ID AND C.��¼����=A.��¼���� AND C.����걾id= [1] AND C.��¼����= [2] And C.ϸ��ID=[3] Order By B.����"
7            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���鼼ʦվ", lngSampleID, mlngValueC, lngBacteriologyID)
8         End If
      '    rsTmp.Sort = "�������"
9         If Not vfgLoadFromRecord(VsfAntibiotic, rsTmp, strErr, imgVsf) Then Exit Sub

10        With VsfAntibiotic
11            .ColWidth(.ColIndex("��������")) = 1800: .ColHidden(.ColIndex("��������")) = False
12            .ColWidth(.ColIndex("���")) = 1500: .ColHidden(.ColIndex("���")) = False
13            .ColWidth(.ColIndex("�������")) = 1500: .ColHidden(.ColIndex("�������")) = False
14        End With


15        Exit Sub
ReadSampleAntibiotic_Error:
16        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(ReadSampleAntibiotic)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear

End Sub
Private Function LoadContrastDBWriteVSF(VSFList As VSFlexGrid, lngSampleID As Long, lngPatientID As Long, SampleReportDate As Date, _
                                        intMaxDay As Integer, Optional strErr As String) As Boolean
      '����                   �����ݿ��ж����ȶ�����д��VSF��
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim lngItemid As Long
          Dim intCol As Integer
          Dim dblTmp As Double
          Dim blnTre As Boolean       '�Ƿ�����������걾

1         On Error GoTo LoadContrastDBWriteVSF_Error



2         If mintVer = 25 Then
3             blnTre = gobjLiscomlib.IsTre(lngSampleID)

4             If blnTre Then
5                 strSQL = "Select b.id, b.������, b.Ӣ����, b.��λ, a.id ����, c.����ʱ��, a.������, e.����ʱ��, b.���챨����, b.�������, a.�����־" & vbNewLine & _
                         "   From ���鱨����ϸ A, ����ָ�� B, ���鱨���¼ C, ��������걾 D, ��������ʱ�䷽�� E" & vbNewLine & _
                         "   Where A.��ĿID = B.ID And A.�걾ID = C.ID And A.ID = D.������ϸid And D.���ܷ���id = e.ID And A.�걾ID = [1]" & vbNewLine & _
                         "   Order By a.id Desc"
6                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ȶ�����", lngSampleID)
7             Else
8                 strSQL = "Select " & vbNewLine & _
                         " B.Id, B.������, B.Ӣ����, B.��λ, A.����, A.����ʱ��, A.������, B.���챨����, B.�������" & vbNewLine & _
                           "From (Select B.��Ŀid ������Ŀid, B.����, B.����ʱ��, B.������" & vbNewLine & _
                         "       From (Select A.Id ����, A.����id, A.�걾����, A.����ʱ��, B.��Ŀid" & vbNewLine & _
                         "              From ���鱨���¼ A, ���鱨����ϸ B" & vbNewLine & _
                         "              Where A.Id = B.�걾id And A.Id = [1] and b.������ is not null ) A," & vbNewLine & _
                         "            (Select A.Id ����, A.����id, A.�걾����, A.����ʱ��, B.��Ŀid, B.������" & vbNewLine & _
                         "              From ���鱨���¼ A, ���鱨����ϸ B" & vbNewLine & _
                         "              Where A.Id = B.�걾id And A.����id = [2] And ����ʱ��+0 Between [3] And [4] and a.id <= [1] and b.������ is not null ) B" & vbNewLine & _
                         "       Where A.����id = B.����id And A.��Ŀid + 0 = B.��Ŀid And Nvl(A.�걾����, 0) = Nvl(B.�걾����, 0) ) A, ����ָ�� B" & vbNewLine & _
                           "Where A.������Ŀid = B.Id" & vbNewLine & _
                           "Order By LPad(B.�������, 10, '0'),b.id, A.���� desc "

9                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ȶ�����", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                         CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
10            End If
11        Else
12            strSQL = "    Select " & vbNewLine & _
                     "       i.Id, i.���� As ������, v.��д As Ӣ����, i.���㵥λ As ��λ, a.����, a.����ʱ��, a.������, v.���챨����, v.�������" & vbNewLine & _
                     "       From (Select b.������Ŀid, b.����, b.����ʱ��, b.������" & vbNewLine & _
                     "              From (Select a.Id ����, a.����id, a.�걾����, a.���ʱ�� ����ʱ��, b.������Ŀid, b.������" & vbNewLine & _
                     "                     From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                     "                     Where a.Id = b.����걾id And a.Id = [1] And a.����id = [2] And b.������ Is Not Null) A," & vbNewLine & _
                     "                   (Select a.Id ����, a.����id, a.�걾����, a.���ʱ�� ����ʱ��, b.������Ŀid, b.������" & vbNewLine & _
                     "                     From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                     "                     Where a.Id = b.����걾id And a.Id < [1] And a.����id = [2]  And  a.����ʱ��+0 Between [3] And [4]  And b.������ Is Not Null) B" & vbNewLine & _
                     "              Where a.����id = b.����id And a.������Ŀid + 0 = b.������Ŀid) A, ������Ŀ V, ���鱨����Ŀ R, ������ĿĿ¼ I" & vbNewLine & _
                     "       Where A.������Ŀid = v.������Ŀid And A.������Ŀid = r.������Ŀid And r.������Ŀid = i.ID And i.�����Ŀ <> 1"
13            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ȶ�����", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                     CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
14        End If
15        vfgSetting 0, VSFList
16        With VSFList

17            .Rows = 1
18            .Cols = 1
19            .FixedRows = 1
              '        .FixedCols = 1
20            .TextMatrix(0, 0) = "������Ŀ": .ColWidth(0) = 2500
21            Do Until rsTmp.EOF
22                If lngItemid <> rsTmp("ID") Then
23                    .Rows = .Rows + 1
24                    intCol = 0
25                    If .Cols - 1 < intCol Then
26                        .Cols = .Cols + 1
27                        .ColWidth(intCol) = 1500
28                    End If

29                    If intCol = 0 Then
                          'д����Ŀ
30                        .TextMatrix(.Rows - 1, intCol) = rsTmp("������") & "(" & rsTmp("Ӣ����") & ")"

31                    End If
32                    intCol = intCol + 1
33                    If .Cols - 1 < intCol Then
34                        .Cols = .Cols + 1
35                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
36                        If Not blnTre Then
37                            .TextMatrix(0, intCol) = "����"
38                        Else
39                            .TextMatrix(0, intCol) = rsTmp("����ʱ��") & ""
40                        End If
41                    End If
                      'д������
42                    .TextMatrix(.Rows - 1, intCol) = rsTmp("������") & ""
43                Else
44                    intCol = intCol + 1
45                    If .Cols - 1 < intCol Then
46                        .Cols = .Cols + 1
47                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
48                        If blnTre Then
49                            .TextMatrix(0, intCol) = rsTmp("����ʱ��") & ""
50                        Else
51                            .TextMatrix(0, intCol) = "��" & intCol - 1 & "��"
52                        End If
53                        dblTmp = Val(CalcVolatility(.TextMatrix(.Rows - 1, 1), .TextMatrix(.Rows - 1, intCol)))
54                        If dblTmp <> 0 And Val(rsTmp("���챨����") & "") <> 0 Then
55                            If dblTmp > Val(rsTmp("���챨����") & "") Then
56                                .Cell(flexcpBackColor, .Rows - 1, intCol) = RGB(248, 194, 169)
57                            End If
58                        End If
59                    End If
                      'д������
60                    .TextMatrix(.Rows - 1, intCol) = rsTmp("������") & ""
61                End If
62                lngItemid = rsTmp("ID")
63                rsTmp.MoveNext
64            Loop
65        End With

66        LoadContrastDBWriteVSF = True


67        Exit Function
LoadContrastDBWriteVSF_Error:
68        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(LoadContrastDBWriteVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
69        Err.Clear

End Function

Private Sub ReadHistorData()
    '����           �������ε�����
    Dim strErr As String
    Call LoadContrastDBWriteVSF(VSFContrast, mlngKey, mlngPatientID, mdteReportDate, 60, strErr)
End Sub
Private Function CalcVolatility(strCalcA As String, strCalcB As String) As String
    '���������

    On Error Resume Next

    If strCalcA = "" Or strCalcB = "" Then
        CalcVolatility = ""
        Exit Function
    End If
    If Val(strCalcA) = 0 Or Val(strCalcB) = 0 Then
        CalcVolatility = ""
    End If

    '����
    CalcVolatility = (Val(strCalcB) - Val(strCalcA)) / Val(strCalcA) * 100
End Function
Private Function LoadVSFContrastToCht(VSFList As VSFlexGrid, chtObj As Chart2D, intRow As Integer, intType As Integer, strErr As String) As Boolean
          '����           ��VSF��������д��Cht�ؼ�
          Dim intCol As Integer
          Dim dblMax As Double

1         On Error GoTo LoadVSFContrastToCht_Error

2         chtObj.ChartGroups(1).Data.NumSeries = 0
3         With chtObj.ChartGroups(1)
4             .ChartType = oc2dTypePlot  '����
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
15            .Axes("X").AnnotationMethod = oc2dAnnotateValueLabels   '��������ʾֵ��ʾ

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
                  Case 1              '������
45                    .Axes("Y").DataMax = Abs(dblMax)
46                    .Axes("Y").DataMin = Abs(dblMax) * -1
47                    .Axes("Y").Origin = 0
48                Case 2              '���ֵ
49                    .Axes("Y").DataMax = Abs(dblMax) + Abs(dblMax) / 100 * 10
50                    .Axes("Y").DataMin = 0
51                    .Axes("Y").Origin = 0
52            End Select
53        End With
54        LoadVSFContrastToCht = True


55        Exit Function
LoadVSFContrastToCht_Error:
56        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(LoadVSFContrastToCht)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
57        Err.Clear

End Function
Private Sub ReadContrastToVsf()
    '����       �������αȶԵ�VSF
    Dim strErr As String

    Me.VSFContrast.Rows = 1: Me.VSFContrast.Rows = 2


    'û�в���IDʱ�˳�
    If mlngPatientID = 0 Then Exit Sub

    Call LoadContrastDBWriteVSF(Me.VSFContrast, mlngKey, mlngPatientID, mdteReportDate, Val(txtMaxDay), strErr)
    Call VSFContrast_SelChange
End Sub

Private Sub InitFace()
    '����           ��ʼ������
    '========================================��ʾ��ɫ����============================================
    '��ʾ��ɫ����
    mSampleShowColour.���� = &H80000005
    mSampleShowColour.ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾƫ����ɫ", 2500, 2500, "8438015"))
    mSampleShowColour.ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾƫ����ɫ", 2500, 2500, "8454143"))
    mSampleShowColour.��ʾƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ��ʾƫ����ɫ", 2500, 2500, "255"))
    mSampleShowColour.��ʾƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ��ʾƫ����ɫ", 2500, 2500, "255"))
    mSampleShowColour.����ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ����ƫ����ɫ", 2500, 2500, "65280"))
    mSampleShowColour.����ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ����ƫ����ɫ", 2500, 2500, "12648384"))
    mSampleShowColour.�쳣 = Val(ComGetPara(Sel_Lis_DB, "��ʾ�쳣��ɫ", 2500, 2500, "16576"))

    '�Ƿ���ʾδ����ѡ��
    chkAudit(0).Visible = InStr(";" & mstrPrivs & ";", ";�鿴δ������;") > 0
    If chkAudit(0).Visible = False Then chkAudit(0).value = 1
    chkAudit(1).Visible = InStr(";" & mstrPrivs & ";", ";�鿴δ������;") > 0
    If chkAudit(1).Visible = False Then chkAudit(1).value = 0
End Sub
Private Function GetValColour(intValType As Integer) As Double
    '����               �����Ӧ�Ľ������1-������2-ƫ�͡�3-ƫ�ߡ�4-����(�쳣)��5-��ʾ���ޡ�6-��ʾ���ޡ�7-�������ޡ�8-��������
    '����               ��Ӧ����ɫ
    Select Case intValType
        Case 1, 0
            GetValColour = mSampleShowColour.����
        Case 2
            GetValColour = mSampleShowColour.ƫ��
        Case 3
            GetValColour = mSampleShowColour.ƫ��
        Case 4
            GetValColour = mSampleShowColour.�쳣
        Case 5
            GetValColour = mSampleShowColour.��ʾƫ��
        Case 6
            GetValColour = mSampleShowColour.��ʾƫ��
        Case 7
            GetValColour = mSampleShowColour.����ƫ��
        Case 8
            GetValColour = mSampleShowColour.����ƫ��
    End Select
End Function

Private Sub CalcReferenceColour()
          '����           ����������ɫ
          Dim intCol As Integer
          Dim intRow As Integer

1         On Error GoTo CalcReferenceColour_Error

2         With vsfCenter
3             For intRow = 1 To .Rows - 1
4                 If .ColIndex("id") <> -1 Then
5                     If Val(.TextMatrix(intRow, .ColIndex("id"))) <> 0 Then
      '                    If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�汾"))) = 25 Then
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
12                        .Cell(flexcpBackColor, intRow, .ColIndex("���"), intRow, .ColIndex("���")) = GetValColour(Val(.TextMatrix(intRow, .ColIndex("�����־"))))
13                    End If
14                End If
15            Next
16        End With


17        Exit Sub
CalcReferenceColour_Error:
18        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(CalcReferenceColour)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
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
7                             If intCol <> .ColIndex("���") Then
8                                 If OldRow <> 0 Then .Cell(flexcpBackColor, OldRow, intCol, OldRow, intCol) = mSampleShowColour.����
9                                 .Cell(flexcpBackColor, NewRow, intCol, NewRow, intCol) = &HFFEBD7
10                            End If
11                        Next
12                        If .Col = .ColIndex("���") Then
13                            .BackColorSel = GetValColour(Val(.TextMatrix(NewRow, .ColIndex("�����־"))))
14                        Else
15                            .BackColorSel = &HFFEBD7
16                        End If

17                        txtSignificance.Text = .TextMatrix(.Row, .ColIndex("�ٴ�����"))
18                    End If
19                End If
20            End With
21        End If


22        Exit Sub
vsfCenter_AfterRowColChange_Error:
23        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(vsfCenter_AfterRowColChange)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
24        Err.Clear
End Sub

Private Sub vsfLeft_AfterSort(ByVal Col As Long, Order As Integer)
    vsfLeft_SelChange
End Sub

Private Sub vsfLeft_Click()
    With Me.vsfLeft
        If .Row < 1 Then Exit Sub
        Me.cbrthis.FindControl(, conFun_Sample_Auditing).Enabled = IIf(Val(.TextMatrix(.Row, .ColIndex("��Ⱦ��"))) = 1, True, False) _
                                                                    And .TextMatrix(.Row, .ColIndex("������")) = "" And .TextMatrix(.Row, .ColIndex("����ʱ��")) = ""
        Me.cbrthis.FindControl(, conFun_Sample_unAuditing).Enabled = IIf(Val(.TextMatrix(.Row, .ColIndex("��Ⱦ��"))) = 1, True, False) _
                                                                    And .TextMatrix(.Row, .ColIndex("������")) <> "" And .TextMatrix(.Row, .ColIndex("����ʱ��")) <> ""
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
            If lngCol = .ColIndex("ѡ��") Then
                If .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 1 Then
                    .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 2
                Else
                    .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 1
                End If
            End If
        End If

        '�������Ҽ�,�������ô���
        If Button = 2 Then
            If lngRow = 0 Then
                Call GetCursorPos(Point)
                strTitle = SetVsfColHiden(Me, Me.vsfLeft, Point.X * 15, Point.Y * 15, "����������ʾ��", 2500, 1013, "�������,id")
                If strTitle <> "" Then
                    SaveDBLog 18, 6, 0, "���鱨���ѯ", "���ñ���е���ʾ������:" & strTitle, 2500, "�ٴ�ʵ���ҹ���"
                    Call ReadPatientList(1)
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfLeft_SelChange()
1         On Error GoTo vsfLeft_SelChange_Error

2         With vsfLeft
3             If .ColIndex("id") <> -1 And .ColIndex("����ID") <> -1 And .ColIndex("����ʱ��") <> -1 Then
4                 If Val(.TextMatrix(.Row, .ColIndex("ID"))) > 0 Then
5                     If Val(.TextMatrix(.Row, .ColIndex("ID"))) <> mlngKey Then
6                         Call clsAllEdit
                          mstrTag = ""
7                         mlngKey = Val(.TextMatrix(.Row, .ColIndex("ID")))
8                         mlngPatientID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
9                         mdteReportDate = .TextMatrix(.Row, .ColIndex("����ʱ��"))
10                        txtComment = .TextMatrix(.Row, .ColIndex("��ע"))
11                        txtDiagnose = .TextMatrix(.Row, .ColIndex("���"))
12                        mlngValueC = .TextMatrix(.Row, .ColIndex("�������"))
13                        mintVer = .TextMatrix(.Row, .ColIndex("�汾"))
14                        If Val(.TextMatrix(.Row, .ColIndex("΢����"))) = 1 Then



15                            If Val(.TextMatrix(.Row, .ColIndex("���Ա���"))) = 1 Then
16                                picGeneral.Visible = False
17                                picMicrobePositive.Visible = True
18                                PicNegative.Visible = False
                                  picRpt.Visible = False
19                                If Val(.TextMatrix(.Row, .ColIndex("�汾"))) = 25 Then
20                                    ReadSampleBacteriology mlngKey, 25
21                                Else
22                                    ReadSampleBacteriology mlngKey, 10
23                                End If
24                            ElseIf Val(.TextMatrix(.Row, .ColIndex("���Ա���"))) = 3 Then
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
34                            If Val(.TextMatrix(.Row, .ColIndex("�汾"))) = 25 Then
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
46        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(vsfLeft_SelChange)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
47        Err.Clear
End Sub

Private Sub findThirdReport(ByVal strAdvice As String)
    Dim strSQL As String
    Dim rsTemp As Recordset
    '����LIS����
    Dim strTag As String

    mstrTag = ""
    strSQL = "select b.id as ����ID,b.������,b.������||','||To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI') as �ĵ�����,c.ҽ��ID,b.����,b.��ӡ���� from ����ҽ����¼ a, ҽ���������� b,����ҽ������ c where b.id=c.����id and a.id=c.ҽ��id and c.����id is not null and b.���� in (0,2) and a.id =[1]"

    Set rsTemp = OpenSQLRecord(Sel_His_DB, strSQL, Me.Caption, strAdvice)
    If rsTemp.RecordCount > 0 Then
        strTag = rsTemp!����ID & ";" & rsTemp!ҽ��id & ";" & rsTemp!���� & "<sTab>" & rsTemp!������
        mstrTag = strTag
        Call WebShow(strTag)
    End If


End Sub

Private Sub ImageTypeSet(intCount As Integer, Optional blnNoRead As Boolean = False)
    '����           ͼ���Ű�(���9��ͼ)
    Dim intloop As Integer
    '����������ͼ�����а�����
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
        '��4��ͼ��������
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
          '����               ���뵱ǰ�걾��ͼ�ε�Cht
          Dim strChart(0 To 8) As String
          Dim strErr As String
          Dim intloop As Integer

          '���Ű�
1         On Error GoTo ReadImages_Error

2         Call ImageTypeSet(9, True)
          '����ͼ������
3         If ReadSampleImage(lngSampleID, strChart, strErr, intVal) = False Then
4             Exit Sub
5         End If
6         For intloop = 0 To 8
7             If strChart(intloop) <> "" Then
8                 chtPic(intloop).Load (strChart(intloop))
9             End If
10        Next
          '����������Ű�
11        Call ImageTypeSet(9)


12        Exit Sub
ReadImages_Error:
13        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(ReadImages)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear

End Sub

Private Sub RefreshTab(Index As Integer)
    '����           ˢ�¶�Ӧ��ҳ
    Select Case Index
        Case 1
            ReadHistorData
        Case 2
            ReadImages mlngKey, mintVer
    End Select
End Sub
Public Function PrintReport(objFrm As Object, lngSampleID As Long, Optional byRunMode As Byte = 2, Optional strErr As String) As Boolean
          '����       ��ӡ����
          Dim intCount As Integer
          Dim strNO As String
          Dim intSel As Integer
          Dim strChart(0 To 8) As String
          Dim strSQL As String
          Dim strTmp As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset

1         On Error GoTo PrintReport_Error

2         strSQL = "select b.id ����id ,b.���� ��������,b.�������,a.������Դ,a.����ʱ��,a.���Ա���,a.�걾��� from ���鱨���¼ a,����������¼ b where a.����id = b.id and a.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ӡ", lngSampleID)

4         If rsTmp.RecordCount = 0 Then Exit Function

5         strSQL = "select id,����,����,���ﵥ��,סԺ����,��쵥��,Ժ�ⵥ��,�����ʽ,סԺ��ʽ,����ʽ,Ժ���ʽ,��ʽ����," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(���ﵥ��, '00000')) || '-2' ���ﵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ����, '00000')) || '-2' סԺ���ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(��쵥��, '00000')) || '-2' ��쵥�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ�ⵥ��, '00000')) || '-2' Ժ�ⵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(�����ʽ, '00000')) || '-2' �����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ��ʽ, '00000')) || '-2' סԺ��ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(����ʽ, '00000')) || '-2' ����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ���ʽ, '00000')) || '-2' Ժ���ʽ��" & vbNewLine & _
                      "from ����������¼ where id = [1] "

6         Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", Val(rsTmp("����ID") & ""))


7         rsReportFormat.Filter = "id=" & Val(rsTmp("����ID") & "")
8         If Val(rsTmp("�������")) = 1 Then
9             If Val(rsTmp("���Ա���") & "") = 1 Then
                  '����
10                intSel = 0
11            Else
                  '����
12                intSel = 1
13            End If
14        Else
15            intCount = GetSampleValCount(lngSampleID)
              'û�н��ʱ��ʾ
16            If intCount = 0 Then
17                Exit Function
18            End If
19            If rsReportFormat.RecordCount > 0 Then
20                If Val(rsReportFormat("��ʽ����") & "") > 0 Then
21                    If intCount > Val(rsReportFormat("��ʽ����") & "") Then
22                        intSel = 0
23                    Else
24                        intSel = 1
25                    End If
26                End If
27            Else
28                intSel = 0
29            End If

30        End If
31        Select Case Val(rsTmp("������Դ") & "")
              Case 1
32                If intSel = 0 Then
33                    strNO = rsReportFormat("���ﵥ�ݺ�")
34                Else
35                    strNO = rsReportFormat("�����ʽ��")
36                End If
37            Case 2
38                If intSel = 0 Then
39                    strNO = rsReportFormat("סԺ���ݺ�")
40                Else
41                    strNO = rsReportFormat("סԺ��ʽ��")
42                End If
43            Case 3
44                If intSel = 0 Then
45                    strNO = rsReportFormat("סԺ���ݺ�")
46                Else
47                    strNO = rsReportFormat("סԺ��ʽ��")
48                End If
49            Case 4
50                If intSel = 0 Then
51                    strNO = rsReportFormat("Ժ�ⵥ�ݺ�")
52                Else
53                    strNO = rsReportFormat("Ժ���ʽ��")
54                End If
55            Case Else
56                If intSel = 0 Then
57                    strNO = rsReportFormat("���ﵥ�ݺ�")
58                Else
59                    strNO = rsReportFormat("�����ʽ��")
60                End If
61        End Select
62        If byRunMode = 3 Then
63            If strNO <> "" Then
64                FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
65            End If
66        Else
             '��ͼ��
67            strTmp = "��ʼ����ͼ��:" & Now & vbCrLf
68            If ReadSampleImage(lngSampleID, strChart, strErr) = False Then
69                Exit Function
70            End If
71            strTmp = strTmp & "����ͼ�����:" & Now & vbCrLf

72            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "�걾ID=" & lngSampleID, "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), _
                      "ͼ��4=" & strChart(3), "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                      "ͼ��9=" & strChart(8), byRunMode
73            strTmp = strTmp & "��ӡ���:" & Now & vbCrLf

              '������˹��ı걾��ʶ
74            strSQL = "Zl_���鱨���ӡ_Edit(1," & lngSampleID & ")"
75            Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
76            strTmp = strTmp & "��ɴ�ӡ:" & Now

77            SaveDBLog 18, 6, lngSampleID, "��ӡ", "�����ӡ", 2500, "�ٴ�ʵ���ҹ���"
78        End If

79        PrintReport = True

          '����ˢ�¿��ڸſ��Ѵ�ӡ��ǩ����
80        Call SendMessage("RefreshDeptSurvey7")


81        Exit Function
PrintReport_Error:
82        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(PrintReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
83        Err.Clear
End Function

Private Sub GetDept(Optional intType As Integer, Optional ByVal strInfo As String)
          '����               ������һ���
          '����               intType 0=���� 1=����
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim blnFind As Boolean

1         On Error GoTo GetDept_Error

2         If strInfo <> "" Then blnFind = True
3         If intType = 0 Then
4             strSQL = "Select Distinct a.����, a.����, a.���� From ���ű� A, ��������˵�� B" & _
                      " Where a.Id = b.����id And a.����ʱ�� Is Not Null And a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd') And" & _
                      " (b.�������� = '�ٴ�' Or b.�������� = '����' Or b.�������� = '����' Or b.�������� = '����')" & _
                      IIf(blnFind, "AND ( A.���� like [1] or A.���� like [2] or A.���� like [2])", "") & "order by a.����"

5         Else
6             strSQL = ""
7         End If
8         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�������", IIf(IsNumeric(strInfo), strInfo, ""), UCase(strInfo))
9         If rsTmp.EOF Then Exit Sub
10        With cboDept
11            .Clear
12            If blnFind = False Then
13                If InStr(mstrPrivs, "���п���") > 0 Then
14                    .AddItem "00-���п���"
15                End If
16            End If
17            Do Until rsTmp.EOF
18                .AddItem Trim(rsTmp("����")) & "-" & Trim(rsTmp("����")) & ""
19                rsTmp.MoveNext
20            Loop
21            .ListIndex = 0
22        End With


23        Exit Sub
GetDept_Error:
24        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(GetDept)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
25        Err.Clear
End Sub


Private Sub CboFind(objcbo As ComboBox, lngID As Long)
    '����           �ҵ�cbo��Ӧ��id
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

    '�Ƿ���ʾ��Ⱦ��ɸѡ��
    picDiseases.Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
    Me.cboDiseases.ListIndex = 2

    With Me.cbrthis
        .FindControl(, conFun_Sample_Auditing).Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
        .FindControl(, conFun_Sample_unAuditing).Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
    End With
    GetDept 0
    Exit Function
errH:
    strErr = "������(ShowMe),������Ϣ:" & Err.Number & " " & Err.Description
End Function

Public Function BHShowMe(lngMain As Long, Optional strErr As String) As Boolean
    On Error GoTo errH
    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 1013)


    gobjLiscomlib.ShowChildWindow Me.hWnd, lngMain
    BHShowMe = True

     '�Ƿ���ʾ��Ⱦ��ɸѡ��
    picDiseases.Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
    Me.cboDiseases.ListIndex = 2

    With Me.cbrthis
        .FindControl(, conFun_Sample_Auditing).Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
        .FindControl(, conFun_Sample_unAuditing).Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
    End With
    GetDept 0

    Exit Function
errH:
    strErr = "������(ShowMe),������Ϣ:" & Err.Number & " " & Err.Description
End Function

Private Sub BatchPrint()
          '����   ������ӡ
          Dim intRow As Integer
          Dim strMsgShow As String
          Dim blnPrint As Boolean '����ѡ�������ʱ,�Ƿ��Ѿ���ӡ���ѳ��ı��� TRUE=�Ѵ�ӡ�ѳ�����,False=δ��ӡ�ѳ�����

1         On Error GoTo BatchPrint_Error

2         If checkDiseases = False Then
3             Exit Sub
4         End If

5         With vsfLeft

              '�ж��Ƿ���δ���ı���
6             For intRow = 1 To .Rows - 1
7                 If .Cell(flexcpChecked, intRow, .ColIndex("ѡ��"), intRow, .ColIndex("ѡ��")) = 1 Then
8                     If Trim(.TextMatrix(intRow, .ColIndex("����"))) = "�ѳ�" Then
9                         blnPrint = True
10                    End If
11                    If Trim(.TextMatrix(intRow, .ColIndex("����"))) = "δ��" Then
12                        strMsgShow = "����δ��,�����ĵȴ�"
13                    End If
14                End If
15            Next
16            If blnPrint = True And strMsgShow <> "" Then
17                strMsgShow = "��δ������,�����ĵȴ�"
18            End If


19            For intRow = 1 To .Rows - 1
20                If .Cell(flexcpChecked, intRow, .ColIndex("ѡ��"), intRow, .ColIndex("ѡ��")) = 1 Then
21                    If Val(.TextMatrix(intRow, .ColIndex("id"))) > 0 And Trim(.TextMatrix(intRow, .ColIndex("����"))) = "�ѳ�" Then
22                        If .TextMatrix(intRow, .ColIndex("�汾")) = 25 Then
23                            If Val(.TextMatrix(intRow, .ColIndex("��Ⱦ��"))) = 1 Then
24                                If .TextMatrix(intRow, .ColIndex("������")) <> "" And .TextMatrix(intRow, .ColIndex("����ʱ��")) <> "" Then
25                                    PrintReport Me, Val(.TextMatrix(intRow, .ColIndex("id")))
26                                End If
27                            Else
28                                PrintReport Me, Val(.TextMatrix(intRow, .ColIndex("id")))
29                            End If
30                        Else
31                            PtintOldReport Me, Val(.TextMatrix(intRow, .ColIndex("id"))), Val(.TextMatrix(intRow, .ColIndex("����id")))
32                        End If
33                    End If
34                End If
35            Next
36        End With

37        If strMsgShow <> "" Then
38            MsgBox strMsgShow, vbInformation, Me.Caption
39        End If

          'ˢ�½���
40        Call ReadPatientList(1)
41        Me.txtPatiNo.SetFocus
42        Call txtPatiNo_GotFocus


43        Exit Sub
BatchPrint_Error:
44        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(BatchPrint)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
45        Err.Clear
End Sub

Private Function checkDiseases() As Boolean
          '��ӡ֮ǰ����Ƿ����û�и��˵Ĵ�Ⱦ������,����,���ӡ�ж�
          Dim intRow As Integer
          Dim blnFindDiseases As Boolean '�Ƿ���ҵ�Ϊ���˵Ĵ�Ⱦ������

1         On Error GoTo checkDiseases_Error

2         blnFindDiseases = False
3         With Me.vsfLeft
4             For intRow = 1 To .Rows - 1
5                 If .Cell(flexcpChecked, intRow, .ColIndex("ѡ��"), intRow, .ColIndex("ѡ��")) = 1 Then
6                     If Val(.TextMatrix(intRow, .ColIndex("id"))) > 0 Then
7                         If .TextMatrix(intRow, .ColIndex("�汾")) = 25 Then
8                             If Val(.TextMatrix(intRow, .ColIndex("��Ⱦ��"))) = 1 Then
9                                 If .TextMatrix(intRow, .ColIndex("������")) = "" Or .TextMatrix(intRow, .ColIndex("����ʱ��")) = "" Then
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
20            MsgBox "���ִ���δ���˵Ĵ�Ⱦ������,��ӡ�ж�", vbInformation, Me.Caption
21            checkDiseases = False
22            Exit Function
23        End If
24        checkDiseases = True


25        Exit Function
checkDiseases_Error:
26        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(checkDiseases)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear

End Function

Private Function PtintOldReport(objFrm As Object, lngSampleID As Long, Optional lngPaintID As Long, Optional byRunMode As Byte = 2, Optional strErr As String) As Boolean
          '��ӡ����
          Dim strReportCode As String
          Dim strReportParaNo As String
          Dim bytReportParaMode As Byte
          Dim rsTmp As New ADODB.Recordset
          Dim blnCurrMoved As Boolean
          Dim lngҽ��ID As Long, lng���ͺ� As Long
          Dim strSQL As String
          Dim strChart(0 To 8) As String

1         On Error GoTo PtintOldReport_Error

2         strSQL = "select ���ͺ�, a.ҽ��id from ����ҽ������ a , ����ҽ����¼ b,����걾��¼  c where b.id = a.ҽ��id and  a.ҽ��id =c.ҽ��id  and c.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�����ӡ", lngSampleID)
4         If rsTmp.EOF = False Then
5             lng���ͺ� = Val("" & rsTmp("���ͺ�"))
6             lngҽ��ID = Val("" & rsTmp("ҽ��id"))
7         End If

8         If GetReportCode(lngҽ��ID, lng���ͺ�, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
9             If byRunMode = 3 Then
10                FunReportPrintSetHis gcnHisOracle, 100, strReportCode, objFrm
11            Else
12                If ReadSampleImage(lngSampleID, strChart, strErr, 10) = False Then
13                    Exit Function
14                End If
15                Call FunReportOpenHis(gcnHisOracle, 100, strReportCode, objFrm, "NO=" & strReportParaNo, "����=" & bytReportParaMode, "ҽ��ID=" & lngҽ��ID, _
                                      "����ID=" & lngPaintID, "�걾ID=" & lngSampleID, _
                                      "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), "ͼ��4=" & strChart(3), _
                                      "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                                      "ͼ��9=" & strChart(8), byRunMode)

                  '������˹��ı걾��ʶ
16                strSQL = "Zl_����걾��¼_�걾�ʿ�(" & lngSampleID & ",'',1)"
17                Call ComExecuteProc(Sel_His_DB, strSQL, "��ӡ�걾")
19            End If
20        End If


21        Exit Function
PtintOldReport_Error:
22        Call writeErrLog("zl9LisInsideComm", "frmPatientReprotFind", "ִ��(PtintOldReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear
End Function



Private Sub VsfMicrobe_SelChange()
    With VsfMicrobe
        If .ColIndex("id") <> -1 And .ColIndex("ϸ��ID") <> -1 Then
            If Val(.TextMatrix(.Row, .ColIndex("ϸ��ID"))) > 0 Then
                If mintVer = 25 Then
                    ReadSampleAntibiotic mlngKey, Val(.TextMatrix(.Row, .ColIndex("ϸ��ID")))
                    txtMicrobePositiveComment.Text = .TextMatrix(.Row, .ColIndex("��������"))
                Else
                    ReadSampleAntibiotic mlngKey, Val(.TextMatrix(.Row, .ColIndex("ϸ��ID"))), 10
                End If
            End If
        End If
    End With
End Sub

Private Sub clsAllEdit()
    '�������������ʾ
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
    'ѡ������Դ�ַ���
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


    '�Ƿ���ʾ��Ⱦ��ɸѡ��
    picDiseases.Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
    Me.cboDiseases.ListIndex = 2

    With Me.cbrthis
        .FindControl(, conFun_Sample_Auditing).Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
        .FindControl(, conFun_Sample_unAuditing).Visible = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
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
    strErr = "������(ShowMe),������Ϣ:" & Err.Number & " " & Err.Description
End Function

Private Sub selAllText(ByVal objCrl As Object)
    With objCrl
        .SelStart = 0
        .SelLength = Len(objCrl.Text)
    End With
End Sub


