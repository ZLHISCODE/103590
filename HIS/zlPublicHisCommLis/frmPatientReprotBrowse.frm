VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPatientReprotBrowse 
   AutoRedraw      =   -1  'True
   Caption         =   "���������"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17085
   Icon            =   "frmPatientReprotBrowse.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   17085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
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
            Caption         =   "�������"
            Height          =   240
            Left            =   10260
            TabIndex        =   55
            Top             =   270
            Width           =   750
         End
         Begin VB.Label Label1 
            Caption         =   "��������"
            Height          =   240
            Left            =   6255
            TabIndex        =   52
            Top             =   270
            Width           =   750
         End
         Begin VB.Label lblPages 
            AutoSize        =   -1  'True
            Caption         =   "סԺ����"
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
            Caption         =   "�ɲ鿴�ѳ�Ժ       ��ļ����¼"
            Height          =   225
            Left            =   8490
            TabIndex        =   49
            Top             =   660
            Width           =   2925
         End
         Begin VB.Label lblDor 
            AutoSize        =   -1  'True
            Caption         =   "����ҽ��"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   645
            Width           =   720
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   4125
            TabIndex        =   47
            Top             =   270
            Width           =   360
         End
         Begin VB.Label lblNo 
            AutoSize        =   -1  'True
            Caption         =   "סԺ�š�"
            Height          =   180
            Left            =   3840
            TabIndex        =   46
            Top             =   660
            Width           =   720
         End
         Begin VB.Label lblDept 
            AutoSize        =   -1  'True
            Caption         =   "������ҡ�"
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
         Caption         =   "�ٴ�����:"
         Height          =   180
         Left            =   30
         TabIndex        =   88
         Top             =   3060
         Width           =   810
      End
      Begin VB.Label lblDiagnose 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���:"
         Height          =   180
         Left            =   30
         TabIndex        =   24
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע:"
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
            TabIndex        =   20
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
            TabIndex        =   13
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
            Caption         =   "ͼ������"
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
            TabIndex        =   77
            Top             =   2160
            Width           =   5250
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
               TabIndex        =   80
               Text            =   "��΢�����"
               Top             =   270
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
               TabIndex        =   79
               Top             =   1350
               Width           =   3915
            End
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
               TabIndex        =   78
               Top             =   690
               Width           =   3915
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
               TabIndex        =   83
               Top             =   300
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
               TabIndex        =   82
               Top             =   1290
               Width           =   960
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
               TabIndex        =   73
               Top             =   210
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
               TabIndex        =   72
               Top             =   975
               Width           =   4065
            End
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
               TabIndex        =   71
               Top             =   1800
               Width           =   4065
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
               TabIndex        =   76
               Top             =   210
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
               TabIndex        =   75
               Top             =   930
               Width           =   960
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
               TabIndex        =   74
               Top             =   1800
               Width           =   960
            End
         End
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
            TabIndex        =   64
            Top             =   1080
            Width           =   5250
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
               TabIndex        =   69
               Top             =   600
               Value           =   -1  'True
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
               Index           =   0
               Left            =   1935
               TabIndex        =   68
               Top             =   600
               Width           =   885
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
               TabIndex        =   67
               Top             =   300
               Width           =   1815
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
               TabIndex        =   66
               Top             =   300
               Width           =   1815
            End
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
               Caption         =   "���䱨��"
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
               TabIndex        =   34
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
               TabIndex        =   35
               Top             =   2820
               Width           =   450
            End
            Begin VB.Label lblAntibiotic 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������:"
               Height          =   180
               Left            =   30
               TabIndex        =   33
               Top             =   1560
               Width           =   630
            End
            Begin VB.Label lblMicrobe 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ϸ��:"
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
               Caption         =   "��ʾ�����Ŀ"
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
                  Caption         =   "���˵��"
                  BeginProperty Font 
                     Name            =   "����"
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
                     Name            =   "����"
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
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":DBDE
            Key             =   "����"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":14440
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":1ACA2
            Key             =   "��ֹ��ӡ"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatientReprotBrowse.frx":21504
            Key             =   "�������"
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
Private mblnShow As Boolean                                         '�����Ƿ���ʾ
Dim mlngKey As Long                                                 '��ǰѡ��ı걾ID
Dim mlngPatientID As Long                                           '����ID
Dim mReportDate As Date                                             '����ʱ��
Dim mlngGetPatientID As Long                                        '�ϼ������Ĳ���ID
Dim mintPatientType As Integer                                      '������Դ
Dim mlngPatientPage As Long                                         '������ҳ
Dim mlngValueC As Long                                              '΢����������
Dim mintVer As Integer                                              '�汾��25-�°� 10-�ϰ�
Dim mrsPatientVal As ADODB.Recordset                                '������Ϣ

Private mrsAntibioticValType As Recordset                           'ҩ���������
Private mstrPrivs As String                                         '������ϼ���Ȩ��


Private mobjFSO As New Scripting.FileSystemObject    'FSO����
Private objImg As Object
Private mblnShowBorder As Boolean                    '�Ƿ���ʾ�����border

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
        mlngPatientPage = Val(Trim(Replace(Replace(.Text, "��", ""), "��", "")))
        If .Text = "����" Then
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
            If MsgBox("�Ƿ�ѡ���Ѵ�ӡ��?", vbYesNo + vbQuestion + vbDefaultButton2, "�������") = vbYes Then
                VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("ѡ��"), 1, True
            Else
                VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("ѡ��"), 1, False
            End If
        Case ConMenu_Browse_ClsAll
            VsfColAllSelAllcls vsfLeft, vsfLeft.ColIndex("ѡ��"), 2, True
        Case ConMenu_Browse_Refresh
            ReadPatientList
        Case ConMenu_Browse_Print                                       '��ӡ
            BatchPrint (2)
        Case ConMenu_Browse_PrintView                                   '��ӡԤ��
            BatchPrint (1)
        Case ConMenu_Browse_PrintSet                                    '��ӡ����
            If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�汾"))) = 25 Then
                PrintReport Me, mlngKey, 3
            Else
                PtintOldReport Me, mlngKey, , 3
            End If
        Case ConMenu_Browse_Exit
            Unload Me
        Case ConMenu_pop_In
            lblNo.Caption = "סԺ�š�"
        Case ConMenu_pop_bed
            lblNo.Caption = "���š�"
        Case ConMenu_pop_Dept
            lblDept.Caption = "������ҡ�"
            InitDepts 0
        Case ConMenu_pop_DeptDistrict
            lblDept.Caption = "���벡����"
            InitDepts 1
        Case ConMenu_Browse_Find                                        '���μ���
            If mlngGetPatientID = 0 Then
                MsgBox "��ѡ��һ�����ˣ�", vbInformation, "������Ϣ"
                cboPatients.SetFocus
            Else
                Call ReadPatientList(1)
            End If
        Case ConMenu_Appfor_ClincHelp       '���Ʋο�
            Call ShowClincHelp
        Case ConMenu_Browse_PrintViewAll        'Ԥ������סԺ����
            Call PrintAll(1)
        Case ConMenu_Browse_PrintAll            '��ӡ����סԺ����
            Call PrintAll(2)
        Case ConMenu_Browse_PrintSetAll         '��ӡ����
            Call PrintAll(3)
        Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99 '��ҹ���ִ��
            Call ExePlugIn(Control.Parameter, mlngKey)
    End Select
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-08-13
'��    ��:  Ԥ��/��ӡ����˵�б���
'��    ��:
'           intType 1=Ԥ����2=��ӡ��3=��ӡ����
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
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
3             Call frmShowPatientAllReport.ShowMe(Me, mlngGetPatientID, Val(Trim(Replace(Replace(cboPages.Text, "��", ""), "��", ""))))
4         ElseIf intType = 2 Then
              '�°汨��
5             strSQL = "Select f_List2str(Cast(Collect(to_char(ID)) As t_Strlist)) �걾ID" & vbCrLf & _
                     "   From ���鱨���¼" & vbCrLf & _
                     "   Where HIS����ID = [1] And ��ҳID = [2] and ����� is not null"
6             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�°汨��", mlngGetPatientID, Val(Trim(Replace(Replace(cboPages.Text, "��", ""), "��", ""))))
7             If Not rsTmp.EOF Then
8                 strNewSampleIDs = rsTmp("�걾ID") & ""
9             End If

              '�ϰ汨��
10            strSQL = "Select f_List2str(Cast(Collect(to_char(ID)) As t_Strlist)) �걾ID" & vbCrLf & _
                     "   From ����걾��¼" & vbCrLf & _
                     "   Where ����ID = [1] And ��ҳID = [2] and ����� is not null"
11            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�ϰ汨��", mlngGetPatientID, Val(Trim(Replace(Replace(cboPages.Text, "��", ""), "��", ""))))
12            If Not rsTmp.EOF Then
13                strOldSampleIDs = rsTmp("�걾ID") & ""
14            End If
              
              '��ӡ�°汨��
15            If strNewSampleIDs <> "" Then
16                FunReportOpen gcnLisOracle, 2500, "ZL25_INSIDE_2500_109", Me, "�걾ID=" & strNewSampleIDs, intType
17                strID = Split(strNewSampleIDs, ",")
18                For i = 0 To UBound(strID)
19                    strSQL = "Zl_���鱨���ӡ_Edit(1," & Val(strID(i)) & ",1)"
20                    Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
21                Next
22            End If
              '��ӡ�ϰ汨��
23            If strOldSampleIDs <> "" Then
24                FunReportOpen gcnHisOracle, 100, "ZL1_INSIDE_1208_9", Me, "�걾ID=" & strOldSampleIDs, intType
25            End If
              
              
26        Else
27            FunReportPrintSet gcnLisOracle, 2500, "ZL25_INSIDE_2500_109", Me
28            FunReportPrintSet gcnHisOracle, 100, "ZL1_INSIDE_1208_9", Me
29        End If


30        Exit Sub
PrintAll_Error:
31        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(PrintAll)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
32        Err.Clear
End Sub

Private Function VsfColAllSelAllcls(objVSF As VSFlexGrid, intCol As Integer, Optional intSel As Integer, Optional blnSelect As Boolean, Optional strErr As String) As Boolean
          '����               ȫѡ��ȫ��ѡ���
          '����               intSel 0=����һ�н����ж� 1=ȫ��ѡ�� 2=ȫ����ѡ��

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
13                 If .Cell(flexcpFontBold, intRow, .ColIndex("ѡ��")) = False Then
14                    If blnSelect = True Then


15                       .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = intSel
16                    Else
17                       If Val(.TextMatrix(intRow, .ColIndex("��ӡ"))) = 0 Then
18                           .Cell(flexcpChecked, intRow, intCol, intRow, intCol) = intSel
19                       End If
20                    End If
21                End If
22            Next
23        End With
24        VsfColAllSelAllcls = True


25        Exit Function
VsfColAllSelAllcls_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(VsfColAllSelAllcls)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear

End Function

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Exit        '�˳�
            Control.Visible = mblnShowBorder
        Case ConMenu_Appfor_ClincHelp       '���Ʋο�
            Control.Visible = VerCompare(gSysInfo.VersionHIS, "10.35.120") <> -1
        Case ConMenu_Browse_PrintViewAll    'Ԥ���������б���
            Control.Visible = lblDept.Caption = "���벡����" And VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1
        Case ConMenu_Browse_PrintSetAll     '��ӡ����
            Control.Visible = lblDept.Caption = "���벡����" And VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1
        Case ConMenu_Browse_PrintAll        '��ӡ�������б���
            Control.Visible = lblDept.Caption = "���벡����" And VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1
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
    Dim blnPrintReport As Boolean   '��ӡ��ť�Ƿ����
    Dim strTemp As String

    If InStr(";" & mstrPrivs & ";", ";��ӡ���鱨��;") > 0 Then
        blnPrintReport = True
    Else
        blnPrintReport = False
    End If

    '�������ƹ������������������ڵĿ���״̬
    strTemp = ComGetPara(Sel_Lis_DB, "������������", 2500, 2001, "1|0")
    If strTemp <> "" Then
        chkApplyDate.value = Val(Split(strTemp, "|")(0))
        chkVerifyDate.value = Val(Split(strTemp, "|")(1))
    End If

    '���ܴ���������
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
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Refresh, "ˢ��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "��ӡԤ��")
        cbrControl.BeginGroup = True
        cbrControl.Enabled = blnPrintReport
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "��ӡ����")
        cbrControl.Enabled = blnPrintReport
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Print, "��ӡ")
        cbrControl.Enabled = blnPrintReport
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintViewAll, "Ԥ��סԺ����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_PrintAll, "��ӡסԺ����")
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSetAll, "��ӡ����  ")
        End With

        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "���μ���")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ClincHelp, "���Ʋο�")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "�˳�")
        cbrControl.BeginGroup = True
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



    dtpS = Now - 7: dtpE = Now
    dtpVS = Now - 7: dtpVE = Now

    picLeft.Width = GetSetWith(1)

    Me.chkGroup.value = Val(ComGetPara(Sel_Lis_DB, "�Ƿ���ʾ�����Ŀ", gSysInfo.SysNo, gSysInfo.ModlNo, 1))

    ReadSampleBacteriology 0
    ReadSampleBacteriology 0
    ReadSampleVal 0

    Set mrsAntibioticValType = GetDictType("ҩ���������")
End Sub

Private Function GetSetWith(ByVal intType As Integer) As Long
    '��ȡ/���ô�����߲��ֵĿ��
    '1-��ȡ,2-����
    If intType = 1 Then
        GetSetWith = ComGetPara(Sel_Lis_DB, "�ٴ����鱨��", 2500, 2500, "5000")
    ElseIf intType = 2 Then
        Call ComSetPara(Sel_Lis_DB, "�ٴ����鱨��", picLeft.Width, 2500, 2500)
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
    Call ComSetPara(Sel_Lis_DB, "������������", chkApplyDate.value & "|" & chkVerifyDate.value, 2500, 2001)
    Call ComSetPara(Sel_Lis_DB, "�Ƿ���ʾ�����Ŀ", Me.chkGroup.value, gSysInfo.SysNo, gSysInfo.ModlNo)
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
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Dept, "�������")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_DeptDistrict, "���벡��")
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
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_In, "סԺ��")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_bed, "����")
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
      '����           ���������������б�
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
          Dim blnReadData As Boolean       '��ʾ��Χ��ѯ����֮���Ƿ������ѯ

1         On Error GoTo ReadPatientList_Error

2         If mblnShow = False Then Exit Sub

3         strTitle = ComGetPara(Sel_Lis_DB, "���鱨����ʾ��", 2500, 2001)

4         Call ReadSampleVal(0)
5         RefreshTab Me.TabPage.Selected.Index

6         If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
7             strSQL = "Select a.id,c.����,0 ѡ��,A.����, Decode(A.�Ա�, '1', '��', '2', 'Ů', '9', 'δ֪', '') �Ա�, A.����, C.���� ������Ŀ,b.�걾���� �걾����, " & _
                     " A.סԺ��,a.�����, A.����,B.����ʱ��,a.����ID,a.����ʱ��,a.���ʱ��,a.��ע,a.���,a.΢����,a.���Ա���," & _
                     " a.����,b.����Id,a.ҽ��վ��ӡ ��ӡ ,b.������,a.������,a.�����, 25 �汾, 0 �������, a.�Ƿ�Ⱦ��,a.���˵��,a.���������,nvl(d.���䱨��״̬,0) ���䱨��" & vbNewLine & _
                     " From ���鱨���¼ A, ����������� B, ���������Ŀ C,���鱨�油���¼ D" & vbNewLine & _
                     " Where A.Id = B.�걾id And B.���id = C.Id(+) and a.id=d.�걾ID(+) and (a.����� is not null or a.��������� is not null) And b.���id Is Not Null "
8         Else
9             strSQL = "Select a.id,c.����,0 ѡ��,A.����, Decode(A.�Ա�, '1', '��', '2', 'Ů', '9', 'δ֪', '') �Ա�, A.����, C.���� ������Ŀ,b.�걾���� �걾����, " & _
                     " A.סԺ��,a.�����, A.����,B.����ʱ��,a.����ID,a.����ʱ��,a.���ʱ��,a.��ע,a.���,a.΢����,a.���Ա���," & _
                     " a.����,b.����Id,a.ҽ��վ��ӡ ��ӡ ,b.������,a.������,a.�����, 25 �汾, 0 �������, a.�Ƿ�Ⱦ��,a.���˵��,a.���������" & vbNewLine & _
                     " From ���鱨���¼ A, ����������� B, ���������Ŀ C" & vbNewLine & _
                     " Where A.Id = B.�걾id And B.���id = C.Id(+) and (a.����� is not null or a.��������� is not null) And b.���id Is Not Null "
10        End If

          '���û������ֹ������Ȩ�ޣ�����������ֹ����뱨��
11        If InStr(";" & mstrPrivs & ";", ";����ֹ�����;") <= 0 Then
12            strSQL = strSQL & " And b.����id Is Not Null "
13        End If
14        If mlngGetPatientID > 0 Then
15            strSQL = strSQL & " and a.HIS����ID = [4] "
              '        If mlngPatientPage <> 0 Then
              '            strSQL = strSQL & " and a.��ҳid = [8] "
              '        End If
16        End If

17        If chkApplyDate.value = 1 Then
18            strSQL = strSQL & " and b.����ʱ�� between [1] and [2] "
19        End If

20        If chkVerifyDate.value = 1 Then
21            strSQL = strSQL & " and a.���ʱ�� between [10] and [11] "
22        End If

23        blnReadData = True
          '    If chkVerifyDate.value = 1 Or chkApplyDate.value = 1 Then
          '�߷�ʱ�����Ʋ�ѯ
24        If chkVerifyDate.value = 0 And chkApplyDate.value = 1 Then
25            If Not funCheckRushHours(2500, 2001, "���������", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59")) Then
26                blnReadData = False
27            Else
28                blnReadData = True
29            End If
30        ElseIf chkVerifyDate.value = 1 And chkApplyDate.value = 0 Then
31            If Not funCheckRushHours(2500, 2001, "���������", CDate(Format(dtpVS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpVE, "yyyy-MM-dd") & " 23:59:59")) Then
32                blnReadData = False
33            Else
34                blnReadData = True
35            End If
36        ElseIf chkVerifyDate.value = 1 And chkApplyDate.value = 1 Then
37            If Not funCheckRushHours(2500, 2001, "���������", CDate(Format(IIf(dtpVS.value < dtpS.value, dtpVS.value, dtpS.value), "yyyy-MM-dd") & " 00:00:00"), CDate(Format(IIf(dtpVE.value < dtpE.value, dtpVE.value, dtpE.value), "yyyy-MM-dd") & " 23:59:59")) Then
38                blnReadData = False
39            Else
40                blnReadData = True
41            End If
42        End If
          '    End If

43        If blnReadData Then

44            If mintPatientType = 2 Then
45                If cboDept <> "00-���п���" Then
46                    If mlngGetPatientID <= 0 Then
47                        If lblDept.Caption = "���벡����" Then
48                            intDeptType = 2
49                        Else
50                            intDeptType = 1
51                        End If
52                        strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
                          '��strPatients���ȴ���3500ʱ,��Ҫ�ֽ�
53                        If Len(strPatients) >= 3500 Then
54                            stridSQL = MidPatients(strPatients)
55                        Else
56                            stridSQL = "Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list('" & strPatients & "') As Zltools.T_Numlist)) b"
57                        End If

58                        strSQL = strSQL & " and a.his����id in (" & stridSQL & ") "
59                    End If
                      '            If lblDept.Caption = "���벡����" Then
                      '        '                strDepts = GetDepts(cboDept.ItemData(cboDept.ListIndex))
                      '        '                If strDepts = "" Then
                      '        '                    strDepts = "0"
                      '        '                End If
                      '                'strSQL = strSQL & " and (b.������ұ��� in  (Select * From Table(Cast(F_Num2list([5]) As Zltools.T_Numlist))) or b.�������� = [6] )"
                      '                strSQL = strSQL & " and a.�������� = [6] "
                      '                strDept = Mid(Me.cboDept.Text, 1, InStr(cboDept, "-") - 1)
                      '            Else
                      '                strSQL = strSQL & " and a.������ұ��� = [6] "
                      '                strDept = Mid(Me.cboDept.Text, 1, InStr(cboDept, "-") - 1)
                      '            End If
60                End If


61            End If
62            If Trim(txtPatiNo <> "") Then
63                If lblNo.Caption = "סԺ�š�" Then
64                    strSQL = strSQL & " and a.סԺ�� = [3] "
65                ElseIf lblNo.Caption = "���š�" Then
66                    strSQL = strSQL & " and a.���� = [3] "
67                Else
68                    strSQL = strSQL & " and a.����� = [3] "
69                End If
70            End If

71            If Trim(cboPages.Text) <> "����" And Trim(cboPages.Text) <> "" Then
72                strSQL = strSQL & " And (a.��ҳid = [8] or a.��ҳid is null)"
73            End If

74            If Trim(cbodor.Text) <> "����" And Trim(cbodor.Text) <> "" Then
75                strSQL = strSQL & " and b.������ = [9] "
76            End If

77            If chkVerifyDate.value = 1 Then
78                strSQL = strSQL & " Order By c.����, a.id, a.����, a.���ʱ�� Desc "
79            Else
80                strSQL = strSQL & " Order By c.����, a.id, a.����, b.����ʱ�� Desc "
81            End If

82            If intType = 1 Then
83                If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then
84                    strSQL = "Select a.id,c.����,0 ѡ��,A.����, Decode(A.�Ա�, 1, '��', 2, 'Ů', 9, 'δ֪', '') �Ա�, A.����, C.���� ������Ŀ,b.�걾���� �걾����, " & _
                             " A.סԺ��,a.�����, A.����,B.����ʱ��,a.����ID,a.����ʱ��,a.���ʱ��,a.��ע,a.���,a.΢����,a.���Ա���," & _
                             " a.����,b.����Id,a.ҽ��վ��ӡ ��ӡ ,b.������,a.������,a.�����, 25 �汾, 0 �������, a.�Ƿ�Ⱦ��,a.���˵��,a.���������,nvl(d.���䱨��״̬,0) ���䱨�� " & _
                             " From ���鱨���¼ A, ����������� B, ���������Ŀ C,���鱨�油���¼ D" & _
                             " Where A.Id = B.�걾id And B.���id = C.Id(+) and a.id=d.�걾ID(+) and (a.����� is not null or a.��������� is not null) and a.����id = [1] " & _
                             " Order By c.����, a.����, a.�Ա�, a.����, a.���ʱ��, c.����"
85                Else
86                    strSQL = "Select a.id,c.����,0 ѡ��,A.����, Decode(A.�Ա�, 1, '��', 2, 'Ů', 9, 'δ֪', '') �Ա�, A.����, C.���� ������Ŀ,b.�걾���� �걾����, " & _
                             " A.סԺ��,a.�����, A.����,B.����ʱ��,a.����ID,a.����ʱ��,a.���ʱ��,a.��ע,a.���,a.΢����,a.���Ա���," & _
                             " a.����,b.����Id,a.ҽ��վ��ӡ ��ӡ ,b.������,a.������,a.�����, 25 �汾, 0 �������, a.�Ƿ�Ⱦ��,a.���˵��,a.���������" & _
                             " From ���鱨���¼ A, ����������� B, ���������Ŀ C" & _
                             " Where A.Id = B.�걾id And B.���id = C.Id(+) and (a.����� is not null or a.��������� is not null) and a.����id = [1] " & _
                             " Order By c.����, a.����, a.�Ա�, a.����, a.���ʱ��, c.����"
87                End If
88                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���벡���б�", mlngGetPatientID)
89            Else
90                Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���벡���б�", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
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
103               .ColKey(1) = "ѡ��": .ColWidth(.ColIndex("ѡ��")) = 500: .ColAlignment(.ColIndex("ѡ��")) = flexAlignCenterCenter    ': .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
104               .ColKey(2) = "����": .ColWidth(.ColIndex("����")) = 250: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
105               .ColKey(3) = "��ӡ": .ColWidth(.ColIndex("��ӡ")) = 250: .ColAlignment(.ColIndex("��ӡ")) = flexAlignCenterCenter
106               .ColKey(4) = "����": .ColWidth(.ColIndex("����")) = 750: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
107               .ColKey(5) = "�Ա�": .ColWidth(.ColIndex("�Ա�")) = 500: .ColAlignment(.ColIndex("�Ա�")) = flexAlignCenterCenter
108               .ColKey(6) = "����": .ColWidth(.ColIndex("����")) = 500: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
109               .ColKey(7) = "������Ŀ": .ColWidth(.ColIndex("������Ŀ")) = 2200: .ColAlignment(.ColIndex("������Ŀ")) = flexAlignCenterCenter
110               .ColKey(8) = "�걾����": .ColWidth(.ColIndex("�걾����")) = 1200: .ColAlignment(.ColIndex("�걾����")) = flexAlignCenterCenter
111               .ColKey(9) = "���ʱ��": .ColWidth(.ColIndex("���ʱ��")) = 2000: .ColAlignment(.ColIndex("���ʱ��")) = flexAlignCenterCenter
112               .ColKey(10) = "סԺ��": .ColWidth(.ColIndex("סԺ��")) = 750: .ColAlignment(.ColIndex("סԺ��")) = flexAlignCenterCenter
113               .ColKey(11) = "����": .ColWidth(.ColIndex("����")) = 500: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter
114               .ColKey(12) = "����ʱ��": .ColWidth(.ColIndex("����ʱ��")) = 2000: .ColAlignment(.ColIndex("����ʱ��")) = flexAlignCenterCenter
115               .ColKey(13) = "����ID": .ColWidth(.ColIndex("����ID")) = 2000: .ColAlignment(.ColIndex("����ID")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ID")) = True
116               .ColKey(14) = "����ʱ��": .ColWidth(.ColIndex("����ʱ��")) = 2000: .ColAlignment(.ColIndex("����ʱ��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����ʱ��")) = True
117               .ColKey(15) = "��ע": .ColWidth(.ColIndex("��ע")) = 2000: .ColAlignment(.ColIndex("��ע")) = flexAlignCenterCenter: .ColHidden(.ColIndex("��ע")) = True
118               .ColKey(16) = "���": .ColWidth(.ColIndex("���")) = 2000: .ColAlignment(.ColIndex("���")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���")) = True
119               .ColKey(17) = "΢����": .ColWidth(.ColIndex("΢����")) = 2000: .ColAlignment(.ColIndex("΢����")) = flexAlignCenterCenter: .ColHidden(.ColIndex("΢����")) = True
120               .ColKey(18) = "���Ա���": .ColWidth(.ColIndex("���Ա���")) = 2000: .ColAlignment(.ColIndex("���Ա���")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���Ա���")) = True
121               .ColKey(19) = "����Id": .ColWidth(.ColIndex("����Id")) = 2000: .ColAlignment(.ColIndex("����Id")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����Id")) = True
122               .ColKey(20) = "������": .ColWidth(.ColIndex("������")) = 750: .ColAlignment(.ColIndex("������")) = flexAlignCenterCenter
123               .ColKey(21) = "������": .ColWidth(.ColIndex("������")) = 750: .ColAlignment(.ColIndex("������")) = flexAlignCenterCenter
124               .ColKey(22) = "�����": .ColWidth(.ColIndex("�����")) = 750: .ColAlignment(.ColIndex("�����")) = flexAlignCenterCenter
125               .ColKey(23) = "�汾": .ColWidth(.ColIndex("�汾")) = 750: .ColAlignment(.ColIndex("�汾")) = flexAlignCenterCenter: .ColHidden(.ColIndex("�汾")) = True
126               .ColKey(24) = "�������": .ColWidth(.ColIndex("�������")) = 750: .ColAlignment(.ColIndex("�������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("�������")) = True
127               .ColKey(25) = "���˵��": .ColWidth(.ColIndex("���˵��")) = 750: .ColAlignment(.ColIndex("���˵��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���˵��")) = True
128               .ColKey(26) = "���������": .ColWidth(.ColIndex("���������")) = 0: .ColAlignment(.ColIndex("���������")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���������")) = True
129               .ColKey(27) = "���䱨��": .ColWidth(.ColIndex("���䱨��")) = 0: .ColAlignment(.ColIndex("���䱨��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���䱨��")) = True
130           Else
131               var_tmp = Split(strTitle, ";")
132               .Rows = 1
133               .Cols = UBound(var_tmp) + 1
134               For lngLoop = LBound(var_tmp) To UBound(var_tmp)
135                   var_SubTmp = Split(var_tmp(lngLoop), ",")
136                   .ColKey(lngLoop) = var_SubTmp(0): .ColWidth(.ColIndex(var_SubTmp(0))) = var_SubTmp(1): .ColAlignment(.ColIndex(var_SubTmp(0))) = flexAlignCenterCenter: .ColHidden(.ColIndex(var_SubTmp(0))) = Not (Val(var_SubTmp(2)) = 1)
                      '                If var_SubTmp(0) = "������" Or var_SubTmp(0) = "����ʱ��" Then
                      '                    .ColHidden(.ColIndex(var_SubTmp(0))) = Not InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
                      '                End If
137               Next
138               If .ColIndex("���䱨��") < 0 Then
139                   .Cols = .Cols + 1
140                   .ColKey(.Cols - 1) = "���䱨��": .ColWidth(.ColIndex("���䱨��")) = 0: .ColAlignment(.ColIndex("���䱨��")) = flexAlignCenterCenter: .ColHidden(.ColIndex("���䱨��")) = True
141               End If
142           End If
143           .TextMatrix(0, .ColIndex("ѡ��")) = ""
144           .TextMatrix(0, .ColIndex("����")) = ""
145           .TextMatrix(0, .ColIndex("��ӡ")) = ""
146           .TextMatrix(0, .ColIndex("����")) = "����"
147           .TextMatrix(0, .ColIndex("�Ա�")) = "�Ա�"
148           .TextMatrix(0, .ColIndex("����")) = "����"
149           .TextMatrix(0, .ColIndex("������Ŀ")) = "������Ŀ"
150           .TextMatrix(0, .ColIndex("�걾����")) = "�걾����"
151           .TextMatrix(0, .ColIndex("���ʱ��")) = "���ʱ��"
152           .TextMatrix(0, .ColIndex("סԺ��")) = "סԺ��"
153           .TextMatrix(0, .ColIndex("����")) = "����"
154           .TextMatrix(0, .ColIndex("����ʱ��")) = "����ʱ��"
155           .TextMatrix(0, .ColIndex("������")) = "������"
156           .TextMatrix(0, .ColIndex("������")) = "������"
157           .TextMatrix(0, .ColIndex("�����")) = "�����"
158           .TextMatrix(0, .ColIndex("�汾")) = "�汾"
159           .TextMatrix(0, .ColIndex("�������")) = "�������"
160           .Row = 0: .Col = .ColIndex("ѡ��"): .CellPicture = imgVsf.ListImages("ѡ��").ExtractIcon
161           .Row = 0: .Col = .ColIndex("����"): .CellPicture = imgVsf.ListImages("����").ExtractIcon
162           .Row = 0: .Col = .ColIndex("��ӡ"): .CellPicture = imgVsf.ListImages("��ӡ").ExtractIcon

163           If blnReadData Then

164               Do Until rsTmp.EOF
165                   If rsTmp("������") & "" <> gUserInfo.Name And rsTmp("������") & "" <> "" And InStr(";" & mstrPrivs & ";", ";�鿴��Ⱦ������;") <= 0 And Val(rsTmp("�Ƿ�Ⱦ��") & "") = 1 Then
                          '���в鿴��Ⱦ������Ȩ��ʱ���ʲ鿴��Ⱦ������,����������Ⱦ���������������һ������
166                       rsTmp.MoveNext
167                   Else
168                       If InStr(strFenLei, rsTmp("����") & ",") > 0 Then
169                           If lngKey <> Val(rsTmp("id") & "") Then
170                               .Rows = .Rows + 1
171                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""

172                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 2
173                               .TextMatrix(.Rows - 1, .ColIndex("��ӡ")) = rsTmp("��ӡ") & ""

174                               If rsTmp("����") & "" = 1 Then
175                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("����")) = imgVsf.ListImages("����").ExtractIcon
176                               End If

177                               If Val(rsTmp("��ӡ") & "") > 0 Then
178                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("��ӡ")) = imgVsf.ListImages("��ӡ").ExtractIcon
179                               End If

180                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
181                               .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsTmp("�Ա�") & ""
182                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
183                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsTmp("������Ŀ") & ""
184                               .TextMatrix(.Rows - 1, .ColIndex("�걾����")) = rsTmp("�걾����") & ""
185                               .TextMatrix(.Rows - 1, .ColIndex("���ʱ��")) = Format(rsTmp("���ʱ��") & "", "yyyy-mm-dd HH:mm:ss")
186                               .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = rsTmp("סԺ��") & ""
187                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
188                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsTmp("����ʱ��") & ""
189                               .TextMatrix(.Rows - 1, .ColIndex("����ID")) = rsTmp("����ID") & ""
190                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsTmp("����ʱ��") & ""
191                               .TextMatrix(.Rows - 1, .ColIndex("��ע")) = rsTmp("��ע") & ""
192                               .TextMatrix(.Rows - 1, .ColIndex("���")) = rsTmp("���") & ""
193                               .TextMatrix(.Rows - 1, .ColIndex("΢����")) = rsTmp("΢����") & ""
194                               .TextMatrix(.Rows - 1, .ColIndex("���Ա���")) = rsTmp("���Ա���") & ""
195                               .TextMatrix(.Rows - 1, .ColIndex("����Id")) = rsTmp("����Id") & ""
196                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
197                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
198                               .TextMatrix(.Rows - 1, .ColIndex("�����")) = rsTmp("�����") & ""
199                               .TextMatrix(.Rows - 1, .ColIndex("�汾")) = rsTmp("�汾") & ""
200                               .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsTmp("�������") & ""
201                               .TextMatrix(.Rows - 1, .ColIndex("���˵��")) = rsTmp("���˵��") & ""
202                               .TextMatrix(.Rows - 1, .ColIndex("���������")) = rsTmp("���������") & ""
203                               If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then .TextMatrix(.Rows - 1, .ColIndex("���䱨��")) = rsTmp("���䱨��") & ""
204                           Else
205                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) & "," & rsTmp("������Ŀ") & ""
206                           End If
207                           lngKey = Val(rsTmp("id") & "")
208                       Else
209                           If lngKey <> Val(rsTmp("id") & "") Then
210                               .Rows = .Rows + 2

211                               For i = 1 To .Cols - 1
212                                   .TextMatrix(.Rows - 2, i) = CStr(rsTmp("����") & "") & "(�°�)"
213                               Next

                                  '�ϲ�
214                               .MergeRow(.Rows - 2) = True
215                               .MergeCellsFixed = flexMergeRestrictRows

                                  '����
216                               .IsSubtotal(.Rows - 2) = True
217                               .RowOutlineLevel(.Rows - 2) = 1

                                  '�Ӵ�
218                               .Cell(flexcpFontBold, .Rows - 2, 1) = True

219                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""

220                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 2
221                               .TextMatrix(.Rows - 1, .ColIndex("��ӡ")) = rsTmp("��ӡ") & ""

222                               If rsTmp("����") & "" = 1 Then
223                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("����")) = imgVsf.ListImages("����").ExtractIcon
224                               End If

225                               If Val(rsTmp("��ӡ") & "") > 0 Then
226                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("��ӡ")) = imgVsf.ListImages("��ӡ").ExtractIcon
227                               End If

228                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
229                               .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsTmp("�Ա�") & ""
230                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
231                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsTmp("������Ŀ") & ""
232                               .TextMatrix(.Rows - 1, .ColIndex("�걾����")) = rsTmp("�걾����") & ""
233                               .TextMatrix(.Rows - 1, .ColIndex("���ʱ��")) = Format(rsTmp("���ʱ��") & "", "yyyy-mm-dd HH:mm:ss")
234                               .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = rsTmp("סԺ��") & ""
235                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
236                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsTmp("����ʱ��") & ""
237                               .TextMatrix(.Rows - 1, .ColIndex("����ID")) = rsTmp("����ID") & ""
238                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsTmp("����ʱ��") & ""
239                               .TextMatrix(.Rows - 1, .ColIndex("��ע")) = rsTmp("��ע") & ""
240                               .TextMatrix(.Rows - 1, .ColIndex("���")) = rsTmp("���") & ""
241                               .TextMatrix(.Rows - 1, .ColIndex("΢����")) = rsTmp("΢����") & ""
242                               .TextMatrix(.Rows - 1, .ColIndex("���Ա���")) = rsTmp("���Ա���") & ""
243                               .TextMatrix(.Rows - 1, .ColIndex("����Id")) = rsTmp("����Id") & ""
244                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
245                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsTmp("������") & ""
246                               .TextMatrix(.Rows - 1, .ColIndex("�����")) = rsTmp("�����") & ""
247                               .TextMatrix(.Rows - 1, .ColIndex("�汾")) = rsTmp("�汾") & ""
248                               .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsTmp("�������") & ""
249                               .TextMatrix(.Rows - 1, .ColIndex("���˵��")) = rsTmp("���˵��") & ""
250                               .TextMatrix(.Rows - 1, .ColIndex("���������")) = rsTmp("���������") & ""
251                               If VerCompare(gSysInfo.VersionLIS, "10.35.140") <> -1 Then .TextMatrix(.Rows - 1, .ColIndex("���䱨��")) = rsTmp("���䱨��") & ""
252                           Else
253                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) & "," & rsTmp("������Ŀ") & ""
254                           End If
255                           lngKey = Val(rsTmp("id") & "")
256                           strFenLei = strFenLei & rsTmp("����") & ","
257                       End If
258                       rsTmp.MoveNext
259                   End If
260               Loop

261               strFenLei = ""
262               Set rsOldLisData = GetOldLisData(intType)
263               If rsOldLisData.RecordCount > 0 Then
264                   Do Until rsOldLisData.EOF
265                       If InStr(strFenLei, rsOldLisData("����") & ",") > 0 Then
266                           If lngKey <> Val(rsOldLisData("id") & "") Then
267                               .Rows = .Rows + 1
268                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsOldLisData("id") & ""
269                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 2
270                               .TextMatrix(.Rows - 1, .ColIndex("��ӡ")) = rsOldLisData("��ӡ") & ""

271                               If rsOldLisData("����") & "" = 1 Then
272                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("����")) = imgVsf.ListImages("����").ExtractIcon
273                               End If

274                               If Val(rsOldLisData("��ӡ") & "") > 0 Then
275                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("��ӡ")) = imgVsf.ListImages("��ӡ").ExtractIcon
276                               End If

277                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
278                               .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsOldLisData("�Ա�") & ""
279                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
280                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsOldLisData("������Ŀ") & ""
281                               .TextMatrix(.Rows - 1, .ColIndex("�걾����")) = rsOldLisData("�걾����") & ""
282                               .TextMatrix(.Rows - 1, .ColIndex("���ʱ��")) = Format(rsOldLisData("���ʱ��") & "", "yyyy-mm-dd HH:mm:ss")
283                               .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = rsOldLisData("סԺ��") & ""
284                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
285                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsOldLisData("����ʱ��") & ""
286                               .TextMatrix(.Rows - 1, .ColIndex("����ID")) = rsOldLisData("����ID") & ""
287                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsOldLisData("����ʱ��") & ""
288                               .TextMatrix(.Rows - 1, .ColIndex("��ע")) = rsOldLisData("��ע") & ""
289                               .TextMatrix(.Rows - 1, .ColIndex("���")) = rsOldLisData("���") & ""
290                               .TextMatrix(.Rows - 1, .ColIndex("΢����")) = rsOldLisData("΢����") & ""
291                               .TextMatrix(.Rows - 1, .ColIndex("���Ա���")) = rsOldLisData("���Ա���") & ""
292                               .TextMatrix(.Rows - 1, .ColIndex("����Id")) = rsOldLisData("����Id") & ""
293                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""
294                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""
295                               .TextMatrix(.Rows - 1, .ColIndex("�����")) = rsOldLisData("�����") & ""
296                               .TextMatrix(.Rows - 1, .ColIndex("�汾")) = rsOldLisData("�汾") & ""
297                               .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsOldLisData("�������") & ""
298                               .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsOldLisData("�������") & ""
299                           Else
300                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) & "," & rsOldLisData("������Ŀ") & ""
301                           End If
302                           lngKey = Val(rsOldLisData("id") & "")
303                       Else
304                           If lngKey <> Val(rsOldLisData("id") & "") Then
305                               .Rows = .Rows + 2
306                               For i = 1 To .Cols - 1
307                                   .TextMatrix(.Rows - 2, i) = CStr(rsOldLisData("����") & "") & "(�ϰ�)"
308                               Next

                                  '�ϲ�
309                               .MergeRow(.Rows - 2) = True
310                               .MergeCellsFixed = flexMergeRestrictRows

                                  '����
311                               .IsSubtotal(.Rows - 2) = True
312                               .RowOutlineLevel(.Rows - 2) = 1

                                  '�Ӵ�
313                               .Cell(flexcpFontBold, .Rows - 2, 1) = True

314                               .TextMatrix(.Rows - 1, .ColIndex("id")) = rsOldLisData("id") & ""
315                               .Cell(flexcpChecked, .Rows - 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 2
316                               .TextMatrix(.Rows - 1, .ColIndex("��ӡ")) = rsOldLisData("��ӡ") & ""

317                               If rsOldLisData("����") & "" = 1 Then
318                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("����")) = imgVsf.ListImages("����").ExtractIcon
319                               End If

320                               If Val(rsOldLisData("��ӡ") & "") > 0 Then
321                                   .Cell(flexcpPicture, .Rows - 1, .ColIndex("��ӡ")) = imgVsf.ListImages("��ӡ").ExtractIcon
322                               End If

323                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
324                               .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = rsOldLisData("�Ա�") & ""
325                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
326                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = rsOldLisData("������Ŀ") & ""
327                               .TextMatrix(.Rows - 1, .ColIndex("�걾����")) = rsOldLisData("�걾����") & ""
328                               .TextMatrix(.Rows - 1, .ColIndex("���ʱ��")) = Format(rsOldLisData("���ʱ��") & "", "yyyy-mm-dd HH:mm:ss")
329                               .TextMatrix(.Rows - 1, .ColIndex("סԺ��")) = rsOldLisData("סԺ��") & ""
330                               .TextMatrix(.Rows - 1, .ColIndex("����")) = rsOldLisData("����") & ""
331                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsOldLisData("����ʱ��") & ""
332                               .TextMatrix(.Rows - 1, .ColIndex("����ID")) = rsOldLisData("����ID") & ""
333                               .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = rsOldLisData("����ʱ��") & ""
334                               .TextMatrix(.Rows - 1, .ColIndex("��ע")) = rsOldLisData("��ע") & ""
335                               .TextMatrix(.Rows - 1, .ColIndex("���")) = rsOldLisData("���") & ""
336                               .TextMatrix(.Rows - 1, .ColIndex("΢����")) = rsOldLisData("΢����") & ""
337                               .TextMatrix(.Rows - 1, .ColIndex("���Ա���")) = rsOldLisData("���Ա���") & ""
338                               .TextMatrix(.Rows - 1, .ColIndex("����Id")) = rsOldLisData("����Id") & ""
339                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""
340                               .TextMatrix(.Rows - 1, .ColIndex("������")) = rsOldLisData("������") & ""
341                               .TextMatrix(.Rows - 1, .ColIndex("�����")) = rsOldLisData("�����") & ""
342                               .TextMatrix(.Rows - 1, .ColIndex("�汾")) = rsOldLisData("�汾") & ""
343                               .TextMatrix(.Rows - 1, .ColIndex("�������")) = rsOldLisData("�������") & ""
344                           Else
345                               .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) = .TextMatrix(.Rows - 1, .ColIndex("������Ŀ")) & "," & rsOldLisData("������Ŀ") & ""
346                           End If
347                           lngKey = Val(rsOldLisData("id") & "")
348                           strFenLei = strFenLei & rsOldLisData("����") & ","
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
364                       .Cell(flexcpSort, .FixedRows, .ColIndex("���ʱ��"), .Rows - 1, .ColIndex("���ʱ��")) = 2
365                       .Cell(flexcpSort, .FixedRows, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = 1
366                   Else
367                       .Cell(flexcpSort, .FixedRows, .ColIndex("����ʱ��"), .Rows - 1, .ColIndex("����ʱ��")) = 2
368                       .Cell(flexcpSort, .FixedRows, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = 1
369                   End If
370               End If
371           End If
372       End With


373       Exit Sub
ReadPatientList_Error:
374       Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(ReadPatientList)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
375       Err.Clear

End Sub

Private Function MidPatients(ByVal strPatients As String)
'��strPatients���ȴ���3500ʱ,��Ҫ�ֽ�
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

2         strSQL = "Select Distinct a.Id, a.�������� ����, 0 ѡ��, a.����, a.�Ա�, a.����, a.������Ŀ ������Ŀ, a.�걾����, a.סԺ��, a.�����, a.����, a.����ʱ��, a.����id, a.����ʱ��," & vbNewLine & _
                  "                a.���ʱ��, a.���鱸ע ��ע, d.���, a.΢����걾 ΢����, 1 ���Ա���, 0 ����, a.ҽ��id ����id, a.��ӡ���� ��ӡ, a.������, a.������, a.�����, 10 �汾," & vbNewLine & _
                  "                a.������ �������" & vbNewLine & _
                  "From ����걾��¼ A, ����ҽ����¼ C," & vbNewLine & _
                  "     (Select b.ҽ��id ҽ��id, f_List2str(Cast(Collect(b.��Ŀ || ':' || b.����) As t_Strlist)) ���" & vbNewLine & _
                  "       From ����걾��¼ A, ����ҽ������ B" & vbNewLine & _
                  "       Where a.ҽ��id = b.ҽ��id  [����]" & vbNewLine & _
                  "       Group By b.ҽ��id) D" & vbNewLine & _
                  "Where a.ҽ��id = c.Id(+) And a.ҽ��id = d.ҽ��id(+) And a.����� Is Not Null"

3         If mlngGetPatientID > 0 Then
4             strWhere = strWhere & " and a.����ID = [4] "
5             strSQL = strSQL & " and a.����ID = [4] "
6         End If

7         If chkApplyDate.value = 1 Then
8             strWhere = strWhere & " and a.����ʱ�� between [1] and [2] "
9             strSQL = strSQL & " and a.����ʱ�� between [1] and [2] "
10        End If

11        If chkVerifyDate.value = 1 Then
12            strWhere = strWhere & " and a.���ʱ�� between [10] and [11] "
13            strSQL = strSQL & " and a.���ʱ�� between [10] and [11] "
14        End If

15        If mintPatientType = 2 Then
16            If cboDept <> "00-���п���" Then
17                If mlngGetPatientID <= 0 Then
18                    If lblDept.Caption = "���벡����" Then
19                        intDeptType = 2
20                    Else
21                        intDeptType = 1
22                    End If
23                    strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
                      '��strPatients���ȴ���3500ʱ,��Ҫ�ֽ�
24                    If Len(strPatients) >= 3500 Then
25                        stridSQL = MidPatients(strPatients)
26                    Else
27                        stridSQL = "Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list('" & strPatients & "') As Zltools.T_Numlist)) b"
28                    End If

29                    strWhere = strWhere & " and a.����id in (" & stridSQL & ") "
30                    strSQL = strSQL & " and a.����id in (" & stridSQL & ") "

31                End If

32            End If


33        End If
34        If Trim(txtPatiNo <> "") Then
35            If lblNo.Caption = "סԺ�š�" Then
36                strWhere = strWhere & " and a.סԺ�� = [3] "
37                strSQL = strSQL & " and a.סԺ�� = [3] "

38            ElseIf lblNo.Caption = "���š�" Then
39                strWhere = strWhere & " and a.���� = [3] "
40                strSQL = strSQL & " and a.���� = [3] "

41            Else
42                strWhere = strWhere & " and a.����� = [3] "
43                strSQL = strSQL & " and a.����� = [3] "

44            End If
45        End If

46        If Trim(cbodor.Text) <> "����" And Trim(cbodor.Text) <> "" Then
47            strWhere = strWhere & " and a.������ = [9] "
48            strSQL = strSQL & " and a.������ = [9] "
49        End If

50        If chkVerifyDate.value = 1 Then
51            strSQL = strSQL & " Order By a.��������, a.id, a.����, a.���ʱ�� Desc "
52        Else
53            strSQL = strSQL & " Order By a.��������, a.id, a.����, a.����ʱ�� Desc "
54        End If

55        strSQL = Replace(strSQL, "[����]", strWhere)

56        If intType = 1 Then
57            strSQL = "Select Distinct a.Id, a.�������� ����, 0 ѡ��, a.����, a.�Ա�, a.����, a.������Ŀ ������Ŀ, a.�걾����, a.סԺ��, a.�����, a.����, a.����ʱ��, a.����id, a.����ʱ��," & vbNewLine & _
                      "                a.���ʱ��, a.���鱸ע ��ע, d.���, a.΢����걾 ΢����, 1 ���Ա���, 0 ����, a.ҽ��id ����id, a.��ӡ���� ��ӡ, a.������, a.������, a.�����, 10 �汾," & vbNewLine & _
                      "                a.������ �������" & vbNewLine & _
                      "From ����걾��¼ A, ����ҽ����¼ C," & vbNewLine & _
                      "     (Select b.ҽ��id ҽ��id, f_List2str(Cast(Collect(b.��Ŀ || ':' || b.����) As t_Strlist)) ���" & vbNewLine & _
                      "       From ����걾��¼ A, ����ҽ������ B" & vbNewLine & _
                      "       Where a.ҽ��id = b.ҽ��id and a.����id = [1]" & vbNewLine & _
                      "       Group By b.ҽ��id) D" & vbNewLine & _
                      "Where a.ҽ��id = c.Id(+) And a.ҽ��id = d.ҽ��id(+) And a.����� Is Not Null and a.����id = [1]" & vbNewLine & _
                      " Order By a.��������, a.����, a.�Ա�, a.����, a.���ʱ��, a.������Ŀ"
58            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���벡���б�", mlngGetPatientID)
59        Else
60            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���벡���б�", CDate(Format(dtpS, "yyyy-MM-dd") & " 00:00:00"), _
                                  CDate(Format(dtpE, "yyyy-MM-dd") & " 23:59:59"), txtPatiNo, mlngGetPatientID, strDepts, strDept, _
                                  strPatients, mlngPatientPage, cbodor.Text, CDate(Format(dtpVS, "yyyy-MM-dd") & " 00:00:00"), CDate(Format(dtpVE, "yyyy-MM-dd") & " 23:59:59"))
61        End If

62        Set GetOldLisData = rsTmp


63        Exit Function
GetOldLisData_Error:
64        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetOldLisData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
65        Err.Clear

End Function

Private Sub ReadPatientVal(lngSampleID As Long)
          Dim strSQL As String
1         On Error GoTo ReadPatientVal_Error

2         If mintPatientType = 2 Then
3             strSQL = "Select Distinct a.����id,b.�����,a.סԺ��,a.��Ժ����,a.����, a.��ҳid," & _
                       " a.��Ժ����, a.��Ժ����  From ������ҳ a,������Ϣ B where a.����id=b.����id and a.����id =[1]  order by ��ҳID"

4             Set mrsPatientVal = ComOpenSQL(Sel_His_DB, strSQL, "���鼼ʦվ", lngSampleID)

5             With Me.cboPages
6                 .Clear
7                 .AddItem "����"
8                 .ItemData(.NewIndex) = 0
9                 Do Until mrsPatientVal.EOF
10                    .AddItem "�� " & mrsPatientVal("��ҳID") & " ��"
11                    .ItemData(.NewIndex) = mrsPatientVal("����id")
12                    mrsPatientVal.MoveNext
13                Loop
14                If mrsPatientVal.RecordCount > 0 Then
15                    mrsPatientVal.MoveLast
16                    mlngPatientPage = mrsPatientVal("��ҳID")
17                    .Text = "�� " & mlngPatientPage & " ��"
18                End If
19                Call readDate(True)
20            End With
21        End If


22        Exit Sub
ReadPatientVal_Error:
23        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(ReadPatientVal)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
24        Err.Clear

End Sub

Private Sub readDate(ByVal blnFilter As Boolean)
1         On Error GoTo readDate_Error

2         If mrsPatientVal.RecordCount > 0 Then
3             mrsPatientVal.MoveFirst
4             If blnFilter = True Then
5                 mrsPatientVal.Filter = "��ҳID=" & mlngPatientPage & ""
6                 If mrsPatientVal.RecordCount <> 0 Then
7                     Me.dtpS.value = IIf(IsNull(mrsPatientVal("��Ժ����")), Currentdate, mrsPatientVal("��Ժ����"))
8                     Me.dtpE.value = IIf(IsNull(mrsPatientVal("��Ժ����")), Currentdate, mrsPatientVal("��Ժ����"))
9                     If lblNo.Caption = "סԺ�š�" Then
10                        txtPatiNo.Text = mrsPatientVal("סԺ��") & ""
11                    ElseIf lblNo.Caption = "�����" Then
12                        txtPatiNo.Text = mrsPatientVal("�����") & ""
13                    Else
14                        txtPatiNo.Text = mrsPatientVal("��Ժ����") & ""
15                    End If
'16                    cboDept.ListIndex = 0
17                End If
18            Else
19                Me.dtpS.value = IIf(IsNull(mrsPatientVal("��Ժ����")), Currentdate, mrsPatientVal("��Ժ����"))
20                mrsPatientVal.MoveLast
21                Me.dtpE.value = IIf(IsNull(mrsPatientVal("��Ժ����")), Currentdate, mrsPatientVal("��Ժ����"))
22                txtPatiNo.Text = ""
'23                cboDept.ListIndex = 0
24                txtPatiNo = ""
25            End If
26            mrsPatientVal.Filter = ""
27        End If


28        Exit Sub
readDate_Error:
29        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(readDate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
30        Err.Clear
End Sub


Private Sub dtpE_Change()
    Me.cboPages.Text = "����"
End Sub

Private Sub dtpS_Change()
    Me.cboPages.Text = "����"
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
3             If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
4                 strSQL = "select * from ( " & vbNewLine & _
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')' || decode(g.����ʱ��,null,'', '(' || g.����ʱ�� || ')')  ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,b.�Ƿ����," & vbNewLine & _
                         "       c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ��,B.ID ���� " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e,��������걾 F,��������ʱ�䷽�� G" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and d.id =e.���id and  b.ID=F.������ϸid(+) and F.���ܷ���id=G.id(+) AND b.���id is not null and e.���id is not null and b.������ is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')' || decode(g.����ʱ��,null,'', '(' || g.����ʱ�� || ')') ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,b.�Ƿ����," & vbNewLine & _
                         "       c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ��,B.ID ���� " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e,��������걾 F,��������ʱ�䷽�� G" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and b.ID=F.������ϸid(+) and F.���ܷ���id=G.id(+) AND e.���id is null and b.���id is null and b.������ is not null and a.id = [1] ) order by ���� desc" & vbNewLine
5             Else
6                 strSQL = "select * from ( " & vbNewLine & _
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')'  ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,b.�Ƿ����," & vbNewLine & _
                         "       c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ��,B.ID ���� " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and d.id =e.���id and b.���id is not null and e.���id is not null and b.������ is not null and a.id = [1] " & vbNewLine & _
                         " union all " & vbNewLine & _
                           "select '' ���,c.id,c.������ || '(' || c.Ӣ���� || ')' ������Ŀ,b.������ ���,b.�ϴν�� �ϴ�," & vbNewLine & _
                         "       c.��λ,b.����ο� �ο�,a.������Դ ��������,e.ҽ��id,e.���id,d.���� �������," & vbNewLine & _
                         "       e.�շ�״̬,e.Ӧ�ս��,e.ʵ�ս��,b.�ο���ֵ,b.�ο���ֵ,c.�������,b.������ ��־���, " & vbNewLine & _
                         "       e.id �������ID,b.�����־, b.OD, b.CUTOFF, b.SCO,c.�������,c.���㹫ʽ,b.�Ƿ����," & vbNewLine & _
                         "       c.ָ�����,c.�ٴ�����,c.��Ŀ���,nvl(c.С��λ��,2) С��λ��,b.�ϴα�־,d.���� ��ϱ���,a.����ID,a.����ʱ��,B.ID ���� " & vbNewLine & _
                           "from ���鱨���¼ a, ���鱨����ϸ b,����ָ�� c,���������Ŀ d,����������� e" & vbNewLine & _
                           "where a.id = b.�걾id and  b.��Ŀid = c.id and  b.���id = d.id(+) and" & vbNewLine & _
                         "      b.�걾id = e.�걾id and e.���id is null and b.���id is null and b.������ is not null and a.id = [1] ) order by ���� desc" & vbNewLine
7             End If
8             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", lngSampleID)
9         Else
10            strSQL = "   Select /*+ rule */" & vbNewLine & _
                     "  Distinct '' ���,a.�걾id, a.������Ŀid, a.����, a.�������, a.�̶���Ŀ, a.Id, a.������Ŀ, a.�ٴ�����, a.��д As Ӣ����, a.Cv," & vbNewLine & _
                     " �����־ , Decode(a.���ν��, '-', '���ԣ�-��', '+', '���ԣ�+��', '*', '*.**', a.���ν��) As ���, Rownum As ���, a.��־, a.����id, a.�걾���," & vbNewLine & _
                     "   a.����ʱ��, a.�걾���, a.�걾����ʾ, a.���鱸ע, a.����, a.�Ա�, a.����, a.�����, a.סԺ��, a.��ǰ����, a.��ҳid, a.�����Χ, Nvl(g.С��λ��, 2) As С��," & vbNewLine & _
                     "    a.��������, a.��������, a.��λ," & vbNewLine & _
                     "   a.����ο� As �ο�, a.Od, a.Cutoff, a.Cov, a.ø���id, a.���챨��, a.���쾯ʾ, a.�������," & vbNewLine & _
                     "   A.����ο�,'' �Ƿ����,a.������Ŀ" & vbNewLine & _
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
                     " From ����걾��¼ A, ����걾��¼ E, ������ͨ��� B, ����������Ŀ C, ������Ŀ D, ������ĿĿ¼ H" & vbNewLine & _
                     " Where a.Id = b.����걾id And b.������Ŀid = c.Id And c.Id = d.������Ŀid And b.������Ŀid = h.Id(+) And b.��¼���� = a.������ And" & vbNewLine & _
                     "       e.Id = a.�ϲ�id And e.id = [1]) A, ����������Ŀ G" & vbNewLine & _
                     "  Where a.����id = g.����id(+) And a.Id = g.��Ŀid(+)" & vbNewLine & _
                     "  Order By a.������ĿID,a.�������,a.����"
12            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���鼼ʦվ", lngSampleID)

13        End If
          '    rsTmp.Sort = "�������"
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
59            .ColHidden(.ColIndex("�Ƿ����")) = True
60            If vsfLeft.Row > 0 Then
61                If Me.vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("���������")) <> "" Then
62                    For i = 1 To .Rows - 1
63                        If Val(.TextMatrix(i, .ColIndex("�Ƿ����"))) = 1 Then
64                            .Cell(flexcpPicture, i, .ColIndex("���")) = Me.imgVsf.ListImages("�������").ExtractIcon
65                            .Cell(flexcpPictureAlignment, i, .ColIndex("���")) = flexAlignRightCenter
66                            .RowHidden(i) = False
67                        Else
68                            .RowHidden(i) = True
69                        End If
70                    Next
71                End If
72            End If

73            For i = 1 To .Rows - 1
74                If .Cell(flexcpFontBold, i, .ColIndex("���")) = True Then
75                    .RowHidden(i) = IIf(Me.chkGroup.value = 1, False, True)
76                End If
77            Next
78        End With

79        CalcReferenceColour

80        Exit Sub
ReadSampleVal_Error:
81        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(ReadSampleVal)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
82        Err.Clear

End Sub

Private Sub ReadSampleBacteriology(lngSampleID As Long, Optional intVal As Integer = 25)
      '����   ��������Ϣ
          Dim strErr As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

1         On Error GoTo ReadSampleBacteriology_Error

2         If intVal = 25 Then
3             strSQL = "select b.id,b.������ || '(' || b.Ӣ���� || ')' ϸ����,a.������,a.�������� ����," & vbNewLine & _
                     "       a.��ҩ����,a.���id," & vbNewLine & _
                       "a.����ʱ��,a.������,a.δ���,a.��������,a.���²���,a.��ϸ��,a.�����豸,a.������," & _
                       "a.����δ���,a.��������,a.��������,a.�����־,a.ϸ��ID,a.�Ƿ񾵼���, a.�������" & vbNewLine & _
                       "from ���鱨��ϸ�� a,����ϸ����¼ b" & vbNewLine & _
                       "where a.ϸ��id = b.id(+) and a.�걾id = [1] "

4             Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", lngSampleID)
5         Else
6             strSQL = "SELECT Distinct B.����, B.ID ϸ��id ,D.������,B.������ AS ϸ����, " & _
                       "A.������ AS ������,A.�������� as ����,A.��ҩ����, d.���鱸ע,d.��ע " & _
                       "FROM ������ͨ��� A,����ϸ�� B,����걾��¼ D  " & _
                       "WHERE A.ϸ��id = B.ID And D.����� is Not null  " & _
                       "AND A.��¼���� = [1]  " & _
                       "AND D.ID=A.����걾ID  " & _
                       "AND D.ID= [2] Order by B.����"
7             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���鼼ʦվ", mlngValueC, lngSampleID)
8         End If
          '    rsTmp.Sort = "�������"
9         If Not vfgLoadFromRecord(VsfMicrobe, rsTmp, strErr, imgVsf) Then Exit Sub

10        With VsfMicrobe
11            If intVal = 25 Then
12                .ColWidth(.ColIndex("ϸ����")) = 3000: .ColHidden(.ColIndex("ϸ����")) = False
13                .ColWidth(.ColIndex("������")) = 2000: .ColHidden(.ColIndex("������")) = False
14                .ColWidth(.ColIndex("����")) = 3000: .ColHidden(.ColIndex("����")) = False
15                .ColWidth(.ColIndex("��ҩ����")) = 3000: .ColHidden(.ColIndex("��ҩ����")) = False
16                If rsTmp.RecordCount > 0 Then
17                    rsTmp.MoveFirst

18                    Me.txtNormalMicrobe = rsTmp("������") & ""
19                    Me.txtNoFindMicrobe = rsTmp("δ���") & ""
20                    Me.txtNormalMicrobes = rsTmp("��������") & ""
21                    Me.chkPathopoiesiaGerm.value = IIf(rsTmp("���²���") = 1, 1, 0)
22                    Me.chkNoGerm.value = IIf(rsTmp("��ϸ��") = 1, 1, 0)
23                    Me.txtMicroscope = rsTmp("�����豸") & ""
24                    Me.txtMicroscopeFinded = rsTmp("������") & ""
25                    Me.txtMicroscopeNOFind = rsTmp("����δ���") & ""
26                    Me.txtMicrobePositiveComment = rsTmp("��������") & ""
27                    Me.txtGermComment = rsTmp("��������") & ""
28                    If Val(rsTmp("�Ƿ񾵼���") & "") = 0 Then
29                        chkMicroscope.value = 0
30                    Else
31                        chkMicroscope.value = 1
32                    End If
33                    If Val(rsTmp("�������") & "") = 0 Then
34                        optReport(1).value = True
35                    Else
36                        optReport(0).value = True
37                    End If
38                    optReportShow

39                    ReadSampleAntibiotic mlngKey, Val(rsTmp("ϸ��ID") & "")
40                End If
41            Else
42                .ColWidth(.ColIndex("ϸ����")) = 3000: .ColHidden(.ColIndex("ϸ����")) = False
43                .ColWidth(.ColIndex("������")) = 2000: .ColHidden(.ColIndex("������")) = False
44                .ColWidth(.ColIndex("����")) = 3000: .ColHidden(.ColIndex("����")) = False
45                .ColWidth(.ColIndex("��ҩ����")) = 3000: .ColHidden(.ColIndex("��ҩ����")) = False
46                If rsTmp.RecordCount > 0 Then
47                    rsTmp.MoveFirst
48                    txtMicrobePositiveComment = rsTmp("��ע") & ""
49                    txtComment = rsTmp("���鱸ע") & ""
50                    ReadSampleAntibiotic mlngKey, Val(rsTmp("ϸ��ID") & ""), 10
51                End If
52            End If
53        End With


54        Exit Sub
ReadSampleBacteriology_Error:
55        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(ReadSampleBacteriology)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
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
          '����           ���뿹����д��VSF
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strErr As String
          Dim i As Integer

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
14            .ColWidth(.ColIndex("ҩ������")) = 1500: .ColHidden(.ColIndex("ҩ������")) = False

15            If .Rows > 1 And Not mrsAntibioticValType Is Nothing And intVal = 25 Then
16                For i = 1 To .Rows - 1
17                    If InStr(.TextMatrix(i, .ColIndex("�������")), "-") > 0 Then
                          'ҩ�����������ɫ��ʶ
18                        mrsAntibioticValType.Filter = ""
19                        mrsAntibioticValType.Filter = "����='" & Split(.TextMatrix(i, .ColIndex("�������")), "-")(0) & "'"
20                        .Cell(flexcpBackColor, i, .ColIndex("�������"), i, .ColIndex("�������")) = Val(mrsAntibioticValType("��ɫ") & "")
21                        mrsAntibioticValType.Filter = ""
22                    Else
23                        .Cell(flexcpBackColor, i, .ColIndex("�������"), i, .ColIndex("�������")) = 0
24                    End If
25                Next
26            End If
27        End With


28        Exit Sub
ReadSampleAntibiotic_Error:
29        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(ReadSampleAntibiotic)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
30        Err.Clear

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
3             blnTre = IsTre(lngSampleID)

4             If blnTre Then
5                 strSQL = "Select b.id, b.������, b.Ӣ����, b.��λ, a.id ����, c.����ʱ��, a.������, e.����ʱ��, b.���챨����, b.�������, a.�����־" & vbNewLine & _
                           "   From ���鱨����ϸ A, ����ָ�� B, ���鱨���¼ C, ��������걾 D, ��������ʱ�䷽�� E" & vbNewLine & _
                           "   Where A.��ĿID = B.ID And A.�걾ID = C.ID And A.ID = D.������ϸid And D.���ܷ���id = e.ID And A.�걾ID = [1]" & vbNewLine & _
                           "   Order By a.id Desc"
6                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ȶ�����", lngSampleID)
7             Else
8                 strSQL = "Select " & vbNewLine & _
                           " B.Id, B.������, B.Ӣ����, B.��λ, A.����, A.����ʱ��, A.������, B.���챨����, B.�������, A.�����־" & vbNewLine & _
                           "From (Select B.��Ŀid ������Ŀid, B.����, B.����ʱ��, B.������, B.�����־" & vbNewLine & _
                           "       From (Select A.Id ����, A.����id, A.�걾����, A.����ʱ��, B.��Ŀid" & vbNewLine & _
                           "              From ���鱨���¼ A, ���鱨����ϸ B" & vbNewLine & _
                           "              Where A.Id = B.�걾id And A.Id = [1] and b.������ is not null ) A," & vbNewLine & _
                           "            (Select A.Id ����, A.����id, A.�걾����, A.����ʱ��, B.��Ŀid, B.������, B.�����־" & vbNewLine & _
                           "              From ���鱨���¼ A, ���鱨����ϸ B" & vbNewLine & _
                           "              Where A.Id = B.�걾id And A.����id = [2] And ����ʱ�� Between [3] And [4] and a.id <= [1] and b.������ is not null ) B" & vbNewLine & _
                           "       Where A.����id = B.����id And A.��Ŀid + 0 = B.��Ŀid And Nvl(A.�걾����, 0) = Nvl(B.�걾����, 0) ) A, ����ָ�� B" & vbNewLine & _
                           "Where A.������Ŀid = B.Id" & vbNewLine & _
                           "Order By LPad(B.�������, 10, '0'),b.id, A.���� desc "

9                 Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ȶ�����", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
                                         CDate(Format(SampleReportDate, "yyyy-MM-dd") & " 23:59:59"))
10            End If
11        Else
12            strSQL = "    Select " & vbNewLine & _
                       "       i.Id, i.���� As ������, v.��д As Ӣ����, i.���㵥λ As ��λ, a.����, a.����ʱ��, a.������, v.���챨����, v.�������, a.�����־" & vbNewLine & _
                       "       From (Select b.������Ŀid, b.����, b.����ʱ��, b.������, b.�����־" & vbNewLine & _
                       "              From (Select a.Id ����, a.����id, a.�걾����, a.���ʱ�� ����ʱ��, b.������Ŀid, b.������, b.�����־" & vbNewLine & _
                       "                     From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                       "                     Where a.Id = b.����걾id And a.Id = [1] And ����id = [2] And b.������ Is Not Null) A," & vbNewLine & _
                       "                   (Select a.Id ����, a.����id, a.�걾����, a.���ʱ�� ����ʱ��, b.������Ŀid, b.������, b.�����־" & vbNewLine & _
                       "                     From ����걾��¼ A, ������ͨ��� B" & vbNewLine & _
                       "                     Where a.Id = b.����걾id And a.Id < [1] And ����id = [2]  And  ���ʱ�� Between [3] And [4]  And b.������ Is Not Null) B" & vbNewLine & _
                       "              Where a.����id = b.����id And a.������Ŀid + 0 = b.������Ŀid) A, ������Ŀ V, ���鱨����Ŀ R, ������ĿĿ¼ I" & vbNewLine & _
                       "       Where A.������Ŀid = v.������Ŀid And A.������Ŀid = r.������Ŀid And r.������Ŀid = i.ID And i.�����Ŀ <> 1"
13            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "����ȶ�����", lngSampleID, lngPatientID, CDate(Format(SampleReportDate - intMaxDay, "yyyy-MM-dd") & " 00:00:00"), _
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
22            .TextMatrix(0, 0) = "������Ŀ": .ColWidth(0) = 2500: .RowHeight(0) = 800
23            Do Until rsTmp.EOF
24                If lngItemid <> rsTmp("ID") Then
25                    .Rows = .Rows + 1
26                    intCol = 0
27                    If .Cols - 1 < intCol Then
28                        .Cols = .Cols + 1
29                        .ColWidth(intCol) = 1500
30                    End If

31                    If intCol = 0 Then
                          'д����Ŀ
32                        .TextMatrix(.Rows - 1, intCol) = rsTmp("������") & "(" & rsTmp("Ӣ����") & ")"

33                    End If
34                    intCol = intCol + 1
35                    If .Cols - 1 < intCol Then
36                        .Cols = .Cols + 1
37                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter

38                        If blnTre Then
39                            .TextMatrix(0, intCol) = rsTmp("����ʱ��") & ""
40                        Else
41                            .TextMatrix(0, intCol) = "����(" & Mid(Mid(rsTmp("����ʱ��"), 3), 1, Len(Mid(rsTmp("����ʱ��"), 3)) - 3) & ")"
42                        End If


43                    End If
                      'д������
44                    .TextMatrix(.Rows - 1, intCol) = rsTmp("������") & ""
45                    .Cell(flexcpBackColor, .Rows - 1, intCol) = GetValColour(Val(rsTmp("�����־" & "")))

46                Else
47                    intCol = intCol + 1
48                    If .Cols - 1 < intCol Then
49                        .Cols = .Cols + 1
50                        .ColWidth(intCol) = 1200: .ColAlignment(intCol) = flexAlignLeftCenter
51                        If blnTre Then
52                            .TextMatrix(0, intCol) = rsTmp("����ʱ��") & ""
53                        Else
54                            .TextMatrix(0, intCol) = "��" & intCol - 1 & "��(" & Mid(Mid(rsTmp("����ʱ��"), 3), 1, Len(Mid(rsTmp("����ʱ��"), 3)) - 3) & ")"
55                        End If
56                        dblTmp = Val(CalcVolatility(.TextMatrix(.Rows - 1, 1), .TextMatrix(.Rows - 1, intCol)))
57                        If dblTmp <> 0 And Val(rsTmp("���챨����") & "") <> 0 Then
58                            If dblTmp > Val(rsTmp("���챨����") & "") Then
59                                .Cell(flexcpBackColor, .Rows - 1, intCol) = RGB(248, 194, 169)
60                            End If
61                        End If
62                    End If
                      'д������
63                    .TextMatrix(.Rows - 1, intCol) = rsTmp("������") & ""
64                    .Cell(flexcpBackColor, .Rows - 1, intCol) = GetValColour(Val(rsTmp("�����־" & "")))
65                End If
66                lngItemid = rsTmp("ID")
67                rsTmp.MoveNext
68            Loop
69        End With

70        LoadContrastDBWriteVSF = True


71        Exit Function
LoadContrastDBWriteVSF_Error:
72        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(LoadContrastDBWriteVSF)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
73        Err.Clear

End Function

Private Sub ReadHistorData()
    '����           �������ε�����
    Dim strErr As String
    Call LoadContrastDBWriteVSF(VSFContrast, mlngKey, mlngPatientID, mReportDate, 60, strErr)
End Sub


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
56        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(LoadVSFContrastToCht)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
57        Err.Clear

End Function
Private Sub ReadContrastToVsf()
    '����       �������αȶԵ�VSF
    Dim strErr As String

    Me.VSFContrast.Rows = 1: Me.VSFContrast.Rows = 2


    'û�в���IDʱ�˳�
    If mlngPatientID = 0 Then Exit Sub

    Call LoadContrastDBWriteVSF(Me.VSFContrast, mlngKey, mlngPatientID, mReportDate, Val(txtMaxDay), strErr)
    Call VSFContrast_SelChange
End Sub

Private Sub InitFace()
    '����           ��ʼ������
    '========================================��ʾ��ɫ����============================================
    '��ʾ��ɫ����
    gSampleShowColour.���� = &H80000005
    gSampleShowColour.ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾƫ����ɫ", 2500, 2500, "8438015"))
    gSampleShowColour.ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾƫ����ɫ", 2500, 2500, "8454143"))
    gSampleShowColour.��ʾƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ��ʾƫ����ɫ", 2500, 2500, "255"))
    gSampleShowColour.��ʾƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ��ʾƫ����ɫ", 2500, 2500, "255"))
    gSampleShowColour.����ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ����ƫ����ɫ", 2500, 2500, "65280"))
    gSampleShowColour.����ƫ�� = Val(ComGetPara(Sel_Lis_DB, "��ʾ����ƫ����ɫ", 2500, 2500, "12648384"))
    gSampleShowColour.�쳣 = Val(ComGetPara(Sel_Lis_DB, "��ʾ�쳣��ɫ", 2500, 2500, "16576"))

    picGeneral.Visible = True
    picMicrobePositive.Visible = False
    PicNegative.Visible = False
End Sub

Private Sub CalcReferenceColour()
          '����           ����������ɫ
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
                          
12                        .Cell(flexcpBackColor, intRow, .ColIndex("���"), intRow, .ColIndex("���")) = GetValColour(Val(.TextMatrix(intRow, .ColIndex("�����־"))))
13                        If Val(vsfLeft.TextMatrix(vsfLeft.Row, vsfLeft.ColIndex("�汾"))) = 25 Then
14                            .Cell(flexcpBackColor, intRow, .ColIndex("�ϴ�"), intRow, .ColIndex("�ϴ�")) = GetValColour(Val(.TextMatrix(intRow, .ColIndex("�ϴα�־"))))
15                        End If
16                    End If
17                End If
18            Next
19        End With
          
          
20        Exit Sub
CalcReferenceColour_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(CalcReferenceColour)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
22        Err.Clear
End Sub

Private Sub vsfCenter_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
          Dim intCol As Integer

1         On Error GoTo vsfCenter_AfterRowColChange_Error

2         If OldRow <> NewRow Or OldCol <> NewCol Then
3             With vsfCenter
4                 If .ColIndex("id") <> -1 Then
5                     If Val(.TextMatrix(.Row, .ColIndex("id"))) <> 0 Then
6                         txtSignificance.Text = .TextMatrix(.Row, .ColIndex("�ٴ�����"))
7                     End If
8                 End If
9             End With
10        End If

11        Exit Sub
vsfCenter_AfterRowColChange_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(vsfCenter_AfterRowColChange)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
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
                If lngCol = .ColIndex("ѡ��") Then
                    If .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 1 And .TextMatrix(lngRow, lngCol) = "" Then
                        .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 2
                    ElseIf .TextMatrix(lngRow, lngCol) = "" Then
                        .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 1
                    End If
                End If
            End If
        End If

        If Button = 2 Then
            If lngRow = 0 Then
                Call GetCursorPos(Point)
                strTitle = SetVsfColHiden(Me, Me.vsfLeft, Point.X * 15, Point.Y * 15, "���鱨����ʾ��", 2500, 2001, "�������,�汾")
                If strTitle <> "" Then
                    SaveDBLog 18, 6, 0, "���������", "���ñ���е���ʾ������:" & strTitle, 2500, "�ٴ�ʵ���ҹ���"
                    Call ReadPatientList
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
7                         mlngKey = Val(.TextMatrix(.Row, .ColIndex("ID")))
8                         mlngPatientID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
9                         mReportDate = .TextMatrix(.Row, .ColIndex("����ʱ��"))
10                        txtComment = .TextMatrix(.Row, .ColIndex("��ע"))
11                        txtDiagnose = .TextMatrix(.Row, .ColIndex("���"))
12                        mlngValueC = .TextMatrix(.Row, .ColIndex("�������"))
13                        mintVer = .TextMatrix(.Row, .ColIndex("�汾"))
14                        Me.txtResultComment.Text = .TextMatrix(.Row, .ColIndex("���˵��"))
15                        If Val(.TextMatrix(.Row, .ColIndex("΢����"))) = 1 Then
16                            If Val(.TextMatrix(.Row, .ColIndex("���Ա���"))) = 1 Then
17                                picGeneral.Visible = False
18                                picMicrobePositive.Visible = True
19                                PicNegative.Visible = False
20                                If Val(.TextMatrix(.Row, .ColIndex("�汾"))) = 25 Then
21                                    ReadSampleBacteriology mlngKey, 25
22                                Else
23                                    ReadSampleBacteriology mlngKey, 10
24                                End If
25                            Else
26                                picGeneral.Visible = False
27                                picMicrobePositive.Visible = False
28                                PicNegative.Visible = True
29                                If Val(.TextMatrix(.Row, .ColIndex("�汾"))) = 25 Then
30                                    ReadSampleBacteriology mlngKey
31                                Else

32                                End If
33                            End If
34                        Else
35                            picGeneral.Visible = True
36                            picMicrobePositive.Visible = False
37                            PicNegative.Visible = False
38                            If Val(.TextMatrix(.Row, .ColIndex("�汾"))) = 25 Then
39                                ReadSampleVal mlngKey, 25

40                            Else
41                                ReadSampleVal mlngKey, 10
42                            End If
43                        End If

                          '���䱨��
44                        If Val(.TextMatrix(.Row, .ColIndex("���䱨��"))) = 3 Then
45                            picSupplement.Visible = True
46                            Call GetSupplementReport(mlngKey, vsfSupplement)    '��ȡ���䱨��
47                            Call EditSampleValueList(vsfCenter, vsfSupplement)
48                        Else
49                            picSupplement.Visible = False
50                        End If
51                        Call picGeneral_Resize


52                        If .TextMatrix(.Row, .ColIndex("����")) <> "1" Then
53                            Call funWriteAdvicesLookState(.TextMatrix(.Row, .ColIndex("����ID")), 1)
54                            .TextMatrix(.Row, .ColIndex("����")) = 1
55                            .Cell(flexcpPicture, .Row, .ColIndex("����")) = imgVsf.ListImages("����").ExtractIcon
56                        End If
      '                    mlngKey = 0
57                        RefreshTab Me.TabPage.Selected.Index
58                    End If
59                End If
60            End If
61        End With


62        Exit Sub
vsfLeft_SelChange_Error:
63        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(vsfLeft_SelChange)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
64        Err.Clear
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
    On Error GoTo ReadImages_Error

    Call ImageTypeSet(9, True)
    '����ͼ������
    If ReadSampleImage(lngSampleID, strChart, strErr, intVal) = False Then Exit Sub
    For intloop = 0 To 8
        If strChart(intloop) <> "" Then
            chtPic(intloop).Load (strChart(intloop))
        End If
    Next
    '����������Ű�
    Call ImageTypeSet(9)


    Exit Sub
ReadImages_Error:
    Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(ReadImages)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear

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
16            End If
17        End If
18        PtintOldReport = True


19        Exit Function
PtintOldReport_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(PtintOldReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear
End Function




Public Function PrintReport(objFrm As Object, lngSampleID As Long, Optional byRunMode As Byte = 2, Optional intRow As Integer, Optional lngPrintCount As Long, Optional strErr As String) As Boolean
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

2         strSQL = "select b.id ����id ,b.���� ��������,b.�������,Nvl(a.������Դ,1) ������Դ,a.����ʱ��,a.���Ա���,a.�걾���,a.ҽ��վ��ӡ from ���鱨���¼ a,����������¼ b where a.����id = b.id and a.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ӡ", lngSampleID)

4         If rsTmp.RecordCount = 0 Then Exit Function

          '�Աȴ�ӡ�����Ͳ���
5         If lngPrintCount > 0 Then
6             If Val(rsTmp("ҽ��վ��ӡ") & "") >= lngPrintCount And Val(rsTmp("������Դ") & "") = 2 Then
7                 With Me.vsfLeft
8                     .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
9                     .Cell(flexcpPicture, intRow, .ColIndex("��ӡ")) = imgVsf.ListImages("��ֹ��ӡ").ExtractIcon
10                End With
11                PrintReport = False
12                Exit Function
13            End If
14        End If

15        strSQL = "select id,����,����,���ﵥ��,סԺ����,��쵥��,Ժ�ⵥ��,�����ʽ,סԺ��ʽ,����ʽ,Ժ���ʽ,��ʽ����," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(���ﵥ��, '00000')) || '-2' ���ﵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ����, '00000')) || '-2' סԺ���ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(��쵥��, '00000')) || '-2' ��쵥�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ�ⵥ��, '00000')) || '-2' Ժ�ⵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(�����ʽ, '00000')) || '-2' �����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ��ʽ, '00000')) || '-2' סԺ��ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(����ʽ, '00000')) || '-2' ����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ���ʽ, '00000')) || '-2' Ժ���ʽ��" & vbNewLine & _
                      "from ����������¼ where id = [1] "

16        Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", Val(rsTmp("����ID") & ""))


17        rsReportFormat.Filter = "id=" & Val(rsTmp("����ID") & "")
18        If Val(rsTmp("�������")) = 1 Then
19            If Val(rsTmp("���Ա���") & "") = 1 Then
                  '����
20                intSel = 0
21            Else
                  '����
22                intSel = 1
23            End If
24        Else
25            intCount = GetSampleValCount(lngSampleID)
              'û�н��ʱ��ʾ
26            If intCount = 0 Then
27                Exit Function
28            End If
29            If rsReportFormat.RecordCount > 0 Then
30                If Val(rsReportFormat("��ʽ����") & "") > 0 Then
31                    If intCount > Val(rsReportFormat("��ʽ����") & "") Then
32                        intSel = 0
33                    Else
34                        intSel = 1
35                    End If
36                End If
37            Else
38                intSel = 0
39            End If

40        End If
41        Select Case Val(rsTmp("������Դ"))
              Case 1
42                If intSel = 0 Then
43                    strNO = rsReportFormat("���ﵥ�ݺ�")
44                Else
45                    strNO = rsReportFormat("�����ʽ��")
46                End If
47            Case 2
48                If intSel = 0 Then
49                    strNO = rsReportFormat("סԺ���ݺ�")
50                Else
51                    strNO = rsReportFormat("סԺ��ʽ��")
52                End If
53            Case 3
54                If intSel = 0 Then
55                    strNO = rsReportFormat("סԺ���ݺ�")
56                Else
57                    strNO = rsReportFormat("סԺ��ʽ��")
58                End If
59            Case 4
60                If intSel = 0 Then
61                    strNO = rsReportFormat("Ժ�ⵥ�ݺ�")
62                Else
63                    strNO = rsReportFormat("Ժ���ʽ��")
64                End If
65            Case Else
66                If intSel = 0 Then
67                    strNO = rsReportFormat("���ﵥ�ݺ�")
68                Else
69                    strNO = rsReportFormat("�����ʽ��")
70                End If
71        End Select
72        If byRunMode = 3 Then
73            If strNO <> "" Then
74                FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, objFrm
75            End If
76        Else
             '��ͼ��
77            strTmp = "��ʼ����ͼ��:" & Now & vbCrLf
78            If ReadSampleImage(lngSampleID, strChart, strErr, 25) = False Then
79                Exit Function
80            End If
81            strTmp = strTmp & "����ͼ�����:" & Now & vbCrLf

82            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, objFrm, "�걾ID=" & lngSampleID, "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), _
                      "ͼ��4=" & strChart(3), "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                      "ͼ��9=" & strChart(8), byRunMode
83            strTmp = strTmp & "��ӡ���:" & Now & vbCrLf

              '������˹��ı걾��ʶ
84            strSQL = "Zl_���鱨���ӡ_Edit(1," & lngSampleID & ",1)"
85            Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
86            strTmp = strTmp & "��ɴ�ӡ:" & Now

87            SaveDBLog 18, 6, lngSampleID, "��ӡ", "�����ӡ", 2500, "�ٴ�ʵ���ҹ���"
88        End If

89        PrintReport = True

          '����ˢ�¿��ڸſ��Ѵ�ӡ��ǩ����
90        Call SendMessage("RefreshDeptSurvey7")


91        Exit Function
PrintReport_Error:
92        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(PrintReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
93        Err.Clear
End Function


Private Sub GetDept(Optional intType As Integer)
          '����               ������һ���
          '����               intType 0=���� 1=����
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
1         On Error GoTo GetDept_Error

2         If intType = 0 Then
3             strSQL = "Select b.id,C.����, C.����" & vbNewLine & _
                      "From ������Ա A, ��Ա�� B, ���ű� C" & vbNewLine & _
                      "Where A.��Աid = B.Id And A.����id = C.Id And (C.����ʱ�� Is Null Or C.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) and a.��Աid = [1] "
4         Else
5             strSQL = ""
6         End If
7         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "�������", gUserInfo.ID)
      '    With cboPatient
      '        .Clear
      '        .AddItem "���п���"
      '        Do Until rsTmp.EOF
      '            .AddItem Trim(rsTmp("����")) & "-" & Trim(rsTmp("����")) & ""
      '            .ItemData(.NewIndex) = rsTmp("id")
      '            rsTmp.MoveNext
      '        Loop
      '        .ListIndex = 0
      '    End With


8         Exit Sub
GetDept_Error:
9         Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetDept)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
10        Err.Clear
End Sub
Private Function InitDepts(intDeptView As Integer, Optional strErr As String) As Boolean
      '���ܣ���ʼ��סԺ�ٴ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          Dim strDeptIDs As String, lngPreDept As Long

1         On Error GoTo InitDepts_Error

2         If cboDept.ListIndex <> -1 Then
3             lngPreDept = cboDept.ItemData(cboDept.ListIndex)
4         End If

5         If intDeptView = 0 Then
              '�����Ҷ�ȡ��ʾ
              '�����ż���۲��ҵĲ��˻�û���ϴ�������ֻ�Դ����в��˵Ŀ��ҵ�����
6             If InStr(mstrPrivs, "ȫԺ����") > 0 Then
7                 strDeptIDs = GetUser����IDs
8                 strSQL = _
                      " Select Distinct A.ID,A.����,A.����" & _
                      " From ���ű� A,��������˵�� B" & _
                      " Where B.����ID=A.ID And B.��������='�ٴ�'" & _
                      " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                      " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                      " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)" & _
                      " Order by A.����"
9             Else
                  '����Ȩ�޵Ŀ��ң��������ڿ���+�������������Ŀ���
10                strSQL = _
                      " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                      " From ���ű� A,��������˵�� B,������Ա C" & _
                      " Where B.����ID=A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                      " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                      " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                      " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)" & _
                      " And B.��������='�ٴ�'"
11                strSQL = strSQL & " Union " & _
                      " Select C.ID,C.����,C.����,Nvl(A.ȱʡ,0) As ȱʡ" & _
                      " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                      " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                      " And Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                      " And Not Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                      " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                      " And (C.վ��='" & gUserInfo.NodeNo & "' Or C.վ�� is Null)"
12                If InStr(mstrPrivs, "ICU����") > 0 Then
13                    strSQL = strSQL & " Union " & _
                          " Select A.ID,A.����,A.����,0 As ȱʡ" & _
                          " From ���ű� A" & _
                          " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                          " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='�ٴ�')" & _
                          " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                          " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)"
14                End If
15                strSQL = "Select ID,����,����,Max(ȱʡ) As ȱʡ From (" & strSQL & ") Group By ID,����,���� Order by ����"
16            End If
17        Else
              '��������ȡ��ʾ
18            If InStr(mstrPrivs, "ȫԺ����") > 0 Then
19                strDeptIDs = GetUser����IDs
20                strSQL = _
                      " Select Distinct A.ID,A.����,A.����" & _
                      " From ���ű� A,��������˵�� B " & _
                      " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
                      " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                      " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)" & _
                      " Order by A.����"
21            Else
                  '����Ȩ������ֱ�����ڲ���+���ڿ�����������
22                strSQL = _
                      " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                      " From ���ű� A,��������˵�� B,������Ա C" & _
                      " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                      " And B.������� in(1,2,3) And B.��������='����'" & _
                      " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                      " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)"
23                strSQL = strSQL & " Union " & _
                      " Select C.ID,C.����,C.����,Nvl(A.ȱʡ,0) as ȱʡ" & _
                      " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                      " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                      " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                      " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                      " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                      " And (C.վ��='" & gUserInfo.NodeNo & "' Or C.վ�� is Null)"
24                If InStr(mstrPrivs, "ICU����") > 0 Then
25                    strSQL = strSQL & " Union " & _
                          " Select A.ID,A.����,A.����,0 As ȱʡ" & _
                          " From ���ű� A" & _
                          " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                          " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='����')" & _
                          " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                          " And (A.վ��='" & gUserInfo.NodeNo & "' Or A.վ�� is Null)"
26                End If
27                strSQL = "Select ID,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,���� Order by ����"
28            End If
29        End If

30        cboDept.Clear
31        If InStr(mstrPrivs, "���п���") > 0 Then cboDept.AddItem "00-���п���"
32        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, gUserInfo.ID)

33        For i = 1 To rsTmp.RecordCount
34            cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
35            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
36            rsTmp.MoveNext
37        Next
38        If rsTmp.RecordCount > 0 Then
39            cboDept.ListIndex = 0
40        End If
41        InitDepts = True


42        Exit Function
InitDepts_Error:
43        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(InitDepts)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
44        Err.Clear

End Function
Public Function GetUser����IDs(Optional ByVal bln���� As Boolean, Optional strErr As String) As String
      '���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
      '�������Ƿ�ȡ���������µĿ���
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser����IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
          'û��ǿ�������ٴ�,����ҽ��������
7         If blnNew Then
8             strSQL = "Select 1 as ���,����ID From ������Ա Where ��ԱID=[1] Union" & _
                      " Select Distinct 2 as ���,B.����ID From ������Ա A,�������Ҷ�Ӧ B" & _
                      " Where A.����ID=B.����ID And A.��ԱID=[1]"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", gUserInfo.ID)
10        End If
11        If bln���� = False Then
12            rsTmp.Filter = "��� = 1"
13        Else
14            rsTmp.Filter = ""
15        End If

16        For i = 1 To rsTmp.RecordCount
17            If InStr("," & GetUser����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
18                GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
19            End If
20            rsTmp.MoveNext
21        Next
22        GetUser����IDs = Mid(GetUser����IDs, 2)


23        Exit Function
GetUser����IDs_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetUser����IDs)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
25        Err.Clear

End Function
Public Function GetUser����IDs(Optional strErr As String) As String
      '���ܣ���ȡ����Ա�����Ĳ���(ֱ�����ڲ��������ڿ��������Ĳ���),�����ж��
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser����IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
7         If blnNew Then
8             strSQL = _
                  "Select Distinct ����ID From (" & _
                  " Select A.����ID as ����ID" & _
                  " From ��������˵�� A,������Ա B" & _
                  " Where A.����ID=B.����ID And B.��ԱID=[1]" & _
                  " And A.������� in(1,2,3) And A.��������='����'" & _
                  " Union" & _
                  " Select A.����ID From �������Ҷ�Ӧ A,������Ա B" & _
                  " Where A.����ID=B.����ID And B.��ԱID=[1])"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", gUserInfo.ID)
10        ElseIf rsTmp.RecordCount > 0 Then
11            rsTmp.MoveFirst
12        End If
13        For i = 1 To rsTmp.RecordCount
14            GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
15            rsTmp.MoveNext
16        Next

17        GetUser����IDs = Mid(GetUser����IDs, 2)


18        Exit Function
GetUser����IDs_Error:
19        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetUser����IDs)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear

End Function

Private Function GetDepts(lngID As Long) As String
          '����           ͨ������IDȡ�ò��������п��ҵĿ��ұ��룬��","�ָ�
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
1         On Error GoTo GetDepts_Error

2         strSQL = "select b.���� from �������Ҷ�Ӧ a,���ű� b where  a.����id = b.id and a.����id = [1] "
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngID)
4         Do Until rsTmp.EOF
5             GetDepts = GetDepts & "," & rsTmp("����")
6             rsTmp.MoveNext
7         Loop
8         If GetDepts <> "" Then
9             GetDepts = Mid(GetDepts, 2)
10        End If


11        Exit Function
GetDepts_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetDepts)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear
End Function

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
        lblDept.Caption = "���벡����"
        InitDepts 1
        If cboDept.ListCount > 0 Then
            CboFind cboDept, lngDeptDistrict
        End If
    Else
        lblDept.Caption = "������ҡ�"
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
        Me.Show , objFrm  '�������ʾ����ı߿����ʾ�ô���ΪǶ��ʽ���ã����ǵ���show����
    Else
        Call YSystemMenu(Me.hWnd)
    End If
    Set objOutFrm = Me

    Exit Function
errH:
    strErr = "������(ShowMe),������Ϣ:" & Err.Number & " " & Err.Description
End Function

Public Function getPartDor()

          Dim rsDeptDor As ADODB.Recordset

1         On Error GoTo getPartDor_Error

2         cbodor.Clear
3         Set rsDeptDor = GetDeptDor(cboDept.ItemData(cboDept.ListIndex))
4         With Me.cbodor
5             .AddItem "����"
6             .ItemData(.NewIndex) = 0
7             Do Until rsDeptDor.EOF
8                 .AddItem rsDeptDor("����") & ""
          '            .ItemData(.NewIndex) = rsTmp("HIS����ID")
9                 rsDeptDor.MoveNext
10            Loop
11            If .ListCount > 0 Then .ListIndex = 0
12        End With


13        Exit Function
getPartDor_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(getPartDor)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
15        Err.Clear

End Function


Private Sub BatchPrint(Optional byRunMode As Byte = 2)
          '����   ������ӡ
          Dim intRow As Integer
          Dim lngPrintCount As Long   'ҽ������վ�����ӡ����Ĵ���
          Dim blnPrint As Boolean     '�Ƿ��ӡ�ɹ�
          Dim blnContinue As Boolean  '�Ƿ������ӡ�����ɴ�ӡ�ı���

          '��ȡ�����е�ҽ��վ��ӡ����
1         On Error GoTo BatchPrint_Error

2         lngPrintCount = Val(ComGetPara(Sel_Lis_DB, "ҽ������վ�����ӡ����", 2500, 2500, 1))

3         With vsfLeft
4             For intRow = 1 To .Rows - 1
5                 If .Cell(flexcpChecked, intRow, .ColIndex("ѡ��"), intRow, .ColIndex("ѡ��")) = 1 Then
6                     If Val(.TextMatrix(intRow, .ColIndex("id"))) > 0 Then
7                         If Val(.TextMatrix(intRow, .ColIndex("�汾"))) = 25 Then
8                             If .TextMatrix(intRow, .ColIndex("���������")) <> "" And .TextMatrix(intRow, .ColIndex("���ʱ��")) = "" Then
9                                 If Not blnContinue Then
10                                    If MsgBox("<" & .TextMatrix(intRow, .ColIndex("������Ŀ")) & ">ֻ����˲���ָ��,�޷���ӡ,�Ƿ������ӡ�����ɴ�ӡ�ı���?", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
11                                        Exit Sub
12                                    Else
13                                        blnContinue = True
14                                    End If
15                                End If
16                            Else
17                                blnPrint = PrintReport(Me, Val(.TextMatrix(intRow, .ColIndex("id"))), byRunMode, intRow, lngPrintCount)
18                            End If
19                        Else
20                            blnPrint = PtintOldReport(Me, Val(.TextMatrix(intRow, .ColIndex("id"))), Val(.TextMatrix(intRow, .ColIndex("����id"))), byRunMode)
21                        End If
22                        If byRunMode = 2 And blnPrint = True Then
23                            .TextMatrix(intRow, .ColIndex("��ӡ")) = 1
24                            .Cell(flexcpPicture, intRow, .ColIndex("��ӡ")) = imgVsf.ListImages("��ӡ").ExtractIcon
25                        End If
26                    End If
27                End If
28            Next
29        End With


30        Exit Sub
BatchPrint_Error:
31        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(BatchPrint)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
32        Err.Clear
End Sub

Private Sub VsfMicrobe_SelChange()
    With VsfMicrobe
        If .ColIndex("ϸ��ID") <> -1 Then
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

Private Sub ShowType(intType As Integer)
    '����       ���inttype <> 2 ������������ʾ��ʽ
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
    lblNo.Caption = "�����"
    lblNo.Left = Label5.Left
    txtPatiNo.Left = dtpVS.Left
    lblTimeOut.Left = dtpVE.Left
    txtDay.Left = lblTimeOut.Left + 1140
    Line1.X1 = Line1.X1 - 2050
    Line1.X2 = Line1.X2 - 2050
End Sub

Private Function GetDeptDor(lngDeptID) As ADODB.Recordset
          '����           ������һ������ض�Ӧ��ҽ����¼��
          '����
          '               lngDeptID ����ID����ID
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
1         On Error GoTo GetDeptDor_Error

2         strSQL = "Select b.����" & vbNewLine & _
                   "From ������Ա A, ��Ա�� B, ���ű� C" & vbNewLine & _
                   "Where A.��Աid = B.Id And A.����id = C.Id And (C.����ʱ�� Is Null Or C.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) and c.id = [1] "

3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngDeptID)
4         Set GetDeptDor = rsTmp


5         Exit Function
GetDeptDor_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetDeptDor)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear
End Function

Private Function GetDeptPatients(intDeptType As Integer, lngDeptID) As String
          '����           ������һ������ض�Ӧ�Ĳ���ID��
          '����           intDeptType = 1 ���� =2 ����
          '               lngDeptID ����ID����ID
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim strPatients As String
1         On Error GoTo GetDeptPatients_Error

2         If txtDay > 99 Then
3             If MsgBox("¼��ĳ�Ժʱ���Ѷ���99�죬�鿴��س�Ժ���ˣ����ܻ��ʱ������" & vbCrLf & "�����Ƿ��������?", vbYesNo + vbQuestion + vbDefaultButton2, "�������") = vbNo Then
4                 Exit Function
5             End If
6         End If
7         If intDeptType = 1 Then
8             strSQL = "select ����id from ��Ժ���� where ����ID = [1] union all select ����id from ������ҳ where ��Ժ����ID = [1] and ��Ժ���� between sysdate - [3] and sysdate "
9         ElseIf intDeptType = 2 Then
10            strSQL = "select ����id from ��Ժ���� where ����ID = [1]  union all select ����id from ������ҳ where ��Ժ����ID = [1] and ��Ժ���� between sysdate - [3] and sysdate "
11        Else
12            strSQL = "select distinct a.����id from ��Ժ���� a,������ҳ b where a.����id = b.����id and b.��Ժ���� is null and  a.����ID = [1] and b.סԺҽʦ = [2] "
13        End If
14        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngDeptID, gUserInfo.Name, Val(txtDay.Text))
15        Do While Not rsTmp.EOF
16            If strPatients = "" Then
17                strPatients = rsTmp("����ID")
18            Else
19                strPatients = strPatients & "," & rsTmp("����ID")
20            End If
21            rsTmp.MoveNext
22        Loop
23        GetDeptPatients = strPatients


24        Exit Function
GetDeptPatients_Error:
25        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetDeptPatients)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
26        Err.Clear
End Function

Private Function GetPatientsList()
          '����               �Ѳ�����Ϣ���뵽ָ������������
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim intDeptType As Integer
          Dim strPatients As String
          Dim rsDeptDor As ADODB.Recordset

          Dim strPatientsCut As String
          Dim strLast As String

1         On Error GoTo GetPatientsList_Error

2         strSQL = "Select a.id,c.����,0 ѡ��,A.����, Decode(A.�Ա�, 1, '��', 2, 'Ů', 9, 'δ֪', '') �Ա�, A.����, C.���� ������Ŀ, " & _
                   " A.סԺ��, A.����,B.����ʱ��,a.����ID,a.����ʱ��,a.���ʱ��,a.��ע,a.���,a.΢����,a.���Ա���,a.ҽ��վ��ӡ " & vbNewLine & _
                  " From ���鱨���¼ A, ����������� B, ���������Ŀ C" & vbNewLine & _
                  " Where A.Id = B.�걾id And B.���id = C.Id(+) and (a.����� is not null or a.��������� is not null) "


3         strSQL = "select /*+ rule */ distinct a.HIS����ID,a.���� from ���鱨���¼ a where "

4         If mintPatientType = 2 Then

5             If InStr(mstrPrivs, "���Ʋ���") > 0 Then
6                 If cboDept <> "" Then
7                     If lblDept.Caption = "���벡����" Then
8                         intDeptType = 2
9                     Else
10                        intDeptType = 1
11                    End If
12                    strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
13                    strSQL = strSQL & "  a.his����id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
14                Else
15                    If lblDept.Caption = "���벡����" Then
16                        intDeptType = 2
17                    Else
18                        intDeptType = 1
19                    End If
20                    strPatients = GetDeptPatients(intDeptType, gUserInfo.DeptID)
21                    strSQL = strSQL & "  a.his����id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
22                End If
23            Else
24                intDeptType = 3
25                If cboDept <> "" Then
26                    strPatients = GetDeptPatients(intDeptType, cboDept.ItemData(cboDept.ListIndex))
27                    strSQL = strSQL & "  a.his����id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
28                Else
29                    strPatients = GetDeptPatients(intDeptType, gUserInfo.DeptID)
30                    strSQL = strSQL & "  a.his����id in (Select /*+cardinality(b,10)*/ b.Column_Value From Table(Cast(F_Num2list([1]) As Zltools.T_Numlist)) b) "
31                End If
32            End If

33            If strPatients <> "" Then
34                With Me.cboPatients
35                    .Clear
36                    .AddItem "����"
37                    .ItemData(.NewIndex) = 0
                      '��strPatients���ȴ���3500ʱ,��Ҫ�ֽ�
38                    Do While Len(strPatients) >= 3500
39                        strPatientsCut = Mid(strPatients, 1, 3500)
40                        strPatients = Mid(strPatients, 3501)
41                        strLast = Mid(strPatientsCut, InStrRev(strPatientsCut, ","))
42                        strPatientsCut = Mid(strPatientsCut, 1, InStrRev(strPatientsCut, ",") - 1)
43                        strPatients = strLast & strPatients

44                        If Mid(strPatientsCut, 1, 1) = "," Then
45                            strPatientsCut = Mid(strPatientsCut, 2)
46                        End If
47                        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���벡���б�", strPatientsCut)
48                        Do Until rsTmp.EOF
49                            .AddItem rsTmp("����") & ""
50                            .ItemData(.NewIndex) = rsTmp("HIS����ID")
51                            rsTmp.MoveNext
52                        Loop
53                    Loop
54                    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���벡���б�", strPatients)
55                    Do Until rsTmp.EOF
56                        .AddItem rsTmp("����") & ""
57                        .ItemData(.NewIndex) = rsTmp("HIS����ID")
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
69                .AddItem "����"
70                .ItemData(.NewIndex) = 0
71                Do Until rsDeptDor.EOF
72                    .AddItem rsDeptDor("����") & ""
          '            .ItemData(.NewIndex) = rsTmp("HIS����ID")
73                    rsDeptDor.MoveNext
74                Loop
75                If .ListCount > 0 Then .ListIndex = 0
76            End With
77        End If
78        strSQL = "select ���� from ������Ϣ where ����ID = [1]  "
79        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���벡���б�", mlngGetPatientID)
80        If rsTmp.EOF = False Then
81            If CheckCboID(Me.cboPatients, mlngGetPatientID) = False Then
82                Me.cboPatients.AddItem rsTmp("����") & ""
83                Me.cboPatients.ItemData(Me.cboPatients.NewIndex) = mlngGetPatientID
84                If cboPatients.ListCount = 1 Then
85                    cboPatients.ListIndex = 0
86                End If
87            End If
      '        Me.cboPatients.Text = rsTmp("����") & ""
88        End If


89        Exit Function
GetPatientsList_Error:
90        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetPatientsList)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
91        Err.Clear

End Function

Private Function CheckCboID(cboobj As ComboBox, lngID As Long) As Boolean
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '����                           ���CBO�ؼ���ID�Ƿ��ظ�
    '����
    '                               cboobj = cbo����
    '                               lngID = ��Ҫ����ID
    '����                           true = �ظ� false= ���ظ�
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
          '����   ���ֵ����ȡָ���ķ���
          Dim strSQL As String

1     On Error GoTo GetDictType_Error

2         strSQL = "Select С��, ����, ����, ����, ����, ��ע, ��ɫ From �����ֵ�� Where ���� = [1]"
3         Set GetDictType = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ֵ�", strType)


4         Exit Function
GetDictType_Error:
5         Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(GetDictType)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
6         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/5/25
'��    ��:����API��̬���ô����border
'��    ��:
'           new_Hwnd    ����ľ��
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-04-19
'��    ��:  ��ʾ���Ʋο�
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Sub ShowClincHelp()
          Dim lngSampleID As Long
          Dim lngVer As Long

1         On Error GoTo ShowClincHelp_Error

2         With Me.vsfLeft
3             If .Row < 1 Then
4                 MsgBox "��ѡ��һ�ݱ���", vbInformation, gSysInfo.AppName
5                 Exit Sub
6             End If
7             If Val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
8                 MsgBox "��ѡ��һ�ݱ���", vbInformation, gSysInfo.AppName
9                 Exit Sub
10            End If
11            lngSampleID = Val(.TextMatrix(.Row, .ColIndex("ID")))
12            lngVer = Val(.TextMatrix(.Row, .ColIndex("�汾")))
13        End With

14        Call funShowClincHelp(Me, lngSampleID, lngVer)

15        Exit Sub
ShowClincHelp_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(ShowClincHelp)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear
End Sub












