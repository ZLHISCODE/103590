VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.12#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareSendCard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ѿ�����"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9840
   Icon            =   "frmSquareSendCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1005
      Index           =   4
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9735
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   5190
      Width           =   9765
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   903
         Left            =   4665
         PasswordChar    =   "*"
         TabIndex        =   32
         Top             =   120
         Width           =   1725
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   904
         Left            =   1140
         TabIndex        =   39
         Top             =   540
         Width           =   2265
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   906
         Left            =   7920
         PasswordChar    =   "*"
         TabIndex        =   43
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox txt 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   905
         Left            =   4665
         PasswordChar    =   "*"
         TabIndex        =   41
         Top             =   540
         Width           =   1725
      End
      Begin zlIDKind.IDKindNew IDKind���� 
         Height          =   300
         Left            =   1140
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   120
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   529
         Appearance      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
         IDKind          =   -1
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   902
         Left            =   1710
         TabIndex        =   35
         Top             =   120
         Width           =   1695
      End
      Begin VB.ComboBox cboԭ������ 
         Height          =   300
         Left            =   4665
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   901
         Left            =   1140
         TabIndex        =   30
         Top             =   120
         Width           =   2265
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�¿���(&C)"
         Height          =   180
         Index           =   905
         Left            =   300
         TabIndex        =   38
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "����(&W)"
         Height          =   180
         Index           =   906
         Left            =   4020
         TabIndex        =   40
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ȷ������(&V)"
         Height          =   180
         Index           =   907
         Left            =   6870
         TabIndex        =   42
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ԭ����(&O)"
         Height          =   180
         Index           =   901
         Left            =   300
         TabIndex        =   29
         Top             =   180
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����(&P)"
         Height          =   180
         Index           =   902
         Left            =   480
         TabIndex        =   33
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ԭ������(&P)"
         Height          =   180
         Index           =   903
         Left            =   3660
         TabIndex        =   31
         Top             =   180
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ԭ����(&O)"
         Height          =   180
         Index           =   904
         Left            =   3840
         TabIndex        =   36
         Top             =   180
         Width           =   810
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Index           =   4
      Left            =   0
      TabIndex        =   60
      Top             =   6090
      Width           =   9825
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   2
         Left            =   2670
         ScaleHeight     =   1425
         ScaleWidth      =   7095
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   60
         Width           =   7125
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   403
            Left            =   720
            TabIndex        =   48
            Top             =   578
            Width           =   2445
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   405
            Left            =   720
            TabIndex        =   50
            Tag             =   "1"
            Top             =   1020
            Width           =   2445
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   406
            Left            =   4290
            TabIndex        =   51
            Tag             =   "1"
            Top             =   1020
            Width           =   2760
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   404
            Left            =   4290
            TabIndex        =   49
            Tag             =   "1"
            Top             =   600
            Width           =   2760
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   401
            Left            =   1770
            TabIndex        =   46
            Top             =   150
            Width           =   1410
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   402
            Left            =   4290
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   150
            Width           =   2760
         End
         Begin VB.ComboBox cbo֧����ʽ 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   150
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "�ɿ���"
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
            Index           =   404
            Left            =   60
            TabIndex        =   100
            Top             =   630
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��"
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
            Index           =   406
            Left            =   60
            TabIndex        =   67
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������"
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
            Index           =   407
            Left            =   3420
            TabIndex        =   66
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Height          =   210
            Index           =   405
            Left            =   3420
            TabIndex        =   65
            Top             =   630
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   402
            Left            =   60
            TabIndex        =   64
            Top             =   195
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "�� ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   403
            Left            =   3690
            TabIndex        =   63
            Top             =   195
            Width           =   570
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   1
         Left            =   30
         ScaleHeight     =   1425
         ScaleWidth      =   2535
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   60
         Width           =   2565
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "����"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   360
            Index           =   401
            Left            =   1665
            TabIndex        =   62
            Top             =   720
            Width           =   780
         End
         Begin XtremeSuiteControls.ShortcutCaption lbl�ɿ� 
            Height          =   375
            Left            =   10
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   10
            Width           =   2535
            _Version        =   589884
            _ExtentX        =   4471
            _ExtentY        =   661
            _StockProps     =   6
            Caption         =   "�տ�ϼ�"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   11.99
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Index           =   1
      Left            =   0
      TabIndex        =   61
      Top             =   7590
      Width           =   10095
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   210
         TabIndex        =   54
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7290
         TabIndex        =   52
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8460
         TabIndex        =   53
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "������0.00"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   5010
         TabIndex        =   101
         Top             =   300
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6045
      Index           =   3
      Left            =   30
      ScaleHeight     =   6015
      ScaleWidth      =   9735
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   30
      Width           =   9765
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Index           =   5
         Left            =   7440
         TabIndex        =   95
         Top             =   1290
         Width           =   2265
         Begin VB.PictureBox pic 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   495
            Index           =   0
            Left            =   120
            Picture         =   "frmSquareSendCard.frx":000C
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   57
            Top             =   3390
            Width           =   495
         End
         Begin VB.Frame fra 
            Caption         =   "    ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1185
            Index           =   52
            Left            =   0
            TabIndex        =   97
            Top             =   3525
            Width           =   2250
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "0.00"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   14.25
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Index           =   601
               Left            =   1470
               TabIndex        =   98
               Top             =   540
               Width           =   660
            End
         End
         Begin VB.Frame fra 
            Caption         =   "�������"
            Height          =   3330
            Index           =   51
            Left            =   0
            TabIndex        =   96
            Top             =   0
            Width           =   2250
            Begin MSComctlLib.ListView lvw������� 
               Height          =   3015
               Left            =   60
               TabIndex        =   44
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   5318
               View            =   3
               Arrange         =   1
               LabelEdit       =   1
               LabelWrap       =   0   'False
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   1
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Key             =   "����"
                  Object.Tag             =   "����"
                  Text            =   "����"
                  Object.Width           =   2540
               EndProperty
            End
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1245
         Index           =   2
         Left            =   10
         TabIndex        =   68
         Top             =   10
         Width           =   9705
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   201
            Left            =   3165
            TabIndex        =   1
            Top             =   90
            Width           =   2055
         End
         Begin VB.ComboBox cbo������ 
            Height          =   300
            Left            =   885
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   90
            Width           =   1455
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   202
            Left            =   5490
            TabIndex        =   2
            Top             =   90
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdd 
            Height          =   300
            Left            =   7650
            Picture         =   "frmSquareSendCard.frx":198E
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   90
            Width           =   300
         End
         Begin VB.CommandButton cmdDelete 
            Enabled         =   0   'False
            Height          =   300
            Left            =   8025
            Picture         =   "frmSquareSendCard.frx":81E0
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   90
            Width           =   300
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfCardNo 
            Height          =   735
            Left            =   30
            TabIndex        =   5
            Top             =   450
            Width           =   9675
            _cx             =   17066
            _cy             =   1296
            Appearance      =   2
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483643
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
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
         Begin VB.Line lineSplit 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   9735
            Y1              =   1230
            Y2              =   1230
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����(&N)"
            Height          =   180
            Index           =   202
            Left            =   2490
            TabIndex        =   73
            Top             =   150
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������(&T)"
            Height          =   180
            Index           =   201
            Left            =   60
            TabIndex        =   72
            Top             =   150
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   203
            Left            =   5280
            TabIndex        =   71
            Top             =   150
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "300"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Index           =   205
            Left            =   8880
            TabIndex        =   69
            Top             =   120
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����    ��"
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
            Index           =   204
            Left            =   8430
            TabIndex        =   70
            Top             =   135
            Width           =   1050
         End
      End
      Begin VB.Frame fra 
         Caption         =   "����ֵ"
         Height          =   705
         Index           =   7
         Left            =   60
         TabIndex        =   92
         Top             =   4080
         Width           =   7275
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   702
            Left            =   4965
            TabIndex        =   24
            Top             =   270
            Width           =   2265
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   701
            Left            =   1080
            TabIndex        =   23
            Top             =   270
            Width           =   2265
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ʵ�����۶�(&J)"
            Height          =   180
            Index           =   702
            Left            =   3720
            TabIndex        =   94
            Top             =   330
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�����(&M)"
            Height          =   180
            Index           =   701
            Left            =   240
            TabIndex        =   93
            Top             =   330
            Width           =   810
         End
      End
      Begin VB.Frame fra 
         Caption         =   "��ֵ��Ϣ"
         Height          =   1170
         Index           =   8
         Left            =   60
         TabIndex        =   86
         Top             =   4830
         Width           =   7275
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Index           =   803
            Left            =   6090
            TabIndex        =   27
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Index           =   802
            Left            =   3510
            TabIndex        =   26
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Index           =   801
            Left            =   1080
            TabIndex        =   25
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   315
            Index           =   804
            Left            =   1080
            TabIndex        =   28
            Top             =   720
            Width           =   6090
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���γ�ֵ(&B)"
            Height          =   180
            Index           =   803
            Left            =   2490
            TabIndex        =   91
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ֵ����(&K)"
            Height          =   180
            Index           =   801
            Left            =   45
            TabIndex        =   90
            Top             =   360
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʵ�ʳ�ֵ�ɿ�(&I)"
            Height          =   180
            Index           =   804
            Left            =   4710
            TabIndex        =   89
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   802
            Left            =   2220
            TabIndex        =   88
            Top             =   345
            Width           =   120
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "��ֵ˵��(&Z)"
            Height          =   180
            Index           =   805
            Left            =   45
            TabIndex        =   87
            Top             =   795
            Width           =   990
         End
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Caption         =   "����������Ϣ"
         Height          =   2835
         Index           =   3
         Left            =   10
         TabIndex        =   74
         Top             =   1320
         Width           =   7425
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   300
            Left            =   1095
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1245
            Width           =   570
            _ExtentX        =   1005
            _ExtentY        =   529
            Appearance      =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   9
            FontName        =   "����"
            IDKind          =   -1
            BackColor       =   -2147483633
         End
         Begin VB.TextBox txt 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   301
            Left            =   1095
            PasswordChar    =   "*"
            TabIndex        =   6
            Top             =   60
            Width           =   2295
         End
         Begin VB.CheckBox chk��ֵ 
            Caption         =   "�����ֵ"
            Height          =   180
            Left            =   1095
            TabIndex        =   8
            Top             =   510
            Width           =   1200
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   311
            Left            =   5040
            TabIndex        =   22
            Top             =   2460
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   310
            Left            =   1095
            TabIndex        =   21
            Top             =   2460
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   308
            Left            =   1095
            TabIndex        =   19
            Top             =   2055
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   307
            Left            =   1095
            TabIndex        =   18
            Top             =   1650
            Width           =   6225
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   270
            Index           =   2
            Left            =   7035
            TabIndex        =   17
            TabStop         =   0   'False
            Tag             =   "�쿨����"
            Top             =   1230
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   270
            Index           =   0
            Left            =   7035
            TabIndex        =   12
            TabStop         =   0   'False
            Tag             =   "����ԭ��"
            Top             =   825
            Width           =   285
         End
         Begin VB.CommandButton cmdSel 
            Caption         =   "��"
            Height          =   270
            Index           =   1
            Left            =   3090
            TabIndex        =   15
            TabStop         =   0   'False
            Tag             =   "�쿨��"
            Top             =   1260
            Width           =   285
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   304
            Left            =   1095
            MaxLength       =   50
            TabIndex        =   11
            Top             =   825
            Width           =   6225
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   306
            Left            =   5040
            TabIndex        =   16
            Top             =   1230
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   309
            Left            =   5040
            TabIndex        =   20
            Top             =   2055
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dtp����Ч���� 
            Height          =   300
            Left            =   5040
            TabIndex        =   10
            Top             =   450
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   529
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   186318851
            CurrentDate     =   40156.0854282407
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            Height          =   300
            Index           =   303
            Left            =   5040
            TabIndex        =   9
            Top             =   450
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox txt 
            Height          =   300
            Index           =   305
            Left            =   1650
            TabIndex        =   14
            Top             =   1245
            Width           =   1725
         End
         Begin VB.TextBox txt 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   302
            Left            =   5040
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ȷ������(&E)"
            Height          =   180
            Index           =   302
            Left            =   3990
            TabIndex        =   85
            Top             =   120
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "����(&W)"
            Height          =   180
            Index           =   301
            Left            =   450
            TabIndex        =   84
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   311
            Left            =   4230
            TabIndex        =   83
            Top             =   2520
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   310
            Left            =   540
            TabIndex        =   82
            Top             =   2520
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�쿨����(&M)"
            Height          =   180
            Index           =   306
            Left            =   3990
            TabIndex        =   81
            Top             =   1290
            Width           =   990
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��������"
            Height          =   180
            Index           =   309
            Left            =   4260
            TabIndex        =   80
            Top             =   2115
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Index           =   308
            Left            =   540
            TabIndex        =   79
            Top             =   2115
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��ע(&S)"
            Height          =   180
            Index           =   307
            Left            =   450
            TabIndex        =   78
            Top             =   1710
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����Ч����(&D)"
            Height          =   180
            Index           =   303
            Left            =   3810
            TabIndex        =   77
            Top             =   510
            Width           =   1170
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�쿨��(&D)"
            Height          =   180
            Index           =   305
            Left            =   270
            TabIndex        =   76
            Top             =   1305
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����ԭ��(&Y)"
            Height          =   180
            Index           =   304
            Left            =   90
            TabIndex        =   75
            Top             =   870
            Width           =   990
         End
      End
   End
End
Attribute VB_Name = "frmSquareSendCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ���
Private mfrmMain As Form, mlngModule As Long, mstrPrivs As String
Private mEditType As gCardEditType
Private mlng����� As Long, mlng��ID As Long
Private mlng��ֵID As Long

'�ؼ�����ö��ֵ
Private Enum mFrameIndex
    fra_���� = 2
    fra_����Ϣ = 3
    fra_����ֵ = 7
    fra_��ֵ��Ϣ = 8
    fra_������� = 51
    fra_������ = 5
    fra_��� = 52
    fra_�ɿ���� = 4
    fra_��ť��� = 1
End Enum
Private Enum mLableIndex
    lbl_������ = 201
    lbl_���� = 202
    lbl_������ = 203
    lbl_������ = 204
    lbl_������2 = 205
    lbl_���� = 301
    lbl_ȷ������ = 302
    lbl_����Ч�� = 303
    lbl_����ԭ�� = 304
    lbl_�쿨�� = 305
    lbl_�쿨���� = 306
    lbl_��ע = 307
    lbl_������ = 308
    lbl_�������� = 309
    lbl_������ = 310
    lbl_�������� = 311
    lbl_����� = 701
    lbl_���۶� = 702
    
    lbl_�ɿ�ϼ� = 401
    lbl_֧����ʽ = 402
    lbl_�Ҳ� = 403
    
    lbl_��� = 0
    
    lbl_ԭ������ = 901
    lbl_�������� = 902
    lbl_ԭ������ = 903
    lbl_����ԭ�� = 904
    lbl_�¿����� = 905
    lbl_�¿����� = 906
    lbl_�¿�����ȷ�� = 907
    lbl_����� = 601
End Enum
Private Enum mTextIndex
    txt_��ʼ���� = 201
    txt_�������� = 202
    txt_���� = 301
    txt_ȷ������ = 302
    txt_����Ч�� = 303
    txt_����ԭ�� = 304
    txt_�쿨�� = 305
    txt_�쿨���� = 306
    txt_��ע = 307
    txt_������ = 308
    txt_�������� = 309
    txt_������ = 310
    txt_�������� = 311
    txt_����� = 701
    txt_���۶� = 702
    
    txt_��ֵ���� = 801
    txt_���γ�ֵ = 802
    txt_��ֵ�ɿ� = 803
    txt_��ֵ˵�� = 804
    
    txt_�ɿ� = 401
    txt_�Ҳ� = 402
    txt_�ɿ��� = 403
    txt_������ = 404
    txt_�ʺ� = 405
    txt_������� = 406
    
    txt_ԭ������ = 901
    txt_�������� = 902
    txt_ԭ������ = 903
    txt_�¿����� = 904
    txt_�¿����� = 905
    txt_�¿�ȷ������ = 906
End Enum
Private Enum mCommandButtonIndex
    cmd_����ԭ�� = 0
    cmd_�쿨�� = 1
    cmd_�쿨���� = 2
End Enum
Private Enum mPictureIndex
    pic_��� = 0
    pic_�ɿ�ϼ� = 1
    pic_�ɿ���Ϣ = 2
    pic_����Ϣ = 3
    pic_������ = 4
End Enum

'ģ�����
Private mblnFirst As Boolean, mintSucces As Integer
Private mblnNotClick As Boolean, mblnChange As Boolean

Private mobjKeyboard As Object
Attribute mobjKeyboard.VB_VarHelpID = -1
Private Type Ty_Para
    bln�ɿ��ӡ As Boolean
    bln������ֵ As Boolean
End Type
Private mTy_MoudlePara As Ty_Para

Private Type Ty_CardType
    str������ As String
    str����ǰ׺ As String
    lng���ų��� As Long
    bln�������� As Boolean
    int���볤�� As Integer
    int���볤������ As Integer
    byt������� As Byte
    bln�ϸ���� As Boolean
    str������� As String
    bln�ض����� As Boolean
    lng�������� As Long
    lng����ID As Long
End Type
Private mCardType As Ty_CardType
Private mrs������ As ADODB.Recordset
Private mobjCard As clsSquareCard '��ǰ����Ϣ
Private mcllCard As Collection '��Ƭ��Ϣ���ϣ�����ʱʹ��
Private mdblʵ�պϼ� As Double '�տ�ϼ�
Private mdbl������� As Double

'֧�����
Private mobjPayCards As Cards
Private mlngPre֧����ʽ As Long
Private Type TY_PayMoney
    str���㷽ʽ  As String
    lng�����ID As Long
    byt�������� As Byte
    strˢ������ As String
    strˢ������ As String
    str������ˮ�� As String
    str����˵�� As String
    lngԭ������� As Long
End Type
Private mCurCardPay As TY_PayMoney '���ο�֧��
Private mrsBalance As ADODB.Recordset 'ԭ�������
Private mBytMoney As Byte '�ֱҴ������

Public Function zlShowCard(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As gCardEditType, ByVal lng����� As Long, Optional ByVal lng��ID As Long, _
    Optional lng��ֵID As Long) As Boolean
    '�������,�鿴�ѷ��������ӷ������޸ķ�����Ϣ
    '��Σ�
    '   frmMain - ������
    '   lngModule - ģ���
    '   strPrivs - Ȩ�޴�
    '   EditType - ��������
    '   lng����� As Long - ���ѿ����
    '   lng��ID - ��ǰ�������ѿ�ID
    '   lng��ֵID - ��ֵ����ʱ���룬��ֵ��¼��ID
    '���أ������ɹ�����True,���򷵻�False
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs:
    mlng����� = lng�����: mEditType = EditType
    mlng��ID = lng��ID
    mlng��ֵID = lng��ֵID
    
    mintSucces = 0
    On Error Resume Next
    Me.Show 1, frmMain
    zlShowCard = mintSucces > 0
End Function

Private Function CardIsValid(ByVal lng��ID As Long) As Boolean
    '��鿨��Ϣ
    Dim strInfo As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str���� As String
    Dim msgBoxStyle As VbMsgBoxStyle
    
    On Error GoTo ErrHandler
    If lng��ID = 0 Then CardIsValid = True: Exit Function
    If mEditType = gEd_���� Then CardIsValid = True: Exit Function
    
    strSQL = _
        "Select ID, ������, �ɷ��ֵ, ����, ���, " & vbNewLine & _
        "       (Select Max(���) From ���ѿ���Ϣ B Where a.���� = b.���� And a.�ӿڱ�� = b.�ӿڱ��) As ������," & vbNewLine & _
        "       To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, To_Char(��Ч��, 'yyyy-mm-dd hh24:mi:ss') As ��Ч��, " & vbNewLine & _
        "       To_Char(ͣ������, 'yyyy-mm-dd hh24:mi:ss') As ͣ������, ��ǰ״̬, ������, ���" & vbNewLine & _
        "From ���ѿ���Ϣ A" & vbNewLine & _
        "Where a.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID)
    
    If rsTemp.EOF Then
        strInfo = Switch(mEditType = gEd_�޸�, "�޸Ŀ���Ϣ", mEditType = gEd_��ֵ, "��ֵ", _
                         mEditType = gEd_�˿�, "�˿�", mEditType = gEd_ȡ���˿�, "ȡ���˿�", _
                         mEditType = gEd_����, "����", mEditType = gEd_����, "����", _
                         mEditType = gEd_��ֵ����, "��ֵ����", _
                         mEditType = gEd_����, "����", mEditType = gEd_ȡ������, "ȡ������", _
                         True, "��������")
        ShowMsgbox mCardType.str������ & "�����Ѿ�������ɾ��������" & strInfo & "��"
        Exit Function
    End If
    str���� = NVL(rsTemp!����)
    
    '��鿨���Ƿ�Ϸ�
    Select Case mEditType
    Case gEd_�޸�
        If Val(NVL(rsTemp!���)) < Val(NVL(rsTemp!������)) Then
            ShowMsgbox "�����޸���ʷ������Ϣ(����Ϊ:" & str���� & ") ��"
            Exit Function
        End If
    Case gEd_��ֵ, gEd_��ֵ����, gEd_�˿�, gEd_ȡ���˿�, gEd_����, gEd_ȡ������, gEd_����, gEd_����
        strInfo = Switch(mEditType = gEd_��ֵ, "��ֵ", mEditType = gEd_��ֵ����, "��ֵ����", _
                         mEditType = gEd_�˿�, "�˿�", mEditType = gEd_ȡ������, "ȡ������", _
                         mEditType = gEd_����, "����", mEditType = gEd_����, "����", _
                         mEditType = gEd_����, "����", _
                         mEditType = gEd_ȡ���˿�, "ȡ���˿�", True, "��������")
        
        If Val(NVL(rsTemp!���)) < Val(NVL(rsTemp!������)) Then
            ShowMsgbox "���ܶ���ʷ���Ž���" & strInfo & "(����Ϊ:" & str���� & ")��"
            Exit Function
        End If
        If mEditType = gEd_ȡ������ Then
            '�������յĿ�����ȡ������
            If Val(NVL(rsTemp!��ǰ״̬)) = 4 Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "Ϊ�������յĿ�������ȡ�����գ�"
                Exit Function
            End If
            If NVL(rsTemp!����ʱ��, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "���ܱ�����ȡ������(����)��������ȡ�����գ�"
                Exit Function
            End If
        ElseIf mEditType = gEd_ȡ���˿� Then
            If NVL(rsTemp!����ʱ��, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "���ܱ�����ȡ���˿�(����)��������ȡ���˿���"
                Exit Function
            End If
        ElseIf mEditType = gEd_�˿� Then
            If NVL(rsTemp!����ʱ��, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�ѱ��˿����������˿���"
                Exit Function
            End If
        Else
            If NVL(rsTemp!����ʱ��, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�ѱ����ջ��˿���������" & strInfo & "��"
                Exit Function
            End If
        End If
        If Not (mEditType = gEd_���� Or mEditType = gEd_ȡ������) Then
            'ͣ�õ�Ҳ���Ի��պ�ȡ������
            If NVL(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�Ѿ�ֹͣʹ�ã�������" & strInfo & "��"
                Exit Function
            End If
        End If
        
        Select Case mEditType
        Case gEd_����
            If Val(NVL(rsTemp!���)) > 0 Then
                If NVL(rsTemp!��Ч��, "3000-01-01 00:00:00") > Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") Then
                    msgBoxStyle = vbQuestion + vbYesNo + vbDefaultButton2
                Else '��ʧЧ��Ĭ�ϻ���
                    msgBoxStyle = vbQuestion + vbYesNo + vbDefaultButton1
                End If
                If MsgBox("����Ϊ:" & str���� & " ��" & mCardType.str������ & "��ǰ�������(" & _
                    FormatEx(Val(NVL(rsTemp!���)), 6, , , 2) & ")����ȷ��Ҫ������", msgBoxStyle) = vbNo Then Exit Function
            End If
        Case gEd_��ֵ
            If Val(NVL(rsTemp!�ɷ��ֵ)) <> 1 Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "���ǳ�ֵ�������ܳ�ֵ��"
                Exit Function
            End If
        Case gEd_��ֵ����
            strSQL = "Select 1" & vbNewLine & _
                "From ���˿������¼ A, ���ѿ���Ϣ B" & vbNewLine & _
                "Where a.���ѿ�id = b.Id And a.Id = [1] And (Nvl(b.���, 0) - Nvl(a.Ӧ�ս��, 0)) >= 0"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ֵID)
            If rsTemp.EOF Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�����㣬���ܳ�ֵ���ˣ�"
                Exit Function
            End If
            
            'ֻ�������Ժ�ĲŽ��м�� And b.������� > 0
            strSQL = "Select 1" & vbNewLine & _
                "From ���˿������¼ A, �ʻ��ɿ���� B" & vbNewLine & _
                "Where a.������� = b.������� And a.���ѿ�id = b.���ѿ�id And (a.Ӧ�ս�� = b.��� Or b.������� <= 0) And a.Id = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ֵID)
            If rsTemp.EOF Then
                ShowMsgbox "�ñʳ�ֵҲ������ʹ�ã����ܳ�ֵ���ˣ�"
                Exit Function
            End If
        Case gEd_�˿�
            If NVL(rsTemp!������) <> UserInfo.���� Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�����㷢�ŵĿ��������˿���"
                Exit Function
            End If
            
            strSQL = "Select 1 From ���˿������¼ Where ���ѿ�id = [1] And ��¼���� = 4 And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID)
            If rsTemp.EOF = False Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�Ѿ��������ѣ��������˿�����ֻ�ܻ��տ�Ƭ��"
                Exit Function
            End If
            
            strSQL = "Select 1 From ���˿������¼��where ���ѿ�id = [1] And ��¼���� = 2 And ��¼״̬ = 1 Having Count(1) > 1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID)
            If rsTemp.EOF = False Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�Ѿ�����γ�ֵ���������˿���ֻ�ܻ��տ�Ƭ��"
                Exit Function
            End If
            
            strSQL = "Select 1 From ���˿������¼��where ���ѿ�id = [1] And ��¼���� = 3 And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID)
            If rsTemp.EOF = False Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�ѽ���������˿�������˿���ֻ�ܻ��տ�Ƭ��"
                Exit Function
            End If
        Case gEd_ȡ���˿�
            If NVL(rsTemp!������) <> UserInfo.���� Then
                ShowMsgbox "����Ϊ:" & str���� & " ��" & mCardType.str������ & "�����㷢�ŵĿ�������ȡ���˿���"
                Exit Function
            End If
        End Select
    End Select
    CardIsValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitFace() As Boolean
    '��ʼ������
    
    On Error GoTo ErrHandler
    Call SetCtlVisible
    Call FormResize
    
    pic(pic_����Ϣ).AutoRedraw = True: zlControl.PicShowFlat pic(pic_����Ϣ)
    pic(pic_�ɿ�ϼ�).AutoRedraw = True: zlControl.PicShowFlat pic(pic_�ɿ�ϼ�)
    pic(pic_�ɿ���Ϣ).AutoRedraw = True: zlControl.PicShowFlat pic(pic_�ɿ���Ϣ)
    pic(pic_������).AutoRedraw = True: zlControl.PicShowFlat pic(pic_������)
    cbo.SetListWidth cbo֧����ʽ, cbo֧����ʽ.Width * 2
    
    Call SetCtlEnable
    Call SetEnabledBackColor(Me)
    lvw�������.BackColor = IIf(lvw�������.Enabled, &H80000005, Me.BackColor)
    txt(txt_�Ҳ�).BackColor = Me.BackColor

    Me.Caption = Switch(mEditType = gEd_����, "����", mEditType = gEd_�޸�, "��Ϣ�޸�", _
                        mEditType = gEd_����, "����", mEditType = gEd_����, "����", _
                        mEditType = gEd_��ѯ, "��Ϣ��ѯ", _
                        mEditType = gEd_��ֵ, "��ֵ����", mEditType = gEd_��ֵ����, "��ֵ����", _
                        mEditType = gEd_����, "���չ���", mEditType = gEd_ȡ������, "ȡ������", _
                        mEditType = gEd_�˿�, "�˿�", mEditType = gEd_ȡ���˿�, "ȡ���˿�", _
                        True, "����") & " - " & mCardType.str������
    
    If mEditType = gEd_��ѯ Then
        cmdOK.Visible = False
        cmdCancel.Caption = cmdOK.Caption
    End If
    
    InitFace = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetCtlVisible()
    '���ÿؼ��Ŀɼ�״̬
    Dim blnVisible As Boolean
    
    On Error GoTo ErrHandler
    '�����������"�ض�����"��������������
    If mEditType <> gEd_���� Or mEditType = gEd_���� And mCardType.bln�ض����� Then
        lbl(lbl_������).Visible = False: txt(txt_��������).Visible = False
        If mEditType <> gEd_���� Then
            cmdAdd.Visible = False: cmdDelete.Visible = False
            lbl(lbl_������).Visible = False: lbl(lbl_������2).Visible = False
            vsfCardNo.Visible = False: lineSplit.Visible = False
        End If
        Call FrameResize(fra_����)
    End If
       
    blnVisible = (mEditType = gEd_����)
    lbl(lbl_����).Visible = blnVisible: txt(txt_����).Visible = blnVisible
    lbl(lbl_ȷ������).Visible = blnVisible: txt(txt_ȷ������).Visible = blnVisible
    
    blnVisible = (mEditType = gEd_���� Or mEditType = gEd_�޸�)
    dtp����Ч����.Visible = blnVisible: txt(txt_����Ч��).Visible = Not blnVisible
    cmdSel(cmd_����ԭ��).Visible = blnVisible
    IDKind.Visible = blnVisible And mCardType.bln�ض�����
    cmdSel(cmd_�쿨��).Visible = blnVisible And mCardType.bln�ض����� = False
    cmdSel(cmd_�쿨����).Visible = blnVisible And mCardType.bln�ض����� = False
    
    blnVisible = (mEditType <> gEd_����)
    lbl(lbl_������).Visible = blnVisible: txt(txt_������).Visible = blnVisible
    lbl(lbl_��������).Visible = blnVisible: txt(txt_��������).Visible = blnVisible
    lbl(lbl_������).Visible = blnVisible: txt(txt_������).Visible = blnVisible
    lbl(lbl_��������).Visible = blnVisible: txt(txt_��������).Visible = blnVisible
    Call FrameResize(fra_����Ϣ)
    
    blnVisible = (mEditType = gEd_���� Or mEditType = gEd_�˿� Or mEditType = gEd_ȡ���˿� Or mEditType = gEd_��ֵ Or mEditType = gEd_��ֵ����)
    fra(fra_��ֵ��Ϣ).Visible = blnVisible
    
    blnVisible = (mEditType = gEd_���� Or mEditType = gEd_�˿� Or mEditType = gEd_ȡ���˿� _
                    Or mEditType = gEd_��ֵ Or mEditType = gEd_��ֵ����)
    fra(fra_�ɿ����).Visible = blnVisible
    
    blnVisible = (mEditType = gEd_���� Or mEditType = gEd_����)
    pic(pic_������).Visible = blnVisible
    If blnVisible Then
        blnVisible = (mEditType = gEd_����)
        lbl(lbl_ԭ������).Visible = blnVisible
        txt(txt_ԭ������).Visible = blnVisible
        lbl(lbl_ԭ������).Visible = blnVisible
        txt(txt_ԭ������).Visible = blnVisible
        
        lbl(lbl_��������).Visible = Not blnVisible
        txt(txt_��������).Visible = Not blnVisible
        IDKind����.Visible = Not blnVisible
        lbl(lbl_����ԭ��).Visible = Not blnVisible
        cboԭ������.Visible = Not blnVisible
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetCtlEnable()
    '���ÿؼ��Ŀ���״̬
    Dim blnEnable As Boolean
    
    On Error GoTo ErrHandler
    blnEnable = (mEditType = gEd_���� Or mEditType = gEd_�޸�)
    cbo������.Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_���� Or mEditType = gEd_���� Or mEditType = gEd_��ֵ)
    txt(txt_��ʼ����).Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_���� Or mEditType = gEd_�޸�)
    chk��ֵ.Enabled = blnEnable And mobjCard.�ѳ�ֵ = False
    txt(txt_����ԭ��).Enabled = blnEnable
    txt(txt_�쿨��).Enabled = blnEnable
    txt(txt_�쿨����).Enabled = blnEnable And mCardType.bln�ض����� = False
    txt(txt_��ע).Enabled = blnEnable
    lvw�������.Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_����)
    txt(txt_�����).Enabled = blnEnable And zlStr.IsHavePrivs(mstrPrivs, "������Ŀ����")
    txt(txt_���۶�).Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_���� Or mEditType = gEd_��ֵ) And chk��ֵ.value = vbChecked
    txt(txt_��ֵ����).Enabled = blnEnable
    txt(txt_���γ�ֵ).Enabled = blnEnable
    txt(txt_��ֵ�ɿ�).Enabled = blnEnable
    txt(txt_��ֵ˵��).Enabled = blnEnable
    
    blnEnable = (mEditType = gEd_���� Or mEditType = gEd_��ֵ Or mEditType = gEd_ȡ���˿�)
    txt(txt_�ɿ���).Enabled = blnEnable
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FrameResize(ByVal Index As Integer)
    Dim sngTop As Single '��һ�пؼ���Topֵ
    Dim sngSplit As Single '�ؼ����м��
    Dim sngDiff As Single '��ǩ�ؼ���Top���ı���ؼ�Top�Ĳ��
    
    On Error Resume Next
    sngDiff = 60
    sngTop = IIf(mEditType = gEd_����, 100, 50)
    sngSplit = IIf(mEditType = gEd_����, 80, 160)
    Select Case Index
    Case fra_����
        If mEditType = gEd_���� And mCardType.bln�ض����� = False Then Exit Sub
        If mEditType = gEd_���� Then
            cmdAdd.Left = txt(txt_��ʼ����).Left + txt(txt_��ʼ����).Width + 100
            cmdDelete.Left = cmdAdd.Left + cmdAdd.Width + 50
            lbl(lbl_������).Left = cmdDelete.Left + cmdDelete.Width + 150
            Call SetLblCaption(lbl_������, True)
            sngTop = sngTop + txt(txt_��ʼ����).Height + 100
            
            vsfCardNo.Top = sngTop: vsfCardNo.Height = 1000
            sngTop = sngTop + vsfCardNo.Height + 100
            
            fra(fra_����).Height = sngTop
            lineSplit.Y1 = fra(fra_����).Height - 10: lineSplit.Y2 = lineSplit.Y1
        Else
            cbo������.Top = sngTop: cbo������.Left = txt(txt_����).Left
            cbo������.Width = txt(txt_����).Width
            lbl(lbl_������).Top = cbo������.Top + sngDiff: lbl(lbl_������).Left = cbo������.Left - lbl(lbl_������).Width - 20
            txt(txt_��ʼ����).Top = cbo������.Top: txt(txt_��ʼ����).Left = txt(txt_ȷ������).Left
            txt(txt_��ʼ����).Width = txt(txt_ȷ������).Width
            lbl(lbl_����).Top = lbl(lbl_������).Top: lbl(lbl_����).Left = txt(txt_��ʼ����).Left - lbl(lbl_����).Width - 60
            sngTop = sngTop + txt(txt_��ʼ����).Height
            
            fra(fra_����).Height = sngTop
            fra(fra_����).Width = fra(fra_����Ϣ).Width
        End If
    Case fra_����Ϣ
        If mEditType = gEd_���� Then
            txt(txt_����).Top = sngTop: lbl(lbl_����).Top = txt(txt_����).Top + sngDiff
            txt(txt_ȷ������).Top = txt(txt_����).Top: lbl(lbl_ȷ������).Top = lbl(lbl_����).Top
            sngTop = sngTop + txt(txt_����).Height + sngSplit
        End If
        
        txt(txt_����Ч��).Top = sngTop: lbl(lbl_����Ч��).Top = sngTop + sngDiff
        dtp����Ч����.Top = txt(txt_����Ч��).Top
        chk��ֵ.Top = lbl(lbl_����Ч��).Top
        sngTop = sngTop + txt(txt_����Ч��).Height + sngSplit
        
        txt(txt_����ԭ��).Top = sngTop: lbl(lbl_����ԭ��).Top = sngTop + sngDiff
        cmdSel(cmd_����ԭ��).Top = txt(txt_����ԭ��).Top
        sngTop = sngTop + txt(txt_����ԭ��).Height + sngSplit
        
        txt(txt_�쿨��).Top = sngTop: lbl(lbl_�쿨��).Top = sngTop + sngDiff
        IDKind.Top = txt(txt_�쿨��).Top: cmdSel(cmd_�쿨��).Top = txt(txt_�쿨��).Top
        txt(txt_�쿨����).Top = sngTop: lbl(lbl_�쿨����).Top = sngTop + sngDiff
        cmdSel(cmd_�쿨����).Top = txt(txt_�쿨����).Top
        sngTop = sngTop + txt(txt_�쿨��).Height + sngSplit
        If Not ((mEditType = gEd_���� Or mEditType = gEd_�޸�) And mCardType.bln�ض�����) Then
            txt(txt_�쿨��).Left = txt(txt_������).Left: txt(txt_�쿨��).Width = txt(txt_������).Width
        End If
        
        txt(txt_��ע).Top = sngTop: lbl(lbl_��ע).Top = sngTop + sngDiff
        sngTop = sngTop + txt(txt_�쿨��).Height + sngSplit
        
        If mEditType <> gEd_���� Then
            txt(txt_������).Top = sngTop: lbl(lbl_������).Top = txt(txt_������).Top + sngDiff
            txt(txt_��������).Top = txt(txt_������).Top: lbl(lbl_��������).Top = lbl(lbl_������).Top
            sngTop = sngTop + txt(txt_������).Height + sngSplit
            
            txt(txt_������).Top = sngTop: lbl(lbl_������).Top = txt(txt_������).Top + sngDiff
            txt(txt_��������).Top = txt(txt_������).Top: lbl(lbl_��������).Top = lbl(lbl_������).Top
            sngTop = sngTop + txt(txt_������).Height + sngSplit
        End If
        
        fra(Index).Height = sngTop
    Case fra_������
        fra(fra_���).Top = fra(Index).Height - fra(fra_���).Height + 10
        pic(pic_���).Top = fra(fra_���).Top - pic(pic_���).Height / 2 + 100
        fra(fra_�������).Top = 50
        fra(fra_�������).Height = pic(pic_���).Top - fra(fra_�������).Top
        lvw�������.Height = fra(fra_�������).Height - lvw�������.Top - sngSplit
    End Select
End Sub

Private Sub FormResize()
    Dim sngTop As Single '��һ�пؼ���Topֵ
    Dim sngSplit As Single '�ؼ����м��
    
    On Error Resume Next
    sngTop = 10: sngSplit = 80
    fra(fra_����).Top = sngTop
    sngTop = sngTop + fra(fra_����).Height + sngSplit
    
    fra(fra_����Ϣ).Top = sngTop
    sngTop = sngTop + fra(fra_����Ϣ).Height
    
    fra(fra_����ֵ).Top = sngTop
    sngTop = sngTop + fra(fra_����ֵ).Height + sngSplit
    
    If mEditType = gEd_���� Or mEditType = gEd_�˿� Or mEditType = gEd_ȡ���˿� Or mEditType = gEd_��ֵ Or mEditType = gEd_��ֵ���� Then
        fra(fra_��ֵ��Ϣ).Top = sngTop
        sngTop = sngTop + fra(fra_��ֵ��Ϣ).Height + sngSplit
    End If
    
    If mEditType = gEd_���� And mCardType.bln�ض����� = False Or mEditType = gEd_���� Then
        fra(fra_������).Top = fra(fra_����Ϣ).Top
    Else
        fra(fra_������).Top = fra(fra_����).Top
    End If
    fra(fra_������).Height = sngTop - fra(fra_������).Top - sngSplit
    Call FrameResize(fra_������)
    
    '��������Ϣ
    pic(pic_����Ϣ).Top = sngSplit
    pic(pic_����Ϣ).Height = sngTop
    sngTop = pic(pic_����Ϣ).Top + pic(pic_����Ϣ).Height + sngSplit
    
    '�����������
    If mEditType = gEd_���� Or mEditType = gEd_���� Then
        pic(pic_������).Top = sngTop
        sngTop = sngTop + pic(pic_������).Height + sngSplit
    End If
    
    '�ɿ����
    If mEditType = gEd_���� Or mEditType = gEd_�˿� Or mEditType = gEd_ȡ���˿� Or mEditType = gEd_��ֵ Or mEditType = gEd_��ֵ���� Then
        fra(fra_�ɿ����).Top = sngTop
        sngTop = sngTop + fra(fra_�ɿ����).Height + sngSplit
    End If
    
    sngTop = sngTop - sngSplit
    fra(fra_��ť���).Top = sngTop
    sngTop = sngTop + fra(fra_��ť���).Height
    
    Me.Height = sngTop + 480
End Sub

Private Sub cboԭ������_Click()
    On Error GoTo ErrHandler
    If Val(cboԭ������.Tag) = cboԭ������.ItemData(cboԭ������.ListIndex) Then Exit Sub
    cboԭ������.Tag = cboԭ������.ItemData(cboԭ������.ListIndex)
    
    LoadCardData 1, cboԭ������.Text
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboԭ������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo֧����ʽ_Click()
    Dim objCard As Card
    Dim ty_Temp As TY_PayMoney
    Dim intSelectIndex As Integer
    
    On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) Then Exit Sub
    
    If (mEditType = gEd_��ֵ���� Or mEditType = gEd_�˿�) And mlngPre֧����ʽ > 0 Then
        '�������ԭ�ɿ���㷽ʽ�оͲ��ü�飬��Ҫ���֧�֡�ת�ʼ����ۡ���
        Set objCard = mobjPayCards(mlngPre֧����ʽ)
        mrsBalance.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
        
        If Not mrsBalance.EOF Then
            mblnNotClick = True
            intSelectIndex = cbo֧����ʽ.ListIndex
            cbo֧����ʽ.ListIndex = cbo.FindIndex(cbo֧����ʽ, mlngPre֧����ʽ)
            If CheckThreeBalanceToCash(objCard) = False Then mblnNotClick = False: Exit Sub
            cbo֧����ʽ.ListIndex = intSelectIndex
            mblnNotClick = False
        End If
    End If
    
    mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    
    mCurCardPay = ty_Temp '�Զ���Type��ʼ��
    With mCurCardPay
        .str���㷽ʽ = objCard.���㷽ʽ
        .lng�����ID = IIf(objCard.�ӿ���� > 0, objCard.�ӿ����, 0)
        .byt�������� = objCard.��������
    End With
    
    txt(txt_�ɿ�).Text = ""
    Call SetControlProperty
    Exit Sub
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlProperty(Optional ByVal blnLoadDefault As Boolean)
    '���ÿؼ�����
    '���:
    '   blnLoadDefault-�Ƿ����ȱʡֵ
    Dim blnDel As Boolean, objCard As Card
    Dim blnEnabled As Boolean
    Dim dblTemp As Double, dblErrMoney As Double
    
    On Error GoTo ErrHandler
    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    
    mdbl������� = 0
    If mEditType = gEd_���� Or mEditType = gEd_��ֵ Then
        If objCard.�������� = 1 Then
            txt(txt_���۶�).Text = Format(CentMoney(Val(txt(txt_���۶�).Text), mBytMoney), "#0.00;-#0.00")
            txt(txt_��ֵ�ɿ�).Text = Format(CentMoney(Val(txt(txt_��ֵ�ɿ�).Text), mBytMoney), "#0.00;-#0.00")
            If Val(txt(txt_��ֵ�ɿ�).Text) > Val(txt(txt_���γ�ֵ).Text) Then
                txt(txt_���γ�ֵ).Text = Format(Val(txt(txt_��ֵ�ɿ�).Text), "#0.00;-#0.00")
            End If
            If Val(txt(txt_���γ�ֵ).Text) <> 0 Then
                txt(txt_��ֵ����).Text = Format((Round(Val(txt(txt_��ֵ�ɿ�).Text) / Val(txt(txt_���γ�ֵ).Text), 6)) * 100, "0.00")
            End If
            Call Calcʵ�պϼ�(True)
        End If
    End If
    
    blnDel = (mdblʵ�պϼ� < 0)
    If blnDel Then
        lbl�ɿ�.Caption = "�˿�ϼ�"
        lbl(lbl_֧����ʽ).Caption = "�� ��"
        lbl(lbl_�ɿ�ϼ�).ForeColor = vbRed
        lbl(lbl_֧����ʽ).ForeColor = vbRed
    Else
        lbl�ɿ�.Caption = "�տ�ϼ�"
        lbl(lbl_֧����ʽ).Caption = "�� ��"
        lbl(lbl_�ɿ�ϼ�).ForeColor = vbBlue
        lbl(lbl_֧����ʽ).ForeColor = vbBlack
    End If
    
    '֧Ʊ��һ��ͨ���ϰ�һ��ͨ��������ɿλ
    '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
    blnEnabled = InStr(",2,7,8,", "," & objCard.�������� & ",") > 0
    txt(txt_������).Enabled = objCard.�������� <> 1
    txt(txt_�ʺ�).Enabled = objCard.�������� <> 1
    txt(txt_�������).Enabled = objCard.�������� <> 1
    If objCard.�������� = 1 Then
        txt(txt_������).Text = ""
        txt(txt_�ʺ�).Text = ""
        txt(txt_�������).Text = ""
        mdbl������� = mdblʵ�պϼ� - CentMoney(mdblʵ�պϼ�, mBytMoney)
    End If
    Call zl_SetCtlBackColor(Array(txt(txt_������), txt(txt_�ʺ�), txt(txt_�������)), Me)
    
    lbl(lbl_�ɿ�ϼ�).Caption = Format(Abs(mdblʵ�պϼ� - mdbl�������), "0.00")
                
    'ȱʡ��������
    txt(txt_�ɿ�).Locked = False
    If objCard.�ӿ���� > 0 Then '��������
        txt(txt_�ɿ�).Text = Format(Abs(mdblʵ�պϼ�), "0.00")
        txt(txt_�ɿ�).Locked = True
    ElseIf objCard.�������� = 1 Then '�ֽ���
        txt(txt_�ɿ�).Text = IIf(blnDel, Format(Abs(mdblʵ�պϼ� - mdbl�������), "0.00"), "")
    Else
        txt(txt_�ɿ�).Text = Format(Abs(mdblʵ�պϼ�), "0.00")
        txt(txt_�ɿ�).Locked = True
    End If
    lbl(lbl_���).Caption = FormatEx(IIf(blnDel, -1, 1) * mdbl�������, 6, , , 2)
    lbl(lbl_���).Visible = Val(lbl(lbl_���).Caption) <> 0
    lbl(lbl_���).Caption = "������" & lbl(lbl_���).Caption
    
    '�����Ҳ�
    Call SetLblCaption(lbl_�Ҳ�)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk��ֵ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdAdd_Click()
    Dim strCardNoRange As String
    Dim lng������ As Long, strCardNos As String, lng��ID As Long
    Dim objListItem As ListItem
    
    On Error GoTo ErrHandler
    If CheckInput����(False, lng������, strCardNos, lng��ID) = False Then Exit Sub
    
    strCardNoRange = Trim(txt(txt_��ʼ����).Text)
    If Trim(txt(txt_��������).Text) <> "" Then
        strCardNoRange = strCardNoRange & "��" & Trim(txt(txt_��������).Text)
    End If
    If mEditType = gEd_���� Then strCardNos = strCardNoRange
    If ZL_vsGrid_AddCell(vsfCardNo, strCardNoRange, Array(lng������, strCardNos, lng��ID)) = False Then Exit Sub
    If mEditType = gEd_���� Then
        mcllCard.Add mobjCard, "K" & mlng��ID
        Call FindDataInGrid(strCardNos, True)
    End If
    With vsfCardNo
        .Redraw = flexRDNone
        Call ZL_vsGrid_AutoSetGridRowAndCol(vsfCardNo)
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftCenter
        .Redraw = flexRDBuffered
    End With
    
    txt(txt_��ʼ����).Text = "": txt(txt_��������).Text = ""
    
    '��ʾ���㵱ǰ������
    Call SetLblCaption(lbl_������, mEditType = gEd_����)
    Call Calcʵ�պϼ�
    
    zlControl.ControlSetFocus txt(txt_��ʼ����)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Dim blnYes As Boolean
    
    On Error GoTo ErrHandler
    If mblnChange Then
        If mEditType = gEd_���� Or mEditType = gEd_��ֵ And mTy_MoudlePara.bln������ֵ Then
            ShowMsgbox "ȷʵҪ�����ǰ��¼���������", True, blnYes
            If blnYes = False Then Exit Sub
            Call ClearCtlData
            mlng��ID = 0
            Call zlControl.ControlSetFocus(txt(txt_��ʼ����))
            Exit Sub
        End If
    End If
    
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    Dim lng������ As Long
    Dim strCardNoRange As String, strCardNos As String
    Dim blnYes As Boolean
    Dim varData As Variant
    
    On Error GoTo ErrHandler
    If ZL_vsGrid_CurrCellHaveData(vsfCardNo) = False Then Exit Sub
    
    strCardNoRange = vsfCardNo.TextMatrix(vsfCardNo.Row, vsfCardNo.Col)
    Call ShowMsgbox("��ȷ��Ҫ��" & IIf(mEditType = gEd_����, "����", "����") & "�б����Ƴ� " & strCardNoRange & " ��", True, blnYes)
    If blnYes = False Then zlControl.ControlSetFocus cmdDelete: Exit Sub
    
    varData = vsfCardNo.Cell(flexcpData, vsfCardNo.Row, vsfCardNo.Col) 'Array(������,�ֽ⿨��,���ѿ�ID)
    lng������ = varData(0)
    
    If ZL_vsGrid_RemoveCell(vsfCardNo) = False Then Exit Sub
    If mEditType = gEd_���� And Not mcllCard Is Nothing Then
        If CollExitsValue(mcllCard, "K" & mlng��ID) Then
            mcllCard.Remove "K" & mlng��ID
        End If
    End If
    With vsfCardNo
        .Redraw = flexRDNone
        Call ZL_vsGrid_AutoSetGridRowAndCol(vsfCardNo)
        .Redraw = flexRDBuffered
    End With
    
    '��ʾ���㵱ǰ������
    Call SetLblCaption(lbl_������, mEditType = gEd_����)
    Call Calcʵ�պϼ�
    
    zlControl.ControlSetFocus txt(txt_��ʼ����)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    If IsValid() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mblnChange = False: mintSucces = mintSucces + 1
    
    If mEditType = gEd_���� Or mEditType = gEd_��ֵ And mTy_MoudlePara.bln������ֵ Then
        Call ClearCtlData
        mlng��ID = 0
        Call zlControl.ControlSetFocus(txt(txt_��ʼ����))
        Exit Sub
    End If
    
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdSel_Click(Index As Integer)
    Dim lngID As Long, str���� As String, str���� As String
    
    On Error GoTo ErrHandler
    Select Case Index
    Case cmd_�쿨��
        'ѡ����Ա
        lngID = Val(txt(txt_�쿨����).Tag)
        If Select��Աѡ����(Me, txt(txt_�쿨��), "", lngID, , True) = False Then Exit Sub
        
        If mEditType = gEd_���� Or mEditType = gEd_�޸� Then
            '�쿨�˾��ǽɿ���
            txt(txt_�ɿ���).Text = txt(txt_�쿨��).Text
            txt(txt_�ɿ���).Tag = txt(txt_�쿨��).Tag
        End If
        '��Ҫ��ȡȱʡ����:
        If zl_From��Ա��ȡȱʡ����(Val(txt(txt_�쿨��).Tag), str����, str����, lngID) Then
            txt(txt_�쿨����).Text = str���� & "-" & str����
            txt(txt_�쿨����).Tag = lngID
        End If
    Case cmd_�쿨����
        'ѡ��ȱʡ����
        lngID = Val(txt(txt_�쿨��).Tag)
        If Select����ѡ����(Me, txt(txt_�쿨����), "", "", IIf(lngID = 0, False, True), "", 0, _
            "����ѡ����", , , , , lngID) = False Then Exit Sub
    Case cmd_����ԭ��
        If zl_SelectAndNotAddItem(Me, txt(txt_����ԭ��), "", "���÷���ԭ��", "���÷���ԭ��ѡ��", True, True) = False Then Exit Sub
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtp����Ч����_Change()
    mblnChange = True
End Sub

Private Sub dtp����Ч����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub ClearCtlData()
    '����ؼ�����
    Dim ctl As Control
    
    On Error GoTo ErrHandler
    mdblʵ�պϼ� = 0
    mdbl������� = 0
    Set mcllCard = New Collection
    Set mobjCard = New clsSquareCard
    mobjCard.��Ч�� = "3000-01-01"
    vsfCardNo.BackColorSel = &HC0C0C0
    
    For Each ctl In Me.Controls
        If UCase(TypeName(ctl)) = "TEXTBOX" Then
            ctl.Text = "": ctl.Tag = ""
        End If
    Next
    Call SetDefaultValue
    
    vsfCardNo.Clear 1
    vsfCardNo.Rows = 0
    vsfCardNo.Cols = 0
    cmdDelete.Enabled = False
    
    '��ʾ���㵱ǰ������
    Call SetLblCaption(lbl_������, mEditType = gEd_����, True)
    If mEditType = gEd_���� Then Call Calcʵ�պϼ�
    
    chk��ֵ.value = vbUnchecked
    
    lbl(lbl_�����).Caption = "0.00": lbl(lbl_�����).Tag = ""
    lbl�ɿ�.Caption = "�տ�ϼ�"
    lbl(lbl_�ɿ�ϼ�).Caption = "0.00"
    Call SetLblCaption(lbl_�Ҳ�, mEditType = gEd_����)
    
    dtp����Ч����.value = "3000-01-01"
    dtp����Ч����.value = Null
    txt(txt_��ֵ����).Text = "0.00"
    
    mblnChange = False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetLblCaption(ByVal Index As Integer, Optional ByVal bln���� As Boolean, _
    Optional blnClear As Boolean)
    '���ù������ѿ�����ҳǩ���ı���ʾ
    '��Σ�
    '   blnClear ��ǰ�Ƿ�Ϊ�����������ʱ����
    Dim lngSpace As Long
    Dim dbl�Ҳ� As Double
    Dim blnDel As Boolean, lngCount As Long
    
    On Error GoTo ErrHandler
    Select Case Index
    Case lbl_������
        lngCount = GetCardsCount()
        lbl(lbl_������2).Caption = CStr(lngCount)
        
        lngSpace = Len(lbl(lbl_������2).Caption) + 1
        If mEditType = gEd_���� Then
            lbl(lbl_������).Caption = "������" & Space(lngSpace) & "��"
            lbl(lbl_������2).Left = lbl(lbl_������).Left + 680
            'û�п�ʱ������濨��Ϣ
            If Val(lbl(lbl_������2).Caption) = 0 And blnClear = False Then
                Call ClearCtlData
            End If
        Else
            lbl(lbl_������).Caption = "����" & Space(lngSpace) & "��"
            lbl(lbl_������2).Left = lbl(lbl_������).Left + 470
        End If
    Case lbl_�Ҳ�
        '�����Ҳ��ı���
        blnDel = mdblʵ�պϼ� - mdbl������� < 0
        dbl�Ҳ� = IIf(blnDel, -1, 1) * Val(txt(txt_�ɿ�).Text) - (mdblʵ�պϼ� - mdbl�������)
        txt(txt_�Ҳ�).Tag = dbl�Ҳ�
        txt(txt_�Ҳ�).Text = Format(dbl�Ҳ�, "0.00")
        lbl(lbl_�Ҳ�).ForeColor = IIf(dbl�Ҳ� >= 0, vbBlack, vbRed)
        txt(txt_�Ҳ�).ForeColor = IIf(dbl�Ҳ� >= 0, vbBlack, vbRed)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitData() As Boolean
    '��ʼ����������
    Dim strValue As String
    Dim rs�շ���� As ADODB.Recordset
    
    On Error GoTo ErrHandler
    Call ClearCtlData
    
    '���ѿ��ֱҴ���ʽ
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 4, 1)))
    
    If mEditType = gEd_���� Or mEditType = gEd_�޸� Then
        mrs������.Filter = ""
        Do While Not mrs������.EOF
            cbo������.AddItem NVL(mrs������!����) & "-" & NVL(mrs������!����)
            mrs������.MoveNext
        Loop
        If cbo������.ListCount > 0 Then cbo������.ListIndex = 0
    End If
    
    lvw�������.ListItems.Clear
    If mEditType = gEd_���� Or mEditType = gEd_�޸� Then
        Set rs�շ���� = zlGet�շ����
        rs�շ����.Filter = 0
        Do While Not rs�շ����.EOF
            lvw�������.ListItems.Add , NVL(rs�շ����!����), NVL(rs�շ����!����) & "-" & NVL(rs�շ����!����)
            rs�շ����.MoveNext
        Loop
        
        Call Load�������(mCardType.str�������)
    End If
    
    If mEditType = gEd_���� Or mEditType = gEd_ȡ���˿� Or mEditType = gEd_��ֵ Then
        If Load֧����ʽ() = False Then Exit Function
    ElseIf mEditType = gEd_�˿� Or mEditType = gEd_��ֵ���� Then
        If Load֧����ʽ(True) = False Then Exit Function
    End If
    
    If (mEditType = gEd_���� Or mEditType = gEd_�޸�) And mCardType.bln�ض����� Then
        Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser)
    End If
    
    If mEditType = gEd_���� Then
        Call IDKind����.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser)
    End If
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadCardData(ByVal bytMode As Byte, _
    Optional ByVal strNO As String, Optional ByVal lng��ID As Long) As Boolean
    '�������ݵ��ؼ�
    '��Σ�
    '   bytMode 0-�����ѿ�ID���أ�1-�����ѿ����ż���
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strWhere As String, lng����� As Long
    
    On Error GoTo ErrHandler
    If bytMode = 1 Then
        strWhere = " And a.���� = [2] And a.�ӿڱ�� = [3]" & vbNewLine & _
                   " And ��� = (Select Max(���) From ���ѿ���Ϣ B Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��)"
    Else
        strWhere = " And a.Id = [1]"
    End If
    strSQL = _
        "Select a.Id, a.������, a.����, a.���, a.�ɷ��ֵ, a.��Ч��, a.����ԭ��, a.����," & vbNewLine & _
        "       a.������, a.�쿨��, a.����id, a.����ʱ��, a.������, a.����ʱ��," & vbNewLine & _
        "       Mod(a.��ǰ״̬, 10) As ��ǰ״̬, a.��ע, a.������, a.���۽��, a.��ֵ�ۿ���," & vbNewLine & _
        "       a.���, a.ͣ����, a.ͣ������, a.�쿨����id, a.�������," & vbNewLine & _
        "       Decode(b.����, Null, '', b.���� || '-' || b.����) As �쿨����," & vbNewLine & _
        "       Nvl((Select 1 From ���˿������¼ Where ���ѿ�ID = a.ID And ��¼���� = 2 And Rownum < 2), 0) As �ѳ�ֵ" & vbNewLine & _
        "From ���ѿ���Ϣ A, ���ű� B" & vbNewLine & _
        "Where a.�쿨����id = b.Id(+)" & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng��ID, strNO, mlng�����)

    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ���ص�" & mCardType.str������ & "��Ϣ�������Ѿ�������ɾ����"
        Exit Function
    End If

    If bytMode = 1 Then
        '���ݿ��Ŷ�ȡ������Ϣ����ÿ���
        If CardIsValid(Val(NVL(rsTemp!id))) = False Then Exit Function
    End If
    
    mlng��ID = Val(NVL(rsTemp!id))
    lng����� = Val(NVL(rsTemp!���))
    Set mobjCard = New clsSquareCard
    With mobjCard
        .������ = NVL(rsTemp!������)
        .���� = NVL(rsTemp!����)
        
        .��ֵ�� = Val(NVL(rsTemp!�ɷ��ֵ)) = 1
        .��Ч�� = Format(NVL(rsTemp!��Ч��), "yyyy-MM-dd")
        .����ԭ�� = NVL(rsTemp!����ԭ��)
        .������ = NVL(rsTemp!������)
        .�쿨�� = NVL(rsTemp!�쿨��)
        .����ID = Val(NVL(rsTemp!����ID))
        .����ʱ�� = Format(NVL(rsTemp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        .�쿨����id = Val(NVL(rsTemp!�쿨����id))
        .�쿨���� = NVL(rsTemp!�쿨����)
        .��ע = NVL(rsTemp!��ע)
        
        .����ֵ = Val(NVL(rsTemp!������))
        .ʵ������ = Val(NVL(rsTemp!���۽��))
        .��ֵ�ۿ��� = Val(NVL(rsTemp!��ֵ�ۿ���))
        .������ = NVL(rsTemp!������)
        .����ʱ�� = Format(NVL(rsTemp!����ʱ��), "yyyy-MM-dd HH:mm:ss")
        .ͣ���� = NVL(rsTemp!ͣ����)
        .ͣ������ = Format(NVL(rsTemp!ͣ������), "yyyy-MM-dd HH:mm:ss")
        .������� = NVL(rsTemp!�������)
        
        .��ǰ״̬ = Val(NVL(rsTemp!��ǰ״̬))
        .����� = Val(NVL(rsTemp!���))
        .�ѳ�ֵ = Val(NVL(rsTemp!�ѳ�ֵ)) = 1
        .ԭ���� = NVL(rsTemp!����)
    End With
    
    Call ShowCardInfo(mobjCard)
    
    '��ȡ��ֵ��Ϣ
    Select Case mEditType
    Case gEd_��ֵ����
        strSQL = _
            "Select a.���㷽ʽ, a.ʵ�ս��, a.����, a.Ӧ�ս��, a.��ע, a.�������," & vbNewLine & _
            "       a.�����id, a.���㿨��, a.������ˮ��, a.����˵��" & vbNewLine & _
            "From ���˿������¼ A" & vbNewLine & _
            "Where a.��¼���� = 2 And ��¼״̬ = 1 And a.Id = [1]"
        Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ֵID)
        If mrsBalance.EOF Then
            ShowMsgbox "δ�ҵ�ԭ��ֵ��¼�����ܼ�����"
            Exit Function
        End If
        
        txt(txt_��ֵ����).Text = Format(Val(NVL(mrsBalance!����)), "#0.00;-#0.00")
        txt(txt_���γ�ֵ).Text = Format(Val(NVL(mrsBalance!Ӧ�ս��)), "#0.00;-#0.00")
        txt(txt_��ֵ�ɿ�).Text = Format(Val(NVL(mrsBalance!ʵ�ս��)), "#0.00;-#0.00")
        txt(txt_��ֵ˵��).Text = NVL(mrsBalance!��ע)
        Call Load֧����ʽ(True)
    Case gEd_�˿�, gEd_ȡ���˿�
        strSQL = _
            "Select a.��¼����, a.���㷽ʽ, a.ʵ�ս��, a.����, a.Ӧ�ս��, a.��ע, a.�������," & vbNewLine & _
            "       a.�����id, a.���㿨��, a.������ˮ��, a.����˵��" & vbNewLine & _
            "From ���˿������¼ A, ���˿������¼ B" & vbNewLine & _
            "Where a.���ѿ�id = b.���ѿ�id And a.������� = b.�������" & vbNewLine & _
            "      And a.��¼״̬ = [2] And a.��¼���� In (1, 2)" & vbNewLine & _
            "      And b.��¼���� = 1 And b.���ѿ�id = [1]" & vbNewLine & _
            "      And b.��� = (Select Max(���) From ���˿������¼ Where ���ѿ�id = b.���ѿ�id And ��¼���� = 1)"
        Set mrsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ID, IIf(mEditType = gEd_�˿�, 1, 3))
        mrsBalance.Filter = "��¼����=2"
        If Not mrsBalance.EOF Then
            txt(txt_��ֵ����).Text = Format(Val(NVL(mrsBalance!����)), "#0.00;-#0.00")
            txt(txt_���γ�ֵ).Text = Format(Val(NVL(mrsBalance!Ӧ�ս��)), "#0.00;-#0.00")
            txt(txt_��ֵ�ɿ�).Text = Format(Val(NVL(mrsBalance!ʵ�ս��)), "#0.00;-#0.00")
            txt(txt_��ֵ˵��).Text = NVL(mrsBalance!��ע)
        End If
        If mEditType = gEd_�˿� Then
            Call Load֧����ʽ(True)
        Else
            Call Calc���
        End If
    End Select
    
    Call Calcʵ�պϼ�

    LoadCardData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowCardInfo(objCard As clsSquareCard, Optional ByVal blnSelectNoList As Boolean)
    '��ʾ��ǰ����Ϣ
    '��Σ�
    '   blnSelectNoList �Ƿ�ѡ�񿨺��б���ʾ����Ϣ
    On Error GoTo ErrHandler

    With objCard
        mblnNotClick = True
        cbo.SeekIndex cbo������, .������
        If cbo������.ListIndex = -1 Then
            cbo������.AddItem .������
            cbo������.ListIndex = cbo������.NewIndex
        End If
        If blnSelectNoList = False Then
            'ѡ���б�ʱ����ʾ���ŵ������ı���
            txt(txt_��ʼ����).Text = .����
        End If
    
        chk��ֵ.value = IIf(.��ֵ��, vbChecked, vbUnchecked)
        txt(txt_����Ч��).Text = Format(.��Ч��, "yyyy-MM-DD")
        If txt(txt_����Ч��).Text <> "" Then
            If CDate(txt(txt_����Ч��).Text) >= CDate("3000-01-01") Then txt(txt_����Ч��).Text = ""
        End If
        If txt(txt_����Ч��).Text <> "" Then
            dtp����Ч����.value = CDate(txt(txt_����Ч��).Text)
        Else
            dtp����Ч����.value = Null
        End If
        
        txt(txt_����ԭ��).Text = .����ԭ��
        txt(txt_�쿨��).Text = .�쿨��
        txt(txt_�쿨��).Tag = .����ID
        txt(txt_�쿨����).Text = .�쿨����
        txt(txt_�쿨����).Tag = .�쿨����id
        txt(txt_��ע).Text = .��ע
        
        txt(txt_������).Text = .������
        txt(txt_��������).Text = .����ʱ��
        txt(txt_������).Text = .������
        txt(txt_��������).Text = .����ʱ��
        If txt(txt_��������).Text <> "" Then
            If CDate(txt(txt_��������).Text) >= CDate("3000-01-01") Then txt(txt_��������).Text = ""
        End If
        
        
        txt(txt_�����).Text = Format(.����ֵ, "#0.00;-#0.00")
        txt(txt_���۶�).Text = Format(.ʵ������, "#0.00;-#0.00")
        txt(txt_���۶�).Tag = .ʵ������
        
        txt(txt_��ֵ����).Text = Format(.��ֵ�ۿ���, "#0.00;-#0.00")
        txt(txt_���γ�ֵ).Text = ""
        txt(txt_��ֵ�ɿ�).Text = ""
        txt(txt_��ֵ˵��).Text = ""
    
        lbl(lbl_�����).Caption = Format(.�����, "#0.00;-#0.00")
        
        If mEditType = gEd_���� Then txt(txt_ԭ������).Text = .����
        If mEditType = gEd_���� Or mEditType = gEd_���� Then
            '�¿�����ȱʡΪԭ������
            txt(txt_�¿�����).Text = .ԭ����
            txt(txt_�¿�ȷ������).Text = .ԭ����
        End If
        
        Call Load�������(.�������)
    End With
    mblnNotClick = False
    Exit Sub
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':��;��?��" & Chr(22), Chr(KeyAscii)) > 0 Then 'Chr(22):Ctrl+V
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    If InitModulePara() = False Then Unload Me: Exit Sub
    If InitData() = False Then Unload Me: Exit Sub
    
    If mlng��ID <> 0 Then
        If CardIsValid(mlng��ID) = False Then Unload Me: Exit Sub
        If LoadCardData(0, , mlng��ID) = False Then Unload Me: Exit Sub
        
        '��ʾ���㵱ǰ������
        If mEditType = gEd_���� Then
            txt(txt_��ʼ����).Tag = 1
            Call SetLblCaption(lbl_������, True)
        End If
    End If
    If InitFace() = False Then Unload Me: Exit Sub
    
    Call CreateObjectKeyboard
    mblnChange = False
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    On Error Resume Next
    '���㶨λ
    Select Case mEditType
    Case gEd_����, gEd_�޸�
        zlControl.ControlSetFocus cbo������
    Case gEd_����
        zlControl.ControlSetFocus txt(txt_��ʼ����)
    Case gEd_��ֵ
        If txt(txt_��ʼ����).Text = "" Then
            zlControl.ControlSetFocus txt(txt_��ʼ����)
        Else
            zlControl.ControlSetFocus txt(txt_���γ�ֵ)
        End If
    Case gEd_����
        If txt(txt_ԭ������).Text = "" Then
            zlControl.ControlSetFocus txt(txt_ԭ������)
        Else
            zlControl.ControlSetFocus txt(txt_ԭ������)
        End If
    Case gEd_����
        zlControl.ControlSetFocus txt(txt_ԭ������)
    Case Else
        zlControl.ControlSetFocus txt(txt_�ɿ�)
    End Select
End Sub

Private Function Load�������(ByVal str������� As String) As Boolean
    '�����������
    '��Σ�
    '   str������� - ��ʽ������ҩ,�г�ҩ,...
    Dim i As Long, j As Long
    Dim varTemp As Variant, blnFind As Boolean
    Dim objItem As ListItem
    
    On Error GoTo ErrHandler
    For j = 1 To lvw�������.ListItems.count
        lvw�������.ListItems(j).Checked = False
    Next
    
    If str������� = "" Then Load������� = True: Exit Function
     
    varTemp = Split(str�������, ",")
    For i = 0 To UBound(varTemp)
        blnFind = False
        For j = 1 To lvw�������.ListItems.count
            If varTemp(i) = lvw�������.ListItems(j).Key Then
                lvw�������.ListItems(j).Checked = True
                blnFind = True: Exit For
            End If
        Next
        
        If blnFind = False Then
            Set objItem = lvw�������.ListItems.Add(, varTemp(i), varTemp(i))
            objItem.Checked = True
        End If
    Next
    Load������� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitModulePara() As Boolean
    '��ʼ��ģ�����
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim str�������� As String, varData As Variant, varTemp As Variant
    Dim i As Integer
    Dim ty_Temp As Ty_CardType
    
    On Error GoTo ErrHandler
    Set rsTemp = zlGet���ѿ��ӿ�()
    rsTemp.Filter = "���=" & mlng�����
    If rsTemp.EOF Then
        ShowMsgbox "δ���ֿ������Ϣ�����ܼ�����"
        Exit Function
    End If
    
    mCardType = ty_Temp '�Զ���Type��ʼ��
    With mCardType
        .str������ = NVL(rsTemp!����)
        .str����ǰ׺ = NVL(rsTemp!ǰ׺�ı�)
        .lng���ų��� = Val(NVL(rsTemp!���ų���))
        .bln�������� = Val(NVL(rsTemp!�Ƿ�����)) = 1
        .int���볤�� = Val(NVL(rsTemp!���볤��))
        .int���볤������ = Val(NVL(rsTemp!���볤������))
        .bln�ϸ���� = Val(NVL(rsTemp!�Ƿ��ϸ����)) = 1
        .byt������� = Val(NVL(rsTemp!�������))
        .str������� = NVL(rsTemp!�������)
        .bln�ض����� = Val(NVL(rsTemp!�Ƿ��ض�����)) = 1
    End With
    
    strSQL = "Select ����, ����, ȱʡ���, ȱʡ�ۿ�, ȱʡ��־ From ���ѿ�����"
    Set mrs������ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If mrs������.RecordCount = 0 Then
        ShowMsgbox "û��������ص����ѿ����ͣ�����[�ֵ����]�����ã�"
        Exit Function
    End If

    With mTy_MoudlePara
        .bln�ɿ��ӡ = Val(zlDatabase.GetPara("�ɿ��ӡ", glngSys, mlngModule)) = 1
        .bln������ֵ = Val(zlDatabase.GetPara("������ֵ", glngSys, mlngModule)) = 1
    End With
    
    str�������� = zlDatabase.GetPara("�������ѿ�����", glngSys, mlngModule)
    '����ID,�����ID|...
    varData = Split(str��������, "|")
    For i = 0 To UBound(varData)
         varTemp = Split(varData(i), ",")
         If Val(varTemp(0)) <> 0 Then
            If Val(varTemp(1)) = mlng����� Then
                mCardType.lng�������� = Val(varTemp(0)): Exit For
            End If
         End If
    Next
    
    If mEditType = gEd_���� Or mEditType = gEd_�˿� Or mEditType = gEd_ȡ���˿� _
        Or mEditType = gEd_��ֵ Or mEditType = gEd_��ֵ���� Then
        If Init֧����ʽ() = False Then Exit Function
    End If
    
    InitModulePara = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mblnFirst Or mblnChange = False Then Exit Sub
    If mEditType = gEd_���� Or mEditType = gEd_�޸� Then
        If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
    If Not mrsBalance Is Nothing Then Set mrsBalance = Nothing
End Sub

Private Sub IDKind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub IDKind����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt_Change(Index As Integer)
    Dim lng������ As Long
    
    If mblnNotClick Then Exit Sub
    mblnChange = True
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_��ʼ����
        mlng��ID = 0
        Set mobjCard = New clsSquareCard
        
        '��ʾ���㵱ǰ������
        If (mEditType = gEd_���� Or mEditType = gEd_����) And txt(txt_��ʼ����).Tag <> "" Then
            txt(txt_��ʼ����).Tag = ""
            Call SetLblCaption(lbl_������, mEditType = gEd_����)
            If mEditType = gEd_���� Then Call Calcʵ�պϼ�
        End If
        
        txt(txt_��������).Text = ""
        txt(txt_��������).Enabled = txt(txt_��ʼ����) <> ""
        Call zl_SetCtlBackColor(txt(txt_��������), Me)
    Case txt_��������
        '��ʾ���㵱ǰ������
        If mEditType = gEd_���� Or mEditType = gEd_���� Then
            lng������ = Val(txt(txt_��ʼ����).Tag)
            If lng������ <> 0 And Val(txt(txt_��������).Tag) <> 0 Then
                txt(txt_��ʼ����).Tag = "1"
                Call SetLblCaption(lbl_������, mEditType = gEd_����)
                If mEditType = gEd_���� Then Call Calcʵ�պϼ�
            End If
            txt(txt_��������).Tag = ""
        End If
    Case txt_�ɿ�
        Call SetLblCaption(lbl_�Ҳ�)
    Case txt_�쿨��
        txt(Index).Tag = ""
        If (mEditType = gEd_���� Or mEditType = gEd_�޸�) And mCardType.bln�ض����� Then
            IDKind.SetAutoReadCard txt(Index).Text = ""
        End If
    Case txt_�쿨����
        txt(Index).Tag = ""
    Case txt_��������
        txt(Index).Tag = ""
        If mEditType = gEd_���� Then
            IDKind����.SetAutoReadCard txt(Index).Text = ""
        End If
    Case txt_����
        txt(txt_ȷ������) = ""
    Case txt_�¿�����
        txt(txt_�¿�ȷ������) = ""
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Set�ɷ��ֵ()
    '���ÿɷ��ֵ
    Dim blnEnabled As Boolean
    
    On Error GoTo ErrHandler
    blnEnabled = (chk��ֵ.value = vbChecked)
    blnEnabled = blnEnabled And IIf(mEditType = gEd_����, zlStr.IsHavePrivs(mstrPrivs, "��ֵ"), mEditType = gEd_��ֵ)
    
    txt(txt_��ֵ����).Enabled = blnEnabled
    txt(txt_���γ�ֵ).Enabled = blnEnabled
    txt(txt_��ֵ�ɿ�).Enabled = blnEnabled
    txt(txt_��ֵ˵��).Enabled = blnEnabled
    
    Call zl_SetCtlBackColor(Array(txt(txt_��ֵ����), txt(txt_���γ�ֵ), txt(txt_��ֵ�ɿ�), txt(txt_��ֵ˵��)), Me)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    If txt(Index).Enabled = False Or txt(Index).Locked Then Exit Sub
    Select Case Index
    Case txt_��ʼ����, txt_��������, txt_ԭ������, txt_�¿�����
        
    Case txt_����, txt_ԭ������, txt_�¿�����
        Call OpenPassKeyboard(txt(Index), False)
    Case txt_ȷ������, txt_�¿�ȷ������
        Call OpenPassKeyboard(txt(Index), True)
    Case txt_�쿨��
        If (mEditType = gEd_���� Or mEditType = gEd_�޸�) And mCardType.bln�ض����� Then
            IDKind.SetAutoReadCard txt(Index).Text = ""
        End If
    Case txt_��������
        If mEditType = gEd_���� Then
            IDKind����.SetAutoReadCard txt(Index).Text = ""
        End If
    End Select
    zlControl.TxtSelAll txt(Index)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim str���� As String, str���� As String, lngID As Long
    Dim strCardNo As String, intIndexTemp As Integer
    
    On Error GoTo ErrHandler
    If KeyCode <> vbKeyReturn Then Exit Sub
    Select Case Index
    Case txt_�쿨��
        If txt(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txt(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        
        If mCardType.bln�ض����� = False Then
            'ѡ����Ա
            lngID = Val(txt(txt_�쿨����).Tag)
            If Select��Աѡ����(Me, txt(Index), Trim(txt(Index).Text), lngID, , True, , , , , , , "") = False Then
                zlCommFun.PressKey vbKeyTab
            End If
            If mEditType = gEd_���� Then
                '�쿨�˾��ǽɿ���
                txt(txt_�ɿ���).Text = txt(txt_�쿨��).Text
                txt(txt_�ɿ���).Tag = txt(txt_�쿨��).Tag
            End If
    
            '��Ҫ��ȡȱʡ����:
            If zl_From��Ա��ȡȱʡ����(Val(txt(txt_�쿨��).Tag), str����, str����, lngID) Then
                txt(txt_�쿨����).Text = str���� & "-" & str����
                txt(txt_�쿨����).Tag = lngID
            End If
        End If
    Case txt_�쿨����
        'ѡ����
        If txt(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If Trim(txt(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        'ѡ��ȱʡ����
        lngID = Val(txt(txt_�쿨��).Tag)
        If Select����ѡ����(Me, txt(Index), Trim(txt(Index).Text), "", IIf(lngID = 0, False, True), "", 0, _
            "����ѡ����", , , , , lngID) = False Then Exit Sub
    Case txt_����ԭ��
        'ѡ�񷢿�ԭ��
        If Trim(txt(Index).Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        
        If zl_SelectAndNotAddItem(Me, txt(Index), Trim(txt(Index).Text), "���÷���ԭ��", _
            "���÷���ԭ��ѡ��", True, True, , , , True) = False Then
            zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
    Case txt_����, txt_�¿�����
        If CheckPassword(txt(Index), , True) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_ȷ������, txt_�¿�ȷ������
        intIndexTemp = IIf(Index = txt_ȷ������, txt_����, txt_�¿�����)
        If CheckPassword(txt(Index), txt(intIndexTemp)) = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_�����
        If CheckInput����� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_���۶�
        If CheckInputʵ�����۶� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_��ֵ�ɿ�
        If CheckInputʵ�ʳ�ֵ�ɿ� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_���γ�ֵ
        If CheckInput���γ�ֵ = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_��ֵ����
        If CheckInput��ֵ���� = False Then Exit Sub
        zlCommFun.PressKey vbKeyTab
    Case txt_��ʼ����, txt_��������, txt_ԭ������, txt_�¿�����
        
    Case txt_�쿨��, txt_��������
        
    Case Else
        zlCommFun.PressKey vbKeyTab
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CreateObjectKeyboard() As Boolean
    '�����������
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    If mobjKeyboard Is Nothing Then Exit Function
    CreateObjectKeyboard = True
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional blnȷ������ As Boolean = False) As Boolean
    '�������������
    On Error GoTo ErrHandler
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, blnȷ������) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '�ر������������
    On Error GoTo ErrHandler
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput�����() As Boolean
    '��鿨���
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_�����).Text), 16, True, False, txt(txt_�����).hWnd, "�����") = False Then
        zlControl.TxtSelAll txt(txt_�����): Exit Function
    End If

    If Val(txt(txt_�����).Text) < Val(txt(txt_���۶�).Text) Then
        txt(txt_���۶�).Text = txt(txt_�����).Text
    End If
    CheckInput����� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInputʵ�����۶�() As Boolean
    '���ʵ�����۶�
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_���۶�).Text), 16, True, False, txt(txt_���۶�).hWnd, "ʵ������") = False Then
        zlControl.TxtSelAll txt(txt_���۶�): Exit Function
    End If
    If Val(txt(txt_�����).Text) < Val(txt(txt_���۶�).Text) Then
        ShowMsgbox "ʵ�����۶�ܴ��ڿ������飡"
        zlControl.ControlSetFocus txt(txt_���۶�)
        zlControl.TxtSelAll txt(txt_���۶�)
        Exit Function
    End If
    CheckInputʵ�����۶� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput��ֵ����() As Boolean
    '����ֵ�����Ƿ�Ϸ�
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_��ֵ����).Text), 3, True, False, txt(txt_��ֵ����).hWnd, "��ֵ����") = False Then
        zlControl.TxtSelAll txt(txt_��ֵ����): Exit Function
    End If
    If Val(txt(txt_��ֵ����).Text) > 100 Then
        ShowMsgbox "��ֵ���ʲ��ܴ���100%�����飡"
        zlControl.ControlSetFocus txt(txt_��ֵ����)
        zlControl.TxtSelAll txt(txt_��ֵ����): Exit Function
    End If
    If Val(txt(txt_��ֵ����).Text) < 0 Then
        ShowMsgbox "��ֵ���ʲ���С��0�����飡"
        zlControl.ControlSetFocus txt(txt_��ֵ����)
        zlControl.TxtSelAll txt(txt_��ֵ����): Exit Function
    End If
    CheckInput��ֵ���� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput���γ�ֵ() As Boolean
    '��鱾�γ�ֵ
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_���γ�ֵ).Text), 16, True, False, txt(txt_���γ�ֵ).hWnd, "���γ�ֵ") = False Then
        zlControl.TxtSelAll txt(txt_���γ�ֵ): Exit Function
    End If
    If Val(txt(txt_���γ�ֵ).Text) < Val(txt(txt_��ֵ�ɿ�).Text) Then
        txt(txt_��ֵ�ɿ�).Text = txt(txt_���γ�ֵ).Text
    End If
    CheckInput���γ�ֵ = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInputʵ�ʳ�ֵ�ɿ�() As Boolean
    '��鱾�γ�ֵ
    On Error GoTo ErrHandler
    If zlDblIsValid(Trim(txt(txt_��ֵ�ɿ�).Text), 16, True, False, txt(txt_��ֵ�ɿ�).hWnd, "ʵ�ʳ�ֵ�ɿ�") = False Then
        zlControl.TxtSelAll txt(txt_��ֵ�ɿ�): Exit Function
    End If
    If Val(txt(txt_���γ�ֵ).Text) < Val(txt(txt_��ֵ�ɿ�).Text) Then
        ShowMsgbox "ʵ�ʳ�ֵ�ɿ�ܴ��ڱ��γ�ֵ�����飡"
        zlControl.ControlSetFocus txt(txt_��ֵ�ɿ�)
        zlControl.TxtSelAll txt(txt_��ֵ�ɿ�): Exit Function
    End If
    CheckInputʵ�ʳ�ֵ�ɿ� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim blnCard As Boolean, blnPass As Boolean
    Dim str���� As String, lng����ID As Long
    Dim objIDKind As IDKindNew
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_��ֵ����
        'ֻ������������
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    Case txt_�����, txt_���۶�, txt_���γ�ֵ, txt_��ֵ�ɿ�, txt_�ɿ�
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m���ʽ)
    Case txt_��ʼ����, txt_��������, txt_ԭ������, txt_�¿�����
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m�ı�ʽ)
        If InStr(1, "'~��|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        
        If Index = txt_�������� Then
            If KeyAscii = 13 And Trim(txt(Index)) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
        End If
        
        'Сд��ĸת��Ϊ��д
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
        Call BrushCard(txt(Index), KeyAscii)
    Case txt_�쿨��, txt_��������
        If mCardType.bln�ض����� Then
            If Index = txt_�쿨�� Then
                Set objIDKind = IDKind
            Else
                Set objIDKind = IDKind����
            End If
            
            If IsCardType(objIDKind, "����") Then
                '105567:���ϴ�,2017/5/25,���ż��ܵ��µ�һ������ƴ�����ܴ������뷨
                blnPass = txt(Index).PasswordChar <> ""
                If Not (InStr("-+*", Left(txt(Index).Text, 1)) > 0 And IsNumeric(Mid(txt(Index).Text, 2))) Then
                    blnCard = zlCommFun.InputIsCard(txt(Index), KeyAscii, objIDKind.ShowPassText)
                End If
                txt(Index).IMEMode = 0
                blnPass = blnPass And txt(Index).PasswordChar = ""
                If blnPass Then
                    If txt(Index).SelLength = Len(txt(Index).Text) Then
                        txt(Index).Text = ""
                    End If
                    SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
                End If
            ElseIf IsCardType(objIDKind, "�����") Or IsCardType(objIDKind, "סԺ��") Or IsCardType(objIDKind, "�ֻ���") Then
                If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                    If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
                End If
            Else
                txt(Index).PasswordChar = IIf(objIDKind.ShowPassText, "*", "")
                '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
                txt(Index).IMEMode = 0
            End If
        
            If blnCard And Len(txt(Index).Text) = objIDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
                Or KeyAscii = 13 And Trim(txt(Index).Text) <> "" Then
                If KeyAscii <> 13 Then
                    txt(Index).Text = txt(Index).Text & Chr(KeyAscii): txt(Index).SelStart = Len(txt(Index).Text)
                End If
                KeyAscii = 0
                
                strTemp = txt(Index).Text
                If Not GetPatient(objIDKind, txt(Index), txt(Index).Text, blnCard, str����, lng����ID) Then
                    '�������
                    If LoadPatientCard(mlng�����, lng����ID) = False Then
                        txt(Index).Text = strTemp
                        zlControl.TxtSelAll txt(Index)
                        Exit Sub
                    End If
                Else
                    If Index = txt_�������� Then
                        '���ز�����Ч��
                        If LoadPatientCard(mlng�����, lng����ID) = False Then
                            txt(Index).Text = strTemp
                            zlControl.TxtSelAll txt(Index)
                            Exit Sub
                        End If
                    End If
                    txt(Index).Text = str����
                    txt(Index).Tag = lng����ID
                    txt(Index).PasswordChar = ""
                    '55766:�ı�����һbug:�����Ϊ������ʾ,�����óɷ�������ʾ��,�����������
                    txt(Index).IMEMode = 0
                    zlCommFun.PressKey vbKeyTab: Exit Sub
                End If
            End If
        End If
    Case txt_����, txt_ȷ������, txt_�¿�����, txt_�¿�ȷ������
        Call CheckInputPassWord(KeyAscii, mCardType.byt������� = 1)
    Case Else
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m�ı�ʽ)
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetPatient(ByVal objIDKind As IDKindNew, txtEdit As TextBox, _
    ByVal strInput As String, ByVal blnCard As Boolean, _
    ByRef str���� As String, ByRef lng����ID As Long) As Boolean
    '��ȡ������Ϣ
    '���:blnCard=��ʾ�Ƿ���￨ˢ��
    '����:
    '����:���˶�ȡ�ɹ�,����true,���򷵻�False
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String
    Dim lng�����ID As Long, blnCancel As Boolean
    Dim strPassWord As String, strErrMsg As String
    Dim blnIsMobileNO As Boolean
    
    On Error GoTo errH
    str���� = "": lng����ID = 0
    blnIsMobileNO = IDKind.IsMobileNO(strInput)
    If blnCard And IsCardType(objIDKind, "����") And (InStr("-+*", Left(strInput, 1)) = 0 And IsNumeric(Mid(strInput, 2))) Then  'ˢ����ȱʡ�Ŀ�
        If objIDKind.Cards.��ȱʡ������ And Not objIDKind.GetfaultCard Is Nothing Then
            lng�����ID = objIDKind.GetfaultCard.�ӿ����
        ElseIf objIDKind.GetCurCard.�ӿ���� > 0 Then
            lng�����ID = objIDKind.GetCurCard.�ӿ����
        Else
            lng�����ID = -1
        End If
        
        '����|ȫ��|������־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�);��
        If GetPatiID(lng�����ID, strInput, True, lng����ID, strPassWord, strErrMsg) = False Then
            If blnIsMobileNO Then
                '�ֻ��Ų���
                If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            Else
                GoTo NotFoundPati:
            End If
        End If
        If lng����ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.����ID=[1]"
        strInput = "-" & lng����ID
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '����ID
        strWhere = strWhere & " And A.����ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '�����
        strWhere = strWhere & " And A.�����=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then 'סԺ��
        strWhere = strWhere & " And A.סԺ��=[1]"
    Else
        Select Case objIDKind.GetCurCard.����
            Case "����", "��������￨"
                strPati = _
                    "Select a.����id As ID, a.����id, a.����, a.�Ա�, a.����, a.�����, a.סԺ��," & vbNewLine & _
                    "       a.��������, a.���֤��, a.��ͥ��ַ, a.������λ" & vbNewLine & _
                    "From ������Ϣ A" & vbNewLine & _
                    "Where a.ͣ��ʱ�� Is Null And a.���� Like [1] And Rownum < 101" & vbNewLine & _
                    "Order By ����"
                
                vRect = zlControl.GetControlRect(txtEdit.hWnd)
                Set rsTemp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "����ѡ��", 1, "", "��ѡ����", False, False, True, _
                    vRect.Left, vRect.Top, txtEdit.Height, blnCancel, False, True, strInput & "%")
                If blnCancel Then Exit Function
                If rsTemp Is Nothing Then GoTo NotFoundPati:
                If rsTemp.State <> 1 Then GoTo NotFoundPati:
                If rsTemp.RecordCount = 0 Then GoTo NotFoundPati:
                If Val(NVL(rsTemp!����ID)) = 0 Then GoTo NotFoundPati:
                
                strInput = "-" & rsTemp!����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "ҽ����"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.ҽ����=[2]"
             Case "���֤��", "���֤", "�������֤"
                strInput = UCase(strInput)
                '�����:54197
                 If GetPatiID("���֤", strInput, False, lng����ID, strPassWord, strErrMsg, , , , False) = False Then lng����ID = 0
                 strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "IC����", "IC��"
                strInput = UCase(strInput)
                If GetPatiID("IC��", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case "�����"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.�����=[2]"
            Case "�ֻ���", "�ֻ�"
                If GetPatiID("�ֻ���", strInput, False, lng����ID, strPassWord, strErrMsg) = False Then lng����ID = 0
                strInput = "-" & lng����ID
                strWhere = strWhere & " And A.����ID=[1]"
            Case Else
                '�������ĺ���
                If Val(objIDKind.GetCurCard.�ӿ����) >= 0 Then
                    lng�����ID = objIDKind.GetCurCard.�ӿ����
                    If GetPatiID(lng�����ID, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng����ID = 0 Then GoTo NotFoundPati:
                Else
                    If GetPatiID(objIDKind.GetCurCard.����, strInput, False, lng����ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng����ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.����ID=[1]"
                strInput = "-" & lng����ID
        End Select
    End If
    
    '��ȡ������Ϣ
    strSQL = "Select A.����id,A.���� From ������Ϣ A Where A.ͣ��ʱ�� is NULL" & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(strInput, 2)), strInput)
    If rsTemp.EOF Then GoTo NotFoundPati:
    
    str���� = NVL(rsTemp!����)
    lng����ID = Val(NVL(rsTemp!����ID))
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
NotFoundPati:
    If blnCard Then
        MsgBox "����ȷ��������Ϣ�������Ƿ���ȷˢ����    ", vbInformation + vbOKOnly, gstrSysName
    Else
        MsgBox "������Ϣδ�ҵ��������Ƿ�������ȷ��", vbInformation + vbOKOnly, gstrSysName
    End If
End Function

Private Function LoadPatientCard(ByVal lng����� As Long, ByVal lng����ID As Long) As Boolean
    '���ܣ����ز��˵�ǰ��Ч���ѿ�
    '��Σ�
    '   lng����� ���ѿ������
    '   lng����ID ����ID
    '���Σ�
    '���أ���ȡ����Ч���ѿ��򷵻�TRUE,���򷵻�FALSE
    '˵����
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mEditType <> gEd_���� Then LoadPatientCard = True: Exit Function
    Call ClearCtlData
    mlng��ID = 0
    cboԭ������.Clear
    cboԭ������.Tag = ""
    
    If lng����ID = 0 Then Exit Function
    strSQL = _
        "Select a.Id, a.����" & vbNewLine & _
        "From ���ѿ���Ϣ A" & vbNewLine & _
        "Where a.��� = (Select Max(���) From ���ѿ���Ϣ B Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��)" & vbNewLine & _
        "      And a.��ǰ״̬ = 1 And a.�ӿڱ�� = [1] And a.����id = [2]" & vbNewLine & _
        "Order By a.����ʱ�� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�����, lng����ID)
        If rsTemp.EOF Then
        ShowMsgbox "δ���ҵ����ڸò��˵���Ч��" & mCardType.str������ & "��Ƭ��Ϣ��"
        Exit Function
    End If
    
    Do While Not rsTemp.EOF
        cboԭ������.AddItem NVL(rsTemp!����)
        cboԭ������.ItemData(cboԭ������.NewIndex) = Val(NVL(rsTemp!id))
        rsTemp.MoveNext
    Loop
    cboԭ������.ListIndex = 0
    
    LoadPatientCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    'ˢ��
    Dim blnCard As Boolean
    Dim lng������ As Long
    
    On Error GoTo ErrHandler
    If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If (mEditType = gEd_���� Or (mEditType = gEd_���� Or mEditType = gEd_����) And objEdit.Index = txt_�¿�����) _
            And Len(objEdit.Text) = mCardType.lng���ų��� - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(objEdit.Text) <> "" Then '�¿��ﵽ���ų��Ȼ������س����ҿ���Ϣ
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        
        If mEditType <> gEd_���� _
            And Not ((mEditType = gEd_���� Or mEditType = gEd_����) And objEdit.Index = txt_�¿�����) Then
            If LoadCardData(1, objEdit.Text) = False Then zlControl.TxtSelAll objEdit: Exit Sub
            If mEditType = gEd_��ֵ Then
                txt(txt_��ֵ����).Enabled = (chk��ֵ.value = vbChecked)
                txt(txt_���γ�ֵ).Enabled = (chk��ֵ.value = vbChecked)
                txt(txt_��ֵ�ɿ�).Enabled = (chk��ֵ.value = vbChecked)
                txt(txt_��ֵ˵��).Enabled = (chk��ֵ.value = vbChecked)
                Call zl_SetCtlBackColor(Array(txt(txt_��ֵ����), txt(txt_���γ�ֵ), txt(txt_��ֵ�ɿ�), txt(txt_��ֵ˵��)), Me)
            End If
        Else
            If CheckInput����(False, lng������) = False Then zlControl.TxtSelAll objEdit: Exit Sub
        End If
        
        '��ʾ���㵱ǰ������
        If mEditType = gEd_���� Then
            txt(txt_��ʼ����).Tag = lng������
            txt(txt_��������).Tag = IIf(lng������ > 1, "1", "") '����Ƿ���������
            Call SetLblCaption(lbl_������)
            Call Calcʵ�պϼ�
        ElseIf mEditType = gEd_���� Then
            txt(txt_��ʼ����).Tag = 1
            Call SetLblCaption(lbl_������, True)
        End If
        zlCommFun.PressKey vbKeyTab
    ElseIf KeyAscii = 13 And Trim(objEdit.Text) = "" Then
        zlCommFun.PressKey vbKeyTab
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Dim intIndexTemp As Integer
    Dim strPassWord As String
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_��ʼ����, txt_��������, txt_�¿�����
        If Not (mEditType <> gEd_���� And Index = txt_��ʼ����) Then
            If Trim(txt(Index).Text) <> "" And Len(Trim(txt(Index).Text)) <> mCardType.lng���ų��� Then
                ShowMsgbox "���ų���ӦΪ" & mCardType.lng���ų��� & "λ�����飡"
                zlControl.ControlSetFocus txt(Index)
                zlControl.TxtSelAll txt(Index)
            End If
        End If
    Case txt_�����
        If CheckInput����� = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        Call Calc���
        Calcʵ�պϼ�
    Case txt_���۶�
        If CheckInputʵ�����۶� = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        Call Calc���
        Calcʵ�պϼ�
    Case txt_���γ�ֵ
        If CheckInput���γ�ֵ = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        txt(txt_��ֵ�ɿ�).Text = Format(Val(txt(Index).Text) * (Round(Val(txt(txt_��ֵ����)) / 100, 6)), "0.00")
        Call Calc���
        Calcʵ�պϼ�
    Case txt_��ֵ����
        If CheckInput��ֵ���� = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        txt(txt_��ֵ�ɿ�).Text = Format(Val(txt(txt_���γ�ֵ).Text) * (Round(Val(txt(txt_��ֵ����)) / 100, 4)), "0.00")
        Call Calc���
        Calcʵ�պϼ�
    Case txt_��ֵ�ɿ�
        If CheckInputʵ�ʳ�ֵ�ɿ� = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
        txt(Index).Text = Format(Val(txt(Index).Text), "0.00")
        If Val(txt(txt_���γ�ֵ).Text) <> 0 Then
            txt(txt_��ֵ����).Text = Format((Round(Val(txt(txt_��ֵ�ɿ�).Text) / Val(txt(txt_���γ�ֵ).Text), 6)) * 100, "0.00")
        Else
             txt(txt_���γ�ֵ).Text = txt(txt_��ֵ�ɿ�).Text
        End If
        Call Calc���
        Calcʵ�պϼ�
    Case txt_����, txt_�¿�����
        If CheckPassword(txt(Index), , True) = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
    Case txt_ȷ������, txt_�¿�ȷ������
        intIndexTemp = IIf(Index = txt_ȷ������, txt_����, txt_�¿�����)
        If CheckPassword(txt(Index), txt(intIndexTemp)) = False Then
            zlControl.ControlSetFocus txt(Index)
            zlControl.TxtSelAll txt(Index)
        End If
    Case txt_ԭ������
        If Trim(txt(Index).Text) <> "" Then
            strPassWord = zlCommFun.zlStringEncode(txt(Index).Text)  '�������
            If mobjCard.ԭ���� <> strPassWord Then
                ShowMsgbox "ԭ����������������������룡"
                zlControl.ControlSetFocus txt(Index)
                zlControl.TxtSelAll txt(Index)
            End If
        End If
    Case txt_�쿨��
        If (mEditType = gEd_���� Or mEditType = gEd_�޸�) And mCardType.bln�ض����� Then
            IDKind.SetAutoReadCard False
        End If
    Case txt_�쿨����
        If txt(Index).Tag = "" Then txt(Index).Text = ""
    Case txt_����, txt_ȷ������, txt_ԭ������, txt_�¿�����, txt_�¿�ȷ������
        Call ClosePassKeyboard(txt(Index))
    Case txt_��������
        If mEditType = gEd_���� Then
            IDKind����.SetAutoReadCard False
        End If
    End Select
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub Calcʵ�պϼ�(Optional blnOnlyCalc As Boolean)
    '����ʵ�պϼ�
    Dim dblʵ�պϼ� As Double
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If Not (mEditType = gEd_���� _
        Or mEditType = gEd_��ֵ Or mEditType = gEd_��ֵ���� _
        Or mEditType = gEd_�˿� Or mEditType = gEd_ȡ���˿�) Then Exit Sub
    
    If mEditType = gEd_���� Or mEditType = gEd_�˿� Or mEditType = gEd_ȡ���˿� Then
        dblʵ�պϼ� = Val(txt(txt_���۶�).Text)
    End If
    If chk��ֵ.value = vbChecked Then
        dblʵ�պϼ� = dblʵ�պϼ� + Val(txt(txt_��ֵ�ɿ�).Text)
    End If
    
    If mEditType = gEd_���� Then
        dblʵ�պϼ� = dblʵ�պϼ� * Val(lbl(lbl_������2).Caption)
    End If
    
    If mEditType = gEd_��ֵ���� Or mEditType = gEd_�˿� Then
        dblʵ�պϼ� = -1 * dblʵ�պϼ� '�˿�
    End If
    
    mdblʵ�պϼ� = dblʵ�պϼ�
    If blnOnlyCalc Then Exit Sub
    
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Calc���()
    '�������
    Dim dbl��� As Double
    
    On Error GoTo ErrHandler
    If Not (mEditType = gEd_���� Or mEditType = gEd_ȡ���˿� Or mEditType = gEd_��ֵ) Then Exit Sub
    If mEditType = gEd_���� Or mEditType = gEd_ȡ���˿� Then
        dbl��� = Val(txt(txt_�����).Text)
    Else
        dbl��� = mobjCard.�����
    End If
    dbl��� = dbl��� + IIf(chk��ֵ.value = 0, 0, Val(txt(txt_���γ�ֵ).Text))
    lbl(lbl_�����).Caption = Format(dbl���, "0.00")
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo������_Click()
    If mblnNotClick Then Exit Sub
    mblnChange = True
    
    On Error GoTo ErrHandler
    '��������ȱʡֵ
    If mEditType <> gEd_���� Then Exit Sub
    Call SetDefaultValue
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk��ֵ_Click()
    If mblnNotClick Then Exit Sub
    
    On Error GoTo ErrHandler
    mblnChange = True
    Call Set�ɷ��ֵ
    Call Calc���
    Call Calcʵ�պϼ�
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDefaultValue()
    '����ȱʡֵ
    On Error GoTo ErrHandler
    mrs������.Filter = "����='" & zlStr.NeedCode(cbo������.Text) & "'"
    If mrs������.EOF Then Exit Sub
    
    txt(txt_��ֵ����).Text = Format(Val(NVL(mrs������!ȱʡ�ۿ�, 100)), "0.00")
    txt(txt_��ֵ����).Tag = txt(txt_��ֵ����).Text
    txt(txt_��ֵ�ɿ�).Text = Format(Val(NVL(mrs������!ȱʡ�ۿ�, 100)) * Val(txt(txt_���γ�ֵ).Text) / 100, "0.00")
    
    txt(txt_�����).Text = Format(Val(NVL(mrs������!ȱʡ���)), "0.00")
    txt(txt_���۶�).Text = Format(Val(NVL(mrs������!ȱʡ���)) * (txt(txt_��ֵ����).Text / 100), "0.00")
    
    Call Calc���
    Call Calcʵ�պϼ�
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Init֧����ʽ() As Boolean
    '��ʼ��֧����ʽ
    '˵����
    '   ֻ�����ֽ�֧Ʊ���������Ľ��㷽ʽ
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim i As Long, objCards As Cards, objCard As Card
    Dim lngKey As Long
    
    On Error GoTo ErrHandler
    Set mobjPayCards = New Cards
    
    Set rsTemp = Get���㷽ʽ("���ѿ�")
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
            '���:bytType-  0-����ҽ�ƿ�;
        '                        1-���õ�ҽ�ƿ�,
        '                        2-���д��������˻���������
        '                        3-���õ������˻���ҽ�ƿ�
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.count
                If objCards(i).���㷽ʽ = NVL(rsTemp!����) Then blnFind = True: Exit For
            Next
            If Not blnFind Then
                If (Val(NVL(rsTemp!����)) = 1 Or Val(NVL(rsTemp!����)) = 2) _
                    And Val(NVL(rsTemp!Ӧ����)) = 0 Then
                    Set objCard = New Card
                    objCard.���� = Mid(NVL(!����), 1, 1)
                    objCard.�ӿڱ��� = NVL(!����)
                    objCard.�ӿڳ����� = ""
                    objCard.�ӿ���� = -1 * lngKey
                    objCard.���㷽ʽ = NVL(!����)
                    objCard.���� = NVL(!����)
                    objCard.���� = True
                    objCard.ȱʡ��־ = Val(NVL(rsTemp!ȱʡ)) = 1
                    objCard.���� = True
                    objCard.�������� = Val(!����)
                    
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '��������
    For i = 1 To objCards.count
        rsTemp.Filter = "����='" & objCards(i).���㷽ʽ & "'"
        If Not rsTemp.EOF And Not objCards(i).���ѿ� Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.count = 0 Then
        ShowMsgbox "���ѿ�����û�п��õĽ��㷽ʽ�����ȵ������㷽ʽ���������á�"
        Exit Function
    End If
    Init֧����ʽ = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load֧����ʽ(Optional ByVal blnDel As Boolean) As Boolean
    '����֧����ʽ
    '˵��:
    '   ȱʡ���㷽ʽ�Ĺ�������˳�����£�
    '   1.���㷽ʽӦ�������õ�ȱʡ��
    '   2.����Ϊ"1-�ֽ���㷽ʽ"�Ľ��㷽ʽ
    Dim objCard As Card, i As Long
    Dim str���㷽ʽ As String, blnExists As Boolean
    
    On Error GoTo ErrHandler
    mlngPre֧����ʽ = 0

    mblnNotClick = True
    With cbo֧����ʽ
        .Clear
        For i = 1 To mobjPayCards.count
            Set objCard = mobjPayCards(i)
            If objCard.���� And Not objCard.���ѿ� _
                And InStr(str���㷽ʽ & "|", "|" & objCard.���㷽ʽ & "|") = 0 Then
                '�����˻���֧����ʽ��ʾΪҽ�ƿ����ƣ�������ʾ���㷽ʽ
                If objCard.�ӿ���� > 0 Then
                    If blnDel And objCard.�Ƿ�ת�ʼ����� = False Then
                        If ExitsInBalance(objCard.���㷽ʽ) Then
                            .AddItem objCard.����
                            .ItemData(.NewIndex) = i
                            If Not (objCard.�Ƿ����� And objCard.�Ƿ�ȱʡ����) Then .ListIndex = .NewIndex
                        End If
                    Else
                        .AddItem objCard.����
                        .ItemData(.NewIndex) = i
                    End If
                Else
                    .AddItem objCard.���㷽ʽ
                    .ItemData(.NewIndex) = i
                    If blnDel Then
                        If ExitsInBalance(objCard.���㷽ʽ) Then .ListIndex = .NewIndex
                    End If
                End If
                
                str���㷽ʽ = str���㷽ʽ & "|" & objCard.���㷽ʽ
            End If
            
            '����ȱʡֵ
            If objCard.ȱʡ��־ And .ListIndex < 0 Then .ListIndex = .NewIndex
            If objCard.�������� = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
        Next
            
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Call cbo֧����ʽ_Click
    Load֧����ʽ = True
    Exit Function
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExitsInBalance(ByVal str���㷽ʽ As String) As Boolean
    '�жϽ��㷽ʽ�Ƿ�������տ���㷽ʽ��
    On Error GoTo ErrHandler
    If mrsBalance Is Nothing Then Exit Function
    With mrsBalance
        .Filter = ""
        Do While Not .EOF
            If NVL(!���㷽ʽ) = str���㷽ʽ Then
                ExitsInBalance = True: Exit Function
            End If
            .MoveNext
        Loop
    End With
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get�������() As String
    '��ȡ�������
    Dim strType As String, i As Long
    
    On Error GoTo ErrHandler
    With lvw�������
         For i = 1 To .ListItems.count
            If .ListItems.Item(i).Checked Then
                strType = strType & "," & lvw�������.ListItems(i).Key
            End If
         Next
         If strType <> "" Then strType = Mid(strType, 2)
    End With
    Get������� = strType
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCurCard(ByRef objCard As Card) As Boolean
    '��ȡ��ǰ֧����
    '����:
    '   objCard-���ص�ǰ�˿��ɿ�Ŀ�����
    '����:�ɹ�,���ؿ�����
    Dim intIndex As Integer
    
    On Error GoTo ErrHandler
    intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    GetCurCard = True
    Exit Function
ErrHandler:
    Set objCard = New Card
End Function
 
'�ؼ������Ƿ�ƥ��
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    On Error GoTo ErrHandler
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "����", "��������￨"
          IsCardType = IDKindCtl.GetCurCard.���� Like "����*"
     Case "���֤", "���֤��", "�������֤"
          IsCardType = IDKindCtl.GetCurCard.���� Like "*���֤*"
     Case "IC����", "IC��"
          IsCardType = IDKindCtl.GetCurCard.���� Like "IC��*"
     Case "ҽ����"
          IsCardType = IDKindCtl.GetCurCard.���� = "ҽ����"
     Case "�����"
          IsCardType = IDKindCtl.GetCurCard.���� = "�����"
     Case "סԺ��"
          IsCardType = IDKindCtl.GetCurCard.���� = "סԺ��"
     Case "�ֻ���"
          IsCardType = IDKindCtl.GetCurCard.���� = "�ֻ���"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.����
            Else
                If IDKindCtl.GetCurCard.�ӿ���� <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.�ӿ���� = Val(strCardName)
            End If
     End Select
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
    '������Ч�Լ��
    Dim strPassWord As String
    
    On Error GoTo ErrHandler
    Select Case mEditType
    Case gEd_����
        If CheckInput����(True) = False Then Exit Function
        If CheckInput() = False Then Exit Function
        If Check�ɿ���� = False Then Exit Function
    Case gEd_�޸�
        If CheckInput����(True) = False Then
            zlControl.ControlSetFocus txt(txt_��ʼ����)
            Exit Function
        End If
        If CheckInput() = False Then Exit Function
    Case gEd_��ֵ
        If Trim(txt(txt_��ʼ����)) = "" Or mlng��ID = 0 Then
            ShowMsgbox "��ˢ���������ֵ���ţ�"
            zlControl.ControlSetFocus txt(txt_��ʼ����)
            Exit Function
        End If
        If mEditType = gEd_��ֵ And Val(txt(txt_���γ�ֵ).Text) = 0 Then
            ShowMsgbox "��ֵ����Ϊ�㣡"
            zlControl.ControlSetFocus txt(txt_���γ�ֵ)
            Exit Function
        End If
        If CheckInput����(True) = False Then
            zlControl.ControlSetFocus txt(txt_��ʼ����)
            Exit Function
        End If
       If Check�ɿ���� = False Then Exit Function
    Case gEd_��ֵ����
        If CheckInput����(True) = False Then
            zlControl.ControlSetFocus txt(txt_��ʼ����)
            Exit Function
        End If
       If Check�ɿ���� = False Then Exit Function
    Case gEd_�˿�
        If CheckInput����(True) = False Then
            zlControl.ControlSetFocus txt(txt_��ʼ����)
            Exit Function
        End If
        If Check�ɿ���� = False Then Exit Function
    Case gEd_ȡ���˿�
        If CheckInput����(True) = False Then
            zlControl.ControlSetFocus txt(txt_��ʼ����)
            Exit Function
        End If
        If Check�ɿ���� = False Then Exit Function
    Case gEd_����
        If CheckInput����(True) = False Then Exit Function
    Case gEd_ȡ������
        If CheckInput����(True) = False Then Exit Function
    Case gEd_����, gEd_����
        If mEditType = gEd_���� Then
            If Trim(txt(txt_ԭ������)) = "" Or mlng��ID = 0 Then
                ShowMsgbox "��ˢ��������ԭ�����ţ�"
                zlControl.ControlSetFocus txt(txt_ԭ������)
                Exit Function
            End If
            
            If CheckInput����(True) = False Then Exit Function
            
            strPassWord = zlCommFun.zlStringEncode(txt(txt_ԭ������).Text)  '�������
            If mobjCard.ԭ���� <> strPassWord Then
                ShowMsgbox "ԭ����������������������룡"
                zlControl.ControlSetFocus txt(txt_ԭ������)
                Exit Function
            End If
        Else
            If Val(txt(txt_��������).Tag) = 0 Then
                ShowMsgbox "����¼�벡�ˣ���ѡ����Ҫ������ԭ�����ţ�"
                zlControl.ControlSetFocus txt(txt_��������)
                Exit Function
            End If
            
            If cboԭ������.Text = "" Or mlng��ID = 0 Then
                ShowMsgbox "��ѡ��ԭ�����ţ�"
                zlControl.ControlSetFocus cboԭ������
                Exit Function
            End If
            
            If CheckInput����(True) = False Then Exit Function
        End If
        
        If CheckPassword(txt(txt_�¿�����), txt(txt_�¿�ȷ������)) = False Then Exit Function
        If mobjCard.ԭ���� <> Trim(txt(txt_�¿�����).Text) Then
            If zlCommFun.StrIsValid(Trim(txt(txt_�¿�����).Text), 20, txt(txt_�¿�����).hWnd, "����") = False Then Exit Function
            If zlCommFun.StrIsValid(Trim(txt(txt_�¿�ȷ������).Text), 20, txt(txt_�¿�ȷ������).hWnd, "ȷ������") = False Then Exit Function
        End If
    End Select
    IsValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPassword(ByVal txtPass As TextBox, Optional ByVal txtVaild As TextBox, _
    Optional blnOnlyCheckOld As Boolean) As Boolean
    '������Ч�Լ��
    'blnOnlyCheckOld �Ƿ�ֻ���ԭ����
    
    On Error GoTo ErrHandler
    If txtPass.Text = "" Or txtPass.Visible = False Then CheckPassword = True: Exit Function
    Select Case mCardType.int���볤������
    Case 0
    Case 1
        If Len(txtPass.Text) <> mCardType.int���볤�� Then
            ShowMsgbox "�����������" & mCardType.int���볤�� & "λ��"
            zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
            Exit Function
         End If
    Case Else
        If Len(txtPass.Text) <= Abs(mCardType.int���볤������) Then
            ShowMsgbox "�����������" & Abs(mCardType.int���볤������) & "λ���ϣ�"
            zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
            Exit Function
         End If
    End Select
    If mCardType.byt������� = 1 Then '����ֻ����Ϊ����
        If (txtPass.Index = txt_�¿����� Or txtPass.Index = txt_�¿�ȷ������) And txtPass.Text = mobjCard.ԭ���� Then
            '���⴦��
        ElseIf IsNumeric(txtPass.Text) = False Then
            ShowMsgbox "����ֻ�ܰ������֣����������룡"
            zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
            Exit Function
        End If
    End If
    If blnOnlyCheckOld Then CheckPassword = True: Exit Function
    
    If txtPass.Text <> txtVaild.Text Then
        ShowMsgbox "������������벻һ�£����飡"
        zlControl.ControlSetFocus txtPass: zlControl.TxtSelAll txtPass
        Exit Function
    End If
    CheckPassword = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
                
Private Function SaveData() As Boolean
    '��������
    Dim lng������� As Long, lngID As Long
    
    Select Case mEditType
    Case gEd_����
        lng������� = zlDatabase.GetNextId("���ѿ���Ϣ")
        If SavePayCard(lng�������) = False Then Exit Function
        SaveData = True
        
        '��ӡ�ɿ
        If mTy_MoudlePara.bln�ɿ��ӡ Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
                "�������=" & lng�������, "�ɿ�=" & Val(txt(txt_�ɿ�).Text), "�Ҳ�=" & Val(txt(txt_�Ҳ�).Tag), _
                "��ֵID=0", "ReportFormat=1", 2)
        End If
    Case gEd_�޸�
        If SaveModifyCard = False Then Exit Function
    Case gEd_����
        If SaveCallBack = False Then Exit Function
    Case gEd_ȡ������
        If SaveCallBack(True) = False Then Exit Function
    Case gEd_�˿�
        If SaveBackCard(False) = False Then Exit Function
    Case gEd_ȡ���˿�
        If SaveBackCard(True) = False Then Exit Function
    Case gEd_��ֵ
        If SaveInFull(lngID) = False Then Exit Function
        SaveData = True
        
        '��ӡ�ɿ
        If mTy_MoudlePara.bln�ɿ��ӡ Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & (glngSys \ 100) & "_BILL_1503", Me, _
                "��ֵID=" & lngID, "�ɿ�=" & Val(txt(txt_�ɿ�).Text), _
                "�Ҳ�=" & Val(txt(txt_�Ҳ�).Tag), "�������=0", "ReportFormat=2", 2)
        End If
    Case gEd_��ֵ����
        If SaveCancelInFull() = False Then Exit Function
    Case gEd_����
        If SaveChangeCard() = False Then Exit Function
    Case gEd_����
        If SaveReissueCard() = False Then Exit Function
    End Select
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveChangeCard() As Boolean
    '����
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_���ѿ���Ϣ_����
    strSQL = "Zl_���ѿ���Ϣ_����("
    '  ԭ��id_In   ���ѿ���Ϣ.Id%Type,
    strSQL = strSQL & "" & mlng��ID & ","
    '  �¿�����_In ���ѿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_�¿�����).Text) & "',"
    '  ����_In     ���ѿ���Ϣ.����%Type,
    If txt(txt_�¿�����).Text = mobjCard.ԭ���� Then
        strSQL = strSQL & "'" & mobjCard.ԭ���� & "',"
    Else
        strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txt(txt_�¿�����).Text) & "',"
    End If
    '  ������_In   ���ѿ���Ϣ.������%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In ���ѿ���Ϣ.����ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ����id_In   ���ѿ���Ϣ.����id%Type := Null
    strSQL = strSQL & "" & mCardType.lng����ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveChangeCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveReissueCard() As Boolean
    '����
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_���ѿ���Ϣ_����
    strSQL = "Zl_���ѿ���Ϣ_����("
    '  ԭ��id_In   ���ѿ���Ϣ.Id%Type,
    strSQL = strSQL & "" & mlng��ID & ","
    '  �¿�����_In ���ѿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_�¿�����).Text) & "',"
    '  ����_In     ���ѿ���Ϣ.����%Type,
    If txt(txt_�¿�����).Text = mobjCard.ԭ���� Then
        strSQL = strSQL & "'" & mobjCard.ԭ���� & "',"
    Else
        strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txt(txt_�¿�����).Text) & "',"
    End If
    '  ������_In   ���ѿ���Ϣ.������%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����ʱ��_In ���ѿ���Ϣ.����ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ����id_In   ���ѿ���Ϣ.����id%Type := Null
    strSQL = strSQL & "" & mCardType.lng����ID & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveReissueCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveInFull(ByRef lngID As Long) As Boolean
    '�����ֵ����
    '����:lngID-���ر��εĳ�ֵ��ID
    '����:��ֵ�ɹ�,����True,���򷵻�False
    Dim strSQL As String, blnTrain As Boolean
    Dim lng������� As Long, objCard As Card
    
    Err = 0: On Error GoTo ErrHandler
    lngID = zlDatabase.GetNextId("���˿������¼")
    lng������� = lngID
    'Zl_���˿������¼_��ֵ
    strSQL = "Zl_���˿������¼_��ֵ("
    '  Id_In         ���˿������¼.Id%Type,
    strSQL = strSQL & "" & lngID & ","
    '  ���ѿ�id_In   ���˿������¼.���ѿ�id%Type,
    strSQL = strSQL & "" & mlng��ID & ","
    '  ��ֵ���_In   ���˿������¼.Ӧ�ս��%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_���γ�ֵ).Text), 4) & ","
    '  ��ֵ�ۿ�_In   ���˿������¼.����%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_��ֵ����).Text), 4) & ","
    '  �ɿ���_In   ���˿������¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_��ֵ�ɿ�).Text), 4) & ","
    '  ��ֵʱ��_In   ���˿������¼.����ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ����Ա���_In ���˿������¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ���˿������¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �ɿ���_In     ���˿������¼.�ɿ�������%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_�ɿ���).Text) & "',"
    '  ��ֵ˵��_In   ���˿������¼.��ע%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_��ֵ˵��).Text) & "',"
    '  ���㷽ʽ_In     ���˿������¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mCurCardPay.str���㷽ʽ & "',"
    '  �������_In   ���˿������¼.Id%Type,
    strSQL = strSQL & "" & lng������� & ","
    '  ������_In       ���˿������¼.��λ������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_������).Text) & "',"
    '  �ʺ�_In         ���˿������¼.��λ�ʺ�%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�ʺ�).Text) & "',"
    '  �������_In     ���˿������¼.�������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�������).Text) & "',"
    '  �����id_In   ���˿������¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng�����ID = 0, "NULL", mCurCardPay.lng�����ID) & ","
    '  ���㿨��_In   ���˿������¼.���㿨��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.strˢ������ & "'") & ","
    '  ������ˮ��_In ���˿������¼.������ˮ��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str������ˮ�� & "'") & ","
    '  ����˵��_In   ���˿������¼.����˵��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str����˵�� & "'") & ","
    '  �ɿ�_In         ���˿������¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�ɿ�).Text), 4), "NULL") & ","
    '  �Ҳ�_In         ���˿������¼.�Ҳ�%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�Ҳ�).Tag), 4), "NULL") & ")"
    
    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '����������
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.�ӿ���� > 0 Then
        If ExecuteThreeSwapPay(objCard, lng�������, mdblʵ�պϼ�) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveInFull = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCancelInFull() As Boolean
    '��ֵ����
    Dim strSQL As String, blnTrain As Boolean
    Dim lng������� As Long, objCard As Card
    
    Err = 0: On Error GoTo Errhand
    lng������� = zlDatabase.GetNextId("���˿������¼")
    'Zl_���˿������¼_��ֵ����
    strSQL = "Zl_���˿������¼_��ֵ����("
    '  Id_In         ���˿������¼.Id%Type,
    strSQL = strSQL & "" & mlng��ֵID & ","
    '  ����Ա���_In ���˿������¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ���˿������¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ���㷽ʽ_In     ���˿������¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mCurCardPay.str���㷽ʽ & "',"
    '  �������_In   ���˿������¼.Id%Type,
    strSQL = strSQL & "" & lng������� & ","
    '  �����_In   ���˿������¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & -1 * mdbl������� & ","
    '  �������_In     ���˿������¼.�������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�������).Text) & "',"
    '  ������_In       ���˿������¼.��λ������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_������).Text) & "',"
    '  �ʺ�_In         ���˿������¼.��λ�ʺ�%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�ʺ�).Text) & "',"
    '  �����id_In   ���˿������¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng�����ID = 0, "NULL", mCurCardPay.lng�����ID) & ","
    '  ���㿨��_In   ���˿������¼.���㿨��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.strˢ������ & "'") & ","
    '  ������ˮ��_In ���˿������¼.������ˮ��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str������ˮ�� & "'") & ","
    '  ����˵��_In   ���˿������¼.����˵��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str����˵�� & "'") & ","
    '  �ɿ�_In         ���˿������¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�ɿ�).Text), 4), "NULL") & ","
    '  �Ҳ�_In         ���˿������¼.�Ҳ�%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�Ҳ�).Tag), 4), "NULL") & ")"
    
    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '����������
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.�ӿ���� > 0 Then
        If ExecuteThreeSwapPay(objCard, lng�������, mdblʵ�պϼ�) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveCancelInFull = True
    Exit Function
Errhand:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckUsedBill(ByVal lng����� As Long, ByVal lng����ID As Long, _
    Optional ByVal strBill As String) As Long
    '���ܣ���鵱ǰ����Ա�Ƿ��п������ѿ�����(���û���),�����ؿ��õ�����ID
    '������
    '      lng�����=���ѿ��ӿڱ��
    '      lng����ID=��һ�μ��ʱΪ�������õĹ�������ID,�Ժ�Ϊ�ϴ�ʹ�õ�����ID
    '      strBill=Ҫ��鷶Χ��Ʊ�ݺ�
    '˵����
    '    1.�ڼ�鷶Χʱ,��������ж�������Ʊ��,��ֻҪ������һ��֮�о�����
    '    2.�ڼ�鷶Χʱ,����Ҳ�ڼ�鷶Χ֮�ڡ�
    '    3.���ж�������ʱ,ȱʡ���ٵ�����,��������,"���ʹ�õ�����"ԭ��
    '���أ�
    '      ������Ʊ������ID>0
    '      0=ʧ��
    '      -1:û������(�����δ����)��Ҳû�й���(δ����)
    '      -2:���õĹ���������
    '      -3:ָ��Ʊ�ݺŲ��ڵ�ǰ���÷�Χ��(������������Ʊ�ݵ����)
    Dim rsTmp As ADODB.Recordset
    Dim rsSelf As ADODB.Recordset
    Dim strSQL As String, blnTmp As Boolean, lngReturn As Long
    
    On Error GoTo errH
    '����Ա��ʣ�������Ʊ�ݼ�
    ' And ʣ������ > 0  ���ѿ������ظ�ʹ��
    strSQL = _
        "Select ID, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʣ������, �Ǽ�ʱ��, ʹ��ʱ��" & vbNewLine & _
        "From ���ѿ����ü�¼" & vbNewLine & _
        "Where �ӿڱ�� = [1] And ʹ�÷�ʽ = 1 And ������ = [2]" & vbNewLine & _
        "Order By Nvl(ʹ��ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) Desc, ��ʼ����"
    Set rsSelf = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�����, UserInfo.����)
    If lng����ID = 0 Then
        '�����е�һ�μ��,��û�����ñ��ع���
        If rsSelf.EOF Then CheckUsedBill = -1: Exit Function 'Ҳû������Ʊ��
        '������Ʊ��,������ԭ�򷵻�
        lngReturn = Val(NVL(rsSelf!id))
    Else
        '�ϴ�ʹ�õ�����ID���һ�μ��Ĺ���ID,���ж�����
        strSQL = _
            "Select ID, ʹ�÷�ʽ, ʣ������, ǰ׺�ı�, ��ʼ����, ��ֹ����" & vbNewLine & _
            "From ���ѿ����ü�¼" & vbNewLine & _
            "Where �ӿڱ�� = [1] And ID = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�����, lng����ID)
        If rsTmp.EOF Then CheckUsedBill = -2: Exit Function
        
        If Val(NVL(rsTmp!ʹ�÷�ʽ)) = 2 Then '����,Ҫ�ȿ���û������
            If Not rsSelf.EOF Then
                '�����õģ�����
                lngReturn = Val(NVL(rsSelf!id))
            Else
                'û������ȡ����
                If Val(NVL(rsTmp!ʣ������)) = 0 Then CheckUsedBill = -2: Exit Function '�����Ѿ�����
                lngReturn = Val(NVL(rsTmp!id))
                blnTmp = True
            End If
        Else
            '����Ʊ��
            If Val(NVL(rsTmp!ʣ������)) > 0 Then
                '��ʣ��
                lngReturn = Val(NVL(rsTmp!id))
            Else
                '������ʣ�������
                If rsSelf.EOF Then CheckUsedBill = -1: Exit Function '��������Ҳû��ʣ��
                lngReturn = Val(NVL(rsSelf!id))
            End If
        End If
    End If
    
    '���Ʊ�ŷ�Χ�Ƿ���ȷ
    If strBill <> "" Then
        If blnTmp Then
            '�ڹ��÷�Χ�ڷ�Χ�ж�
            If Left(strBill, Len(NVL(rsTmp!ǰ׺�ı�))) <> NVL(rsTmp!ǰ׺�ı�) Then
                lngReturn = -3
            ElseIf Not (strBill >= NVL(rsTmp!��ʼ����) And strBill <= NVL(rsTmp!��ֹ����) _
                And Len(strBill) = Len(NVL(rsTmp!��ʼ����))) Then
                lngReturn = -3
            End If
        Else
            '�ڿ������÷�Χ���ж�
            blnTmp = False
            rsSelf.Filter = "ID=" & lngReturn
            If Left(strBill, Len(NVL(rsSelf!ǰ׺�ı�))) <> NVL(rsSelf!ǰ׺�ı�) Then
                blnTmp = True
            ElseIf Not (strBill >= NVL(rsSelf!��ʼ����) And strBill <= NVL(rsSelf!��ֹ����) _
                And Len(strBill) = Len(NVL(rsSelf!��ʼ����))) Then
                blnTmp = True
            End If
            If blnTmp Then
                '����������,�������������м��
                lngReturn = -3
                rsSelf.Filter = "ID<>" & lngReturn
                Do While Not rsSelf.EOF
                    blnTmp = False
                    If Left(strBill, Len(NVL(rsSelf!ǰ׺�ı�))) <> NVL(rsSelf!ǰ׺�ı�) Then
                        blnTmp = True
                    ElseIf Not (strBill >= NVL(rsSelf!��ʼ����) And strBill <= NVL(rsSelf!��ֹ����) _
                        And Len(strBill) = Len(NVL(rsSelf!��ʼ����))) Then
                        blnTmp = True
                    End If
                    If Not blnTmp Then lngReturn = rsSelf!id: Exit Do
                    rsSelf.MoveNext
                Loop
            End If
        End If
    End If
    CheckUsedBill = lngReturn
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    CheckUsedBill = 0
End Function

Private Function Check��������(Optional ByVal strCardNo As String) As Boolean
    '����:����ϸ���ƿ����Ƿ�����Ч����������
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If Not mCardType.bln�ϸ���� Then Check�������� = True: Exit Function
    
    mCardType.lng����ID = CheckUsedBill(mlng�����, _
        IIf(mCardType.lng����ID > 0, mCardType.lng����ID, mCardType.lng��������), strCardNo)
    If mCardType.lng����ID <= 0 Then
        Select Case mCardType.lng����ID
            Case 0 '����ʧ��
            Case -1
                If strCardNo <> "" Then ShowMsgbox "����û�����ü����õ�" & mCardType.str������ & "�����ܷ��ţ�" & vbCrLf & _
                    "�����ڱ������ù������λ�����һ���¿�! "
                Exit Function
            Case -2
                If strCardNo <> "" Then ShowMsgbox "���ع��õ�" & mCardType.str������ & "�����꣬���ܷ��ţ�" & vbCrLf & _
                    "���������ñ��ع��ÿ����λ�����һ���¿���"
                Exit Function
            Case -3
                ShowMsgbox "���ſ�Ƭ" & IIf(strCardNo = "", "", "��" & strCardNo & "��") & "������Ч��Χ�ڣ������Ƿ���ȷˢ����"
                Exit Function
        End Select
    End If
    
    '����Ƿ�Ҳ������
    strSQL = _
        "Select 1 From ���ѿ�ʹ�ü�¼" & vbNewLine & _
        "Where �ӿڱ�� = [1] And ����id = [2] And ���� = 1 And ԭ�� = 5 And ���� = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�����, mCardType.lng����ID, strCardNo)
    If rsTemp.EOF = False Then
        ShowMsgbox "���ſ���" & IIf(strCardNo = "", "", "��" & strCardNo & "��") & "�ѱ����𣬲�����ʹ�ã�"
        Exit Function
    End If
    
    Check�������� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckInput����(ByVal blnSaveData As Boolean, _
    Optional ByRef lng������ As Long, Optional ByRef strCardNos As String, _
    Optional ByRef lng��ID As Long) As Boolean
    '����:�������Ŀ����Ƿ�Ϸ�
    '���:
    '   blnSaveData - �Ƿ񱣴�����ǰ�ļ��
    '����:
    '   lng������ - ���η�����Χ�п�����
    '   strCardNos - �ֽ�Ŀ��ţ�����ö��š�,���ָ�
    '   lng��ID - ���տ�ʱ���ѿ�ID
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strTable As String, varPara() As Variant
    Dim i As Long, j As Long, k As Long
    Dim strCardNoTemp As String, strCardNo As String
    Dim strCardNoStart As String, strCardNoEnd As String, strInCardNos As String
    Dim objListItem As ListItem, strInfo As String
    Dim varData As Variant, varListData As Variant
    Dim strFindNo As String
    Dim strTemp As String
    
    On Error GoTo ErrHandler
    lng������ = 0
    If mEditType = gEd_���� Then
        strCardNoStart = Trim(txt(txt_��ʼ����).Text)
        strCardNoEnd = Trim(txt(txt_��������).Text)
        If strCardNoStart = "" And blnSaveData = False Then
            ShowMsgbox "��ˢ�������뿨�ţ�": GoTo DataInvalid
        End If
        
        If strCardNoStart <> "" And Len(strCardNoStart) <> mCardType.lng���ų��� Then
            ShowMsgbox "���ų���ӦΪ" & mCardType.lng���ų��� & "λ�����飡": GoTo DataInvalid
        End If
        
        If strCardNoEnd <> "" Then
            If Len(strCardNoEnd) <> mCardType.lng���ų��� Then
                ShowMsgbox "���ų���ӦΪ" & mCardType.lng���ų��� & "λ�����飡"
                zlControl.ControlSetFocus txt(txt_��������)
                Exit Function
            End If
            If strCardNoEnd <= strCardNoStart Then
                ShowMsgbox "�������ű�����ڿ�ʼ���ţ����飡"
                zlControl.ControlSetFocus txt(txt_��������)
                Exit Function
            End If
            
            If Check��������(strCardNoStart) = False Then Exit Function
            If Check��������(strCardNoEnd) = False Then Exit Function
            
            If SplitCardNos(strCardNoStart & "��" & strCardNoEnd, strInCardNos) = False Then Exit Function
        Else
            If strCardNoStart <> "" Then
                If Check��������(strCardNoStart) = False Then Exit Function
                strInCardNos = strCardNoStart
            End If
        End If
        
        '���ÿһ�ſ��Ž��м��
        varData = Split(strInCardNos, ",")
        lng������ = UBound(varData) + 1
        For k = 0 To UBound(varData)
            strCardNo = varData(k)
            If FindDataInGrid(strCardNo) Then
                ShowMsgbox "����Ϊ��" & strCardNo & " ��" & mCardType.str������ & "�Ѵ����ڷ����б��У�": GoTo DataInvalid
            End If
        Next
        
        If strInCardNos <> "" Then strCardNos = strCardNos & "," & strInCardNos
        If blnSaveData Then
            If CheckCardsInGrid() = False Then Exit Function
            strTemp = GetCardsFromGrid()
            If strTemp <> "" Then strCardNos = strCardNos & "," & strTemp
        End If
        If strCardNos <> "" Then strCardNos = Mid(strCardNos, 2)
        If strCardNos = "" Then
            ShowMsgbox "��ˢ�������뿨�ţ�": GoTo DataInvalid
        End If
        
        varPara = Array(mlng�����, mCardType.lng����ID)
        If FromStringListBulidSQL(0, strCardNos, varPara, strTable, "����", 3) = False Then Exit Function
        strSQL = _
            "Select a.ID, a.������, a.�ɷ��ֵ, b.����, a.���," & vbNewLine & _
            "       (Select Max(���) From ���ѿ���Ϣ B Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��) As ������," & vbNewLine & _
            "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
            "       To_Char(a.ͣ������, 'yyyy-mm-dd hh24:mi:ss') As ͣ������" & vbNewLine & _
            "From ���ѿ���Ϣ A, (" & strTable & ") B" & vbNewLine & _
            "Where a.���� = b.���� And a.�ӿڱ��(+) = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
        Do While Not rsTemp.EOF
            If Val(NVL(rsTemp!id)) <> 0 Then
                If NVL(rsTemp!����ʱ��, "3000-01-01") >= "3000-01-01" Then
                    ShowMsgbox "����Ϊ:" & NVL(rsTemp!����) & " ��" & mCardType.str������ & "����ʹ�ã������ٷ�����"
                    strFindNo = NVL(rsTemp!����): GoTo DataInvalid
                End If
                If NVL(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                    ShowMsgbox "����Ϊ:" & NVL(rsTemp!����) & " ��" & mCardType.str������ & "�Ѿ�ֹͣʹ�ã������ٷ�����"
                    strFindNo = NVL(rsTemp!����): GoTo DataInvalid
                End If
            End If
            rsTemp.MoveNext
        Loop
        
        '�������Ͽ�Ƭʱ����Ƿ��ѱ�����һ�ż����ŵ��Ѽ��
        If mCardType.lng����ID > 0 And UBound(Split(strCardNos, ",")) > 1 Then
            strTemp = ""
            strSQL = _
                "Select Distinct a.����" & vbNewLine & _
                "From ���ѿ�ʹ�ü�¼ A, (" & strTable & ") B" & vbNewLine & _
                "Where a.���� = b.���� And a.�ӿڱ�� = [1] And a.����id = [2] And a.���� = 1 And a.ԭ�� = 5"
            Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
            Do While Not rsTemp.EOF
                strTemp = strTemp & "," & NVL(rsTemp!����)
                rsTemp.MoveNext
            Loop
            If strTemp <> "" Then
                strTemp = Mid(strTemp, 2)
                ShowMsgbox "���¿�Ƭ�ѱ����𣬲�����ʹ�ã�" & vbCrLf & strTemp
                Exit Function
            End If
        End If
        CheckInput���� = True
        Exit Function
    End If
    
    If mEditType = gEd_���� Then
        strCardNo = Trim(txt(txt_��ʼ����).Text)
        If strCardNo = "" And blnSaveData = False Then
            ShowMsgbox "��ˢ�������뿨�ţ�": GoTo DataInvalid
        End If
        If FindDataInGrid(strCardNo) Then
            ShowMsgbox "����Ϊ��" & strCardNo & " ��" & mCardType.str������ & "�Ѵ����ڻ����б��У�": GoTo DataInvalid
        End If
        
        If strCardNo <> "" Then
            lng������ = 1
            strCardNos = strCardNos & "," & strCardNo
        End If
        If blnSaveData Then
            strTemp = GetCardsFromGrid()
            If strTemp <> "" Then strCardNos = strCardNos & "," & strTemp
        End If
        If strCardNos <> "" Then strCardNos = Mid(strCardNos, 2)
        If strCardNos = "" Then
            ShowMsgbox "��ˢ�������뿨�ţ�": GoTo DataInvalid
        End If
         
        varPara = Array(mlng�����)
        If FromStringListBulidSQL(0, strCardNos, varPara, strTable, "����", 2) = False Then Exit Function
        strSQL = _
            "Select a.ID, a.������, a.�ɷ��ֵ, b.����, a.���," & vbNewLine & _
            "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
            "       To_Char(a.ͣ������, 'yyyy-mm-dd hh24:mi:ss') As ͣ������" & vbNewLine & _
            "From ���ѿ���Ϣ A, (" & strTable & ") B" & vbNewLine & _
            "Where a.����(+) = b.���� And a.�ӿڱ��(+) = [1]" & vbNewLine & _
            "      And (a.��� Is Null Or a.��� = (Select Max(���) From ���ѿ���Ϣ Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��))"
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, Me.Caption, varPara)
        If rsTemp.EOF Then
            ShowMsgbox "δ���ҵ����������Ϣ�������Ѿ�������ɾ�������ܻ��գ�": GoTo DataInvalid
        Else
            Do While Not rsTemp.EOF
                strFindNo = NVL(rsTemp!����)
                If Val(NVL(rsTemp!id)) = 0 Then
                    '��������
                    ShowMsgbox mCardType.str������ & "(����Ϊ:" & NVL(rsTemp!����) & ")�����Ѿ�������ɾ�������ܻ��գ�": GoTo DataInvalid
                Else
                    If NVL(rsTemp!����ʱ��, "3000-01-01") < "3000-01-01" Then
                        ShowMsgbox "����Ϊ:" & NVL(rsTemp!����) & " ��" & mCardType.str������ & "�ѱ����գ������ٻ��գ�": GoTo DataInvalid
                    End If
                    'ͣ�õ�Ҳ���Ի���
                    'If Nvl(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                    '    ShowMsgbox "����Ϊ:" & Nvl(rsTemp!����) & " ��" & mCardType.str������ & "�Ѿ�ֹͣʹ�ã����ܻ��գ�": GoTo DataInvalid
                    'End If
                End If
                lng��ID = Val(NVL(rsTemp!id))
                rsTemp.MoveNext
            Loop
        End If
        CheckInput���� = True
        Exit Function
    End If
    
    If CardIsValid(mlng��ID) = False Then GoTo DataInvalid
    
    If mEditType = gEd_���� Or mEditType = gEd_���� Then
        If Trim(txt(txt_�¿�����).Text) = "" Then
            ShowMsgbox "��ˢ���������¿����ţ�"
            zlControl.ControlSetFocus txt(txt_�¿�����)
            Exit Function
        End If
        
        If Check��������(txt(txt_�¿�����).Text) = False Then
            zlControl.ControlSetFocus txt(txt_�¿�����)
            zlControl.TxtSelAll txt(txt_�¿�����)
            Exit Function
        End If
        
        strSQL = _
            "Select a.ID, a.������, a.�ɷ��ֵ, a.����, a.���," & vbNewLine & _
            "       (Select Max(���) From ���ѿ���Ϣ B Where ���� = a.���� And �ӿڱ�� = a.�ӿڱ��) As ������," & vbNewLine & _
            "       To_Char(a.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��," & vbNewLine & _
            "       To_Char(a.ͣ������, 'yyyy-mm-dd hh24:mi:ss') As ͣ������" & vbNewLine & _
            "From ���ѿ���Ϣ A" & vbNewLine & _
            "Where a.�ӿڱ�� = [1] And a.���� = [2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�����, txt(txt_�¿�����))
        If Not rsTemp.EOF Then
            If NVL(rsTemp!����ʱ��, "3000-01-01") >= "3000-01-01" Then
                ShowMsgbox "����Ϊ:" & NVL(rsTemp!����) & " ��" & mCardType.str������ & "����ʹ�ã������ٷ�����"
                zlControl.ControlSetFocus txt(txt_�¿�����)
                zlControl.TxtSelAll txt(txt_�¿�����)
                Exit Function
            End If
            If NVL(rsTemp!ͣ������, "3000-01-01") < "3000-01-01" Then
                ShowMsgbox "����Ϊ:" & NVL(rsTemp!����) & " ��" & mCardType.str������ & "�Ѿ�ֹͣʹ�ã������ٷ�����"
                zlControl.ControlSetFocus txt(txt_�¿�����)
                zlControl.TxtSelAll txt(txt_�¿�����)
                Exit Function
            End If
        End If
    End If
    
    CheckInput���� = True
    Exit Function
DataInvalid:
    If (mEditType = gEd_���� Or mEditType = gEd_����) And strFindNo <> "" Then
        If FindDataInGrid(strFindNo) Then
            zlControl.ControlSetFocus vsfCardNo
        Else
            zlControl.ControlSetFocus txt(txt_��ʼ����)
            zlControl.TxtSelAll txt(txt_��ʼ����)
        End If
    ElseIf mEditType = gEd_���� Then
        zlControl.ControlSetFocus txt(txt_ԭ������)
        zlControl.TxtSelAll txt(txt_ԭ������)
    ElseIf mEditType = gEd_���� Then
        zlControl.ControlSetFocus cboԭ������
    Else
        zlControl.ControlSetFocus txt(txt_��ʼ����)
        zlControl.TxtSelAll txt(txt_��ʼ����)
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCardsFromGrid(Optional ByVal blnGetID As Boolean) As String
    '�ӱ���л�ȡ����
    '��Σ�
    '   blnGetID �Ƿ��ȡ���ѿ�ID,�����ȡ����
    Dim i As Long, j As Long
    Dim varData As Variant
    Dim strCards As String
    
    On Error GoTo ErrHandler
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            If vsfCardNo.TextMatrix(i, j) <> "" Then
                varData = vsfCardNo.Cell(flexcpData, i, j) 'Array(������,�ֽ⿨��,���ѿ�ID)
                strCards = strCards & "," & IIf(blnGetID, varData(2), varData(1))
            End If
        Next
    Next
    If strCards <> "" Then strCards = Mid(strCards, 2)
    GetCardsFromGrid = strCards
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckCardsInGrid() As Boolean
    '������еĿ�����Ч��
    Dim i As Long, j As Long
    Dim strCardNoStart As String, strCardNoEnd As String
    
    On Error GoTo ErrHandler
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            If vsfCardNo.TextMatrix(i, j) <> "" Then
                strCardNoStart = Split(vsfCardNo.TextMatrix(i, j) & "��", "��")(0)
                strCardNoEnd = Split(vsfCardNo.TextMatrix(i, j) & "��", "��")(1)
                
                If Check��������(strCardNoStart) = False Then Exit Function
                If strCardNoEnd <> "" Then
                    If Check��������(strCardNoEnd) = False Then Exit Function
                End If
            End If
        Next
    Next
    CheckCardsInGrid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCardsCount() As Long
    '�ӿ�������
    Dim i As Long, j As Long
    Dim varData As Variant
    Dim lngCount As Long
    
    On Error GoTo ErrHandler
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            If vsfCardNo.TextMatrix(i, j) <> "" Then
                varData = vsfCardNo.Cell(flexcpData, i, j) 'Array(������,�ֽ⿨��,���ѿ�ID)
                lngCount = lngCount + Val(varData(0))
            End If
        Next
    Next
    '����δ�������е�����
    lngCount = lngCount + Val(txt(txt_��ʼ����).Tag)
    
    GetCardsCount = lngCount
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function Check�ɿ����() As Boolean
    '����:���ɿ����
    Dim objCard As Card
    Dim strTitle As String, lng��Ƭ���� As Long
    Dim blnYes As Boolean
    
    On Error GoTo ErrHandler
    If mdblʵ�պϼ� < 0 Then
        strTitle = "�˿�"
    Else
        strTitle = "�տ�"
    End If
    
    If GetCurCard(objCard) = False Then
        ShowMsgbox "��ǰ" & strTitle & "��ʽδѡ�����飡"
        zlControl.ControlSetFocus cbo֧����ʽ
        Exit Function
    End If
    
    If zlDblIsValid(Trim(txt(txt_�ɿ�).Text), 16, True, False, txt(txt_�ɿ�).hWnd, strTitle) = False Then Exit Function
    
    If objCard.�������� = 1 Then
        If Val(txt(txt_�ɿ�).Text) = 0 And RoundEx(mdblʵ�պϼ� - mdbl�������, 6) <> 0 Then
            ShowMsgbox "�㻹δ����" & strTitle & "���Ƿ������", True, blnYes
            If blnYes = False Then
                zlControl.ControlSetFocus txt(txt_�ɿ�)
                Exit Function
            End If
        End If
    Else
        If RoundEx(mdblʵ�պϼ�, 6) = 0 Then
            ShowMsgbox "��ǰ" & strTitle & "���Ϊ�㣬����ʹ�÷��ֽ���㷽ʽ��"
            zlControl.ControlSetFocus cbo֧����ʽ
            Exit Function
        End If
        
        If Val(txt(txt_�ɿ�).Text) = 0 Then
            ShowMsgbox "δ����" & strTitle & "�����飡"
            zlControl.ControlSetFocus txt(txt_�ɿ�)
            Exit Function
        End If
    End If
    If Val(txt(txt_�ɿ�).Text) <> 0 Then
        If Val(txt(txt_�ɿ�).Text) < RoundEx(Abs(mdblʵ�պϼ� - mdbl�������), 6) Then
            ShowMsgbox strTitle & "���(" & Format(Val(txt(txt_�ɿ�).Text), "0.00") & ")���㱾��δ�����(" & _
                FormatEx(Abs(mdblʵ�պϼ� - mdbl�������), 6, , , 2) & ")�����飡"
            zlControl.ControlSetFocus txt(txt_�ɿ�)
            Exit Function
        End If
        
        If objCard.�������� <> 1 And Val(txt(txt_�ɿ�).Text) > Val(Format(Abs(mdblʵ�պϼ� - mdbl�������), "0.00")) Then
            ShowMsgbox strTitle & "���(" & Format(Val(txt(txt_�ɿ�).Text), "0.00") & ")�����˱���δ�����(" & _
                FormatEx(Abs(mdblʵ�պϼ� - mdbl�������), 6, , , 2) & ")�����飡"
            zlControl.ControlSetFocus txt(txt_�ɿ�)
            Exit Function
        End If
    End If
    
    If mEditType = gEd_���� And mCardType.bln�ض����� = False Then
        lng��Ƭ���� = Val(lbl(lbl_������2).Caption)
    Else
        lng��Ƭ���� = 1
    End If
    If CheckThreeSwapIsValied(objCard, mdblʵ�պϼ�, lng��Ƭ����) = False Then
        zlControl.ControlSetFocus cbo֧����ʽ
        Exit Function
    End If
    
    If zlCommFun.StrIsValid(Trim(txt(txt_�ɿ���).Text), 20, txt(txt_�ɿ���).hWnd, "�ɿ���") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_������).Text), 50, txt(txt_������).hWnd, "������") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_�ʺ�).Text), 20, txt(txt_�ʺ�).hWnd, "�ʺ�") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_�������).Text), 30, txt(txt_�������).hWnd, "�������") = False Then Exit Function
    Check�ɿ���� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckInput() As Boolean
    '����:����������Ƿ�Ϸ�
    '����:�Ϸ�����true,���򷵻�False
    On Error GoTo ErrHandler
    If zlCommFun.StrIsValid(Trim(txt(txt_����ԭ��).Text), 50, txt(txt_����ԭ��).hWnd, "����ԭ��") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_��ע).Text), 100, txt(txt_��ע).hWnd, "��ע") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_��ֵ˵��).Text), 100, txt(txt_��ֵ˵��).hWnd, "��ֵ˵��") = False Then Exit Function
    
    If mEditType = gEd_���� Then
        If CheckPassword(txt(txt_ȷ������), txt(txt_����)) = False Then Exit Function
        If zlCommFun.StrIsValid(Trim(txt(txt_����).Text), 20, txt(txt_����).hWnd, "����") = False Then Exit Function
        If zlCommFun.StrIsValid(Trim(txt(txt_ȷ������).Text), 20, txt(txt_ȷ������).hWnd, "ȷ������") = False Then Exit Function
    End If
    If zlCommFun.StrIsValid(Trim(txt(txt_�쿨��).Text), 20, txt(txt_�쿨��).hWnd, "�쿨��") = False Then Exit Function
    If mCardType.bln�ض����� And Val(txt(txt_�쿨��).Tag) = 0 Then
        ShowMsgbox "�쿨����Ч�����������룡ע�⣬�쿨�˱����ǽ������ˡ�"
        zlControl.ControlSetFocus txt(txt_�쿨��)
        Exit Function
    End If
    If Trim(txt(txt_�쿨����).Text) <> "" And Val(txt(txt_�쿨����).Tag) = 0 Then
        ShowMsgbox " ��������쿨�����������飡"
        zlControl.ControlSetFocus txt(txt_�쿨����)
        Exit Function
    End If
    If mEditType = gEd_�޸� Then CheckInput = True: Exit Function
    
    '�����
    If CheckInput����� = False Then Exit Function
    If CheckInputʵ�����۶� = False Then Exit Function
    If CheckInputʵ�ʳ�ֵ�ɿ� = False Then Exit Function
    If CheckInput��ֵ���� = False Then Exit Function
    If CheckInput���γ�ֵ = False Then Exit Function
    CheckInput = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SavePayCard(ByVal lng������� As Long) As Boolean
    '����:���淢����Ϣ
    '����:����ɹ�,����true,���򷵻�False
    Dim strCardNoStart As String, strCardNoEnd As String, strInCardNos As String
    Dim cllPro As New Collection, str����ʱ�� As String
    Dim strCardNos As String, varCardNos As Variant
    Dim blnTrain As Boolean
    Dim i As Long, objCard As Card
    Dim lng������� As Long, lng��¼ID As Long
    Dim lng��Ƭ���� As Long
    
    On Error GoTo ErrHandler
    'ȡ��δ������Ŀ���
    strCardNoStart = Trim(txt(txt_��ʼ����).Text)
    strCardNoEnd = Trim(txt(txt_��������).Text)
    If strCardNoEnd <> "" Then
        If SplitCardNos(strCardNoStart & "��" & strCardNoEnd, strInCardNos) = False Then Exit Function
    Else
        If strCardNoStart <> "" Then strInCardNos = strCardNoStart
    End If
    
    strCardNos = GetCardsFromGrid()
    If strInCardNos <> "" Then strCardNos = strCardNos & IIf(strCardNos = "", "", ",") & strInCardNos
    
    Set cllPro = New Collection
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    lng��¼ID = zlDatabase.GetNextId("���˿������¼")
    lng������� = lng��¼ID '����һ�� ���˿������¼.ID ��Ϊ�������
    If zlCommFun.ActualLen(strCardNos) > 4000 Then
        varCardNos = Split(strCardNos, ",")
        strCardNos = ""
        For i = 0 To UBound(varCardNos)
            If zlCommFun.ActualLen(strCardNos & "," & varCardNos(i)) > 4000 Then
                strCardNos = Mid(strCardNos, 2)
                If AddCardDataSQL(lng�������, strCardNos, str����ʱ��, cllPro, lng�������, lng��¼ID) = False Then Exit Function
                lng��¼ID = 0 '����һ��ID�⣬�����Ķ���������ȥȡ
                strCardNos = ""
            End If
            strCardNos = strCardNos & "," & varCardNos(i)
        Next
        If strCardNos <> "" Then strCardNos = Mid(strCardNos, 2)
    End If
    If strCardNos <> "" Then
        If AddCardDataSQL(lng�������, strCardNos, str����ʱ��, cllPro, lng�������, lng��¼ID) = False Then Exit Function
    End If
    If cllPro.count = 0 Then
        ShowMsgbox " ��û��¼���κη������ţ����飡"
        Exit Function
    End If
    
    blnTrain = True
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    
    '����������
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.�ӿ���� > 0 Then
        If mEditType = gEd_���� And mCardType.bln�ض����� = False Then
            lng��Ƭ���� = Val(lbl(lbl_������2).Caption)
        Else
            lng��Ƭ���� = 1
        End If
        If ExecuteThreeSwapPay(objCard, lng�������, mdblʵ�պϼ�, lng��Ƭ����) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SavePayCard = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AddCardDataSQL(ByVal lng������� As Long, ByVal strCardNos As String, _
    ByVal str����ʱ�� As String, ByRef cllPro As Collection, ByVal lng������� As Long, _
    Optional ByVal lng��¼ID As Long) As Boolean
    '����:��ȡ����SQL���
    '���:lng�������-��Ҫ�Ǳ���һ������ʱ�ķ������,�Ա��ӡ
    '����:
    '����:
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_���ѿ���Ϣ_Insert
    strSQL = "Zl_���ѿ���Ϣ_Insert("
    '  �ӿڱ��_In     ���ѿ���Ϣ.�ӿڱ��%Type,
    strSQL = strSQL & "" & mlng����� & ","
    '  ����_In         Varchar2,--����_In ����ö���,�ָ�
    strSQL = strSQL & "'" & strCardNos & "',"
    '  ������_In       ���ѿ���Ϣ.������%Type,
    strSQL = strSQL & "'" & zlStr.NeedName(cbo������.Text) & "',"
    '  ����_In         ���ѿ���Ϣ.����%Type,
    strSQL = strSQL & "'" & zlCommFun.zlStringEncode(txt(txt_����).Text) & "',"
    '  �������_In     ���ѿ���Ϣ.�������%Type,
    strSQL = strSQL & "'" & Get�������() & "',"
    '  �ɷ��ֵ_In     ���ѿ���Ϣ.�ɷ��ֵ%Type,
    strSQL = strSQL & "" & IIf(chk��ֵ.value = vbChecked, 1, 0) & ","
    '  ��Ч��_In       ���ѿ���Ϣ.��Ч��%Type,
    If IsNull(dtp����Ч����.value) Then
        strSQL = strSQL & "NULL,"
    Else
        strSQL = strSQL & "To_Date('" & Format(dtp����Ч����.value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
    End If
    '  ����ԭ��_In     ���ѿ���Ϣ.����ԭ��%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_����ԭ��).Text) & "',"
    '  ������_In       ���ѿ���Ϣ.������%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �����˱��_In   ���˿������¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����ʱ��_In     ���ѿ���Ϣ.����ʱ��%Type,
    strSQL = strSQL & "To_Date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
    '  �쿨��_In       ���ѿ���Ϣ.�쿨��%Type,
    strSQL = strSQL & "'" & txt(txt_�쿨��).Text & "',"
    '  ����id_In       ���ѿ���Ϣ.����id%Type,
    strSQL = strSQL & "" & IIf(mCardType.bln�ض�����, Val(txt(txt_�쿨��).Tag), "NULL") & ","
    '  �쿨����id_In   ���ѿ���Ϣ.�쿨����id%Type,
    strSQL = strSQL & "" & IIf(txt(txt_�쿨����).Tag = "", "NULL", Val(txt(txt_�쿨����).Tag)) & ","
    '  ��ע_In         ���ѿ���Ϣ.��ע%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_��ע).Text) & "',"
    '  ������_In     ���ѿ���Ϣ.������%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_�����).Text), 4) & ","
    '  ���۽��_In     ���ѿ���Ϣ.���۽��%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_���۶�).Text), 4) & ","
    '  �������_In     ���ѿ���Ϣ.�������%Type,
    strSQL = strSQL & "" & lng������� & ","
    '  ����id_In       ���ѿ���Ϣ.����id%Type,
    strSQL = strSQL & "" & IIf(mCardType.lng����ID = 0, "NULL", mCardType.lng����ID) & ","
    '  ��ֵ�ۿ���_In   ���ѿ���Ϣ.��ֵ�ۿ���%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_��ֵ����).Text) * IIf(chk��ֵ.value = vbChecked, 1, 0), 4) & ","
    '  ��ֵ���_In     ���˿������¼.Ӧ�ս��%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_���γ�ֵ).Text) * IIf(chk��ֵ.value = vbChecked, 1, 0), 4) & ","
    '  ��ֵ�ɿ���_In ���˿������¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & Round(Val(txt(txt_��ֵ�ɿ�).Text) * IIf(chk��ֵ.value = vbChecked, 1, 0), 4) & ","
    '  ��ֵ˵��_In     ���˿������¼.��ע%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_��ֵ˵��).Text) & "',"
    '  �ɿ���_In       ���˿������¼.�ɿ�������%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_�ɿ���).Text) & "',"
    '  ���㷽ʽ_In     ���˿������¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mCurCardPay.str���㷽ʽ & "',"
    '  �������_In     ���˿������¼.�������%Type,
    strSQL = strSQL & "" & lng������� & ","
    '  ��¼id_In       ���˿������¼.Id%Type := Null,
    strSQL = strSQL & "" & IIf(lng��¼ID = 0, "NULL", lng��¼ID) & ","
    '  �������_In     ���˿������¼.�������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�������).Text) & "',"
    '  ������_In       ���˿������¼.��λ������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_������).Text) & "',"
    '  �ʺ�_In         ���˿������¼.��λ�ʺ�%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�ʺ�).Text) & "',"
    '  �����id_In   ���˿������¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng�����ID = 0, "NULL", mCurCardPay.lng�����ID) & ","
    '  ���㿨��_In   ���˿������¼.���㿨��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.strˢ������ & "'") & ","
    '  ������ˮ��_In ���˿������¼.������ˮ��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str������ˮ�� & "'") & ","
    '  ����˵��_In   ���˿������¼.����˵��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str����˵�� & "'") & ","
    '  �ɿ�_In         ���˿������¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�ɿ�).Text), 4), "NULL") & ","
    '  �Ҳ�_In         ���˿������¼.�Ҳ�%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�Ҳ�).Tag), 4), "NULL") & ")"
    AddArray cllPro, strSQL
    AddCardDataSQL = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveModifyCard() As Boolean
    '����:���濨Ƭ�޸���Ϣ
    '����:�޸ĳɹ�,����True,���򷵻�False
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    'Zl_���ѿ���Ϣ_Update
    strSQL = "Zl_���ѿ���Ϣ_Update("
    '  Id_In         ���ѿ���Ϣ.Id%Type,
    strSQL = strSQL & "" & mlng��ID & ","
    '  ������_In     ���ѿ���Ϣ.������%Type,
    strSQL = strSQL & "'" & zlStr.NeedName(cbo������.Text) & "',"
    '  �ɷ��ֵ_In   ���ѿ���Ϣ.�ɷ��ֵ%Type,
    strSQL = strSQL & "" & IIf(chk��ֵ.value = vbChecked, 1, 0) & ","
    '  ��Ч��_In     ���ѿ���Ϣ.��Ч��%Type,
    If IsNull(dtp����Ч����.value) Then
        strSQL = strSQL & "NULL,"
    Else
        strSQL = strSQL & "To_Date('" & Format(dtp����Ч����.value, "yyyy-MM-dd") & "','yyyy-mm-dd'),"
    End If
    '  ����ԭ��_In   ���ѿ���Ϣ.����ԭ��%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_����ԭ��).Text) & "',"
    '  �쿨��_In     ���ѿ���Ϣ.�쿨��%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_�쿨��).Text) & "',"
    '  ����id_In     ���ѿ���Ϣ.����id%Type,
    strSQL = strSQL & "" & IIf(mCardType.bln�ض�����, Val(txt(txt_�쿨��).Tag), "NULL") & ","
    '  �쿨����id_In ���ѿ���Ϣ.�쿨����id%Type,
    strSQL = strSQL & "" & IIf(Val(txt(txt_�쿨����).Tag) = 0, "NULL", Val(txt(txt_�쿨����).Tag)) & ","
    '  ��ע_In       ���ѿ���Ϣ.��ע%Type,
    strSQL = strSQL & "'" & Trim(txt(txt_��ע).Text) & "',"
    '  �������_In     ���ѿ���Ϣ.�������%Type
    strSQL = strSQL & "'" & Get�������() & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    SaveModifyCard = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCallBack(Optional blnCancelCallBack As Boolean = False) As Boolean
    '����:���մ���
    '���:
    '   blnCancelCallBack-�Ƿ�ȡ������
    Dim cllPro As New Collection, strIDs As String, varIDs As Variant
    Dim strSQL As String, blnTrain As Boolean
    Dim strNow As String
    Dim i As Long
    
    On Error GoTo ErrHandler
    strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    If blnCancelCallBack Then
        'Zl_���ѿ���Ϣ_Callback
        strSQL = "Zl_���ѿ���Ϣ_Callback("
        '  Ids_In       varchar2,
        strSQL = strSQL & "" & mlng��ID & ","
        '  ������_In   ���ѿ���Ϣ.������%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����ʱ��_In ���ѿ���Ϣ.����ʱ��%Type,
        strSQL = strSQL & "to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ȡ������_In Number := 0
        strSQL = strSQL & "" & 1 & ")"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        SaveCallBack = True: Exit Function
    End If
    
    '���ܴ����������ղ���,���Ҫ����ID�ļ���
    strIDs = GetCardsFromGrid(True)
    If Trim(txt(txt_��ʼ����).Text) <> "" Then
        'ȡ��δ������Ŀ���
        strIDs = strIDs & IIf(strIDs = "", "", ",") & mlng��ID
    End If
    
    Set cllPro = New Collection
    If zlCommFun.ActualLen(strIDs) > 4000 Then
        varIDs = Split(strIDs, ",")
        strIDs = ""
        For i = 0 To UBound(varIDs)
            If zlCommFun.ActualLen(strIDs & "," & varIDs(i)) > 4000 Then
                strIDs = Mid(strIDs, 2)
                'Zl_���ѿ���Ϣ_Callback
                strSQL = "Zl_���ѿ���Ϣ_Callback("
                '  Ids_In     varchar2,
                strSQL = strSQL & "'" & strIDs & "',"
                '  ������_In   ���ѿ���Ϣ.������%Type,
                strSQL = strSQL & "'" & UserInfo.���� & "',"
                '  ����ʱ��_In ���ѿ���Ϣ.����ʱ��%Type,
                strSQL = strSQL & "to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),"
                '  ȡ������_In Number := 0
                strSQL = strSQL & "" & 0 & ")"
                AddArray cllPro, strSQL
                strIDs = ""
            End If
            strIDs = strIDs & "," & varIDs(i)
        Next
        If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    End If
    If strIDs <> "" Then
        'Zl_���ѿ���Ϣ_Callback
        strSQL = "Zl_���ѿ���Ϣ_Callback("
        '  Ids_In       varchar2,
        strSQL = strSQL & "'" & strIDs & "',"
        '  ������_In   ���ѿ���Ϣ.������%Type,
        strSQL = strSQL & "'" & UserInfo.���� & "',"
        '  ����ʱ��_In ���ѿ���Ϣ.����ʱ��%Type,
        strSQL = strSQL & "to_date('" & strNow & "','yyyy-mm-dd hh24:mi:ss'),"
        '  ȡ������_In Number := 0
        strSQL = strSQL & "" & 0 & ")"
        AddArray cllPro, strSQL
    End If
    If cllPro.count = 0 Then
        ShowMsgbox " ��û��ˢҪ���յ�" & mCardType.str������ & "�����飡"
        Exit Function
    End If
    
    blnTrain = True
    ExecuteProcedureArrAy cllPro, Me.Caption
    blnTrain = False
    SaveCallBack = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveBackCard(Optional blnCancelBackCard As Boolean) As Boolean
    '���ܣ��˿�����
    '���:
    '   blnCancelBackCard-�Ƿ�ȡ���˿�
    Dim strSQL As String, blnTrain As Boolean
    Dim lng������� As Long, objCard As Card
    
    Err = 0: On Error GoTo ErrHandler
    lng������� = zlDatabase.GetNextId("���˿������¼")
    'Zl_���ѿ���Ϣ_Backcard
    strSQL = "Zl_���ѿ���Ϣ_Backcard("
    '  ���ѿ�Id_In         ���ѿ���Ϣ.Id%Type,
    strSQL = strSQL & "" & mlng��ID & ","
    '  ����Ա���_In ���˿������¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ���˿������¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  �˿�ʱ��_In   ���ѿ���Ϣ.����ʱ��%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  ���㷽ʽ_In     ���˿������¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & mCurCardPay.str���㷽ʽ & "',"
    '  �������_In   ���˿������¼.Id%Type,
    strSQL = strSQL & "" & lng������� & ","
    '  �����_In   ���˿������¼.ʵ�ս��%Type,
    strSQL = strSQL & "" & IIf(blnCancelBackCard, 1, -1) * mdbl������� & ","
    '  ȡ���˿�_In   Number := 0,
    strSQL = strSQL & "" & IIf(blnCancelBackCard, 1, 0) & ","
    '  �������_In     ���˿������¼.�������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�������).Text) & "',"
    '  ������_In       ���˿������¼.��λ������%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_������).Text) & "',"
    '  �ʺ�_In         ���˿������¼.��λ�ʺ�%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_�ʺ�).Text) & "',"
    '  �����id_In   ���˿������¼.�����id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng�����ID = 0, "NULL", mCurCardPay.lng�����ID) & ","
    '  ���㿨��_In   ���˿������¼.���㿨��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.strˢ������ & "'") & ","
    '  ������ˮ��_In ���˿������¼.������ˮ��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str������ˮ�� & "'") & ","
    '  ����˵��_In   ���˿������¼.����˵��%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng�����ID = 0, "NULL", "'" & mCurCardPay.str����˵�� & "'") & ","
    '  �ɿ�_In         ���˿������¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�ɿ�).Text), 4), "NULL") & ","
    '  �Ҳ�_In         ���˿������¼.�Ҳ�%Type := Null
    strSQL = strSQL & "" & IIf(mCurCardPay.byt�������� = 1, _
        IIf(mdblʵ�պϼ� < 0, -1, 1) * Round(Val(txt(txt_�Ҳ�).Tag), 4), "NULL") & ")"
    
    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '����������
    If GetCurCard(objCard) = False Then Set objCard = New Card
    If objCard.�ӿ���� > 0 Then
        If ExecuteThreeSwapPay(objCard, lng�������, mdblʵ�պϼ�) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveBackCard = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCardNo_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim varData As Variant, lng��ID As Long
    
    On Error GoTo ErrHandler
    cmdDelete.Enabled = ZL_vsGrid_CurrCellHaveData(vsfCardNo, NewRow, NewCol)
    If mEditType = gEd_���� And NewRow >= 0 And NewCol >= 0 Then
        If vsfCardNo.TextMatrix(NewRow, NewCol) <> "" Then
            varData = vsfCardNo.Cell(flexcpData, NewRow, NewCol)
            lng��ID = varData(2) 'Array(������,�ֽ⿨��,���ѿ�ID)
            If CollExitsValue(mcllCard, "K" & lng��ID) Then
                Call ShowCardInfo(mcllCard("K" & lng��ID), True)
            End If
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function FindDataInGrid(ByVal strCardNo As String, Optional ByVal blnSetFocus As Boolean = True) As Boolean
    '�ڵ�Ԫ��������ݣ�����λ
    '��Σ�
    '   strCardNo - ����
    '   blnSetFocus - �Ƿ�����ѡ��Ԫ��
    Dim i As Long, j As Long
    Dim varData As Variant, strTemp As String
    
    On Error GoTo ErrHandler
    If strCardNo = "" Then Exit Function
    For i = 0 To vsfCardNo.Rows - 1
        For j = 0 To vsfCardNo.Cols - 1
            strTemp = vsfCardNo.TextMatrix(i, j)
            If strTemp <> "" Then
                strTemp = strTemp & "��" & strTemp '�����ǿ��ŷ�Χ��ת��Ϊ���ŷ�Χ
                varData = Split(strTemp, "��")
                If varData(0) <= strCardNo And strCardNo <= varData(1) Then
                    If blnSetFocus Then
                        vsfCardNo.Row = i
                        vsfCardNo.Col = j
                    End If
                    FindDataInGrid = True
                    Exit Function
                End If
            End If
        Next
    Next
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfCardNo_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    If Trim(vsfCardNo.TextMatrix(NewRow, NewCol)) = "" Then Cancel = True
End Sub
 

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, ByVal dblMoney As Double, _
    Optional ByVal lng��Ƭ���� As Long = 1) As Boolean
    '����:������ˢ����֤
    '���:objCard-��ǰ��
    '����:ˢ���ɹ�,����true,���򷵻�False
    Dim strXMLExpend As String, strBalanceIDs As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If objCard.�ӿ���� <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If dblMoney = 0 Then CheckThreeSwapIsValied = True: Exit Function
    
    If dblMoney > 0 Then
        'zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl��� As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln�˷� As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln���� As Boolean = False, _
        Optional ByVal bln�����ֹ As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal blnתԤ�� As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:����ָ��֧�����,����ˢ������
        '���:rsClassMoney:�շ����,���
        '        lngCardTypeID-Ϊ��ʱ,Ϊ��һ��ͨˢ��
        '       bln�����ֹ-Ŀǰֻ������ѿ�,��ʾ����ʱ,��ֹ��������,������������֧��
        '       dblBrushTotaled-������Ч,��ʾ�Ѿ�ˢ���ѿ��ܶ�(��Ҫ���ڶ��ˢ��)
        '       str�ϴ��������-�ϴ�ˢ����ʱ���������(ͬ�ζ��ˢ���ѿ�ʱ,��Ҫ��鱾��ˢ��������ϴ�����Ƿ�һ��,��һ�²�����ˢ������)
        '       varSquareBalance- Collection����,��ǰ�Ѿ�ˢ������Ϣ(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ����� ))
        '       blnԤ��-�Ƿ�תԤ��
        '       blnAllPay-�Ƿ����ȫ֧����true-����δ֧���겻����ɽ��㣬false-����ֻ֧�����ֲ�����
        '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
        '       <IN>
        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
        '       </IN>
        '����:str�������-�������(���ѿ�����)
        '        lng��ID-���ѿ���Ϣ.ID(���ѿ�����)
        '       strCardNO-����ˢ���Ŀ���
        '       strPassWord-����ˢ������Ӧ������
        '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
        '����:�ɹ�,����true,���򷵻�False
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
            objCard.�ӿ����, False, "", "", "", dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
            False, False, False, True, Nothing, False, True, "<IN><CZLX>0</CZLX></IN>") = False Then Exit Function
        
        '����ǰ,һЩ���ݼ��
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNos As String, _
        Optional ByVal strXMLExpend As String) As Boolean
        '����:�ʻ��ۿ�׼��
        '���:frmMain-���õ�������
        '       lngModule-���õ�ģ���
        '       strCardTypeID-�����ID
        '       strCardNo-����
        '       dblMoney-֧�����(�˿�ʱΪ����)
        '       strNos-����֧�����漰�ĵ���
        '       strXMLExpend-(XML��:��֤����:��������)
        '        ���ѿ��տ�ʱ������XML����
        '        <IN>
        '            <MZXSJE>��ֵ���۽��</MZXSJE>
        '            <CZJKJE>��ֵ�ɿ���</CZJKJE>
        '        </IN>
        '����:
        '   strXMLExpend-(XML��:������Ϣ)
        '����:�ۿ�Ϸ�,����true,���򷵻�Flase
        strXMLExpend = ""
        strXMLExpend = strXMLExpend & "<IN>"
            strXMLExpend = strXMLExpend & "<MZXSJE>" & IIf(mEditType = gEd_��ֵ, "0", lng��Ƭ���� * Val(txt(txt_���۶�).Text)) & "</MZXSJE>"
            strXMLExpend = strXMLExpend & "<CZJKJE>" & lng��Ƭ���� * Val(txt(txt_��ֵ�ɿ�).Text) & "</CZJKJE>"
        strXMLExpend = strXMLExpend & "</IN>"
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, objCard.�ӿ����, _
            False, mCurCardPay.strˢ������, dblMoney, "", strXMLExpend) = False Then Exit Function
    Else
        mrsBalance.Filter = ""
        If mrsBalance.EOF Then
            ShowMsgbox "δ���ҵ�" & objCard.���� & "ԭ֧��������Ϣ�������˻أ�"
            Exit Function
        End If
        mCurCardPay.lngԭ������� = Val(NVL(mrsBalance!�������))
        
        If objCard.�Ƿ�ת�ʼ����� Then
            '   zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln���ѿ� As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByRef dbl��� As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln�˷� As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln���� As Boolean = False, _
            Optional ByVal bln�����ֹ As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal blnתԤ�� As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
            '       <IN>
            '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
            '       </IN>
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                objCard.�ӿ����, False, "", "", "", dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
                False, False, False, True, Nothing, False, True, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
            
            '����ת�ʽӿ�
            'zlTransferAccountsCheck ת�ʼ��ӿ�
            '������  ��������    ��/��   ��ע
            'frmMain Object  In  ���õ�������
            'lngModule   Long    In  HIS����ģ���
            'lngCardTypeID   Long    In  �����ID
            'strCardNo   String  In  ����
            'dblMoney    Double  In  ת�ʽ��(����ʱΪ����)
            'strBalanceID    String  In  ԭ֧���������,���ò����¼.������Ż���Ԥ����¼.�������
            'strXMLExpend String In   XML��:
            '                            <IN>
            '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��
            '                                       2-����ҵ��;3-�����˷�ҵ��4-�����˷�ҵ��5-���ѿ������˷�ҵ��
            '                            </IN>
            '                    Out  XML��:
            '                            <OUT>
            '                               <ERRMSG>������Ϣ</ERRMSG >
            '                            </OUT>
            '    Boolean ��������    �������ݺϷ�,����True:���򷵻�False
            '˵��:
            '��. ������ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
            '��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
            '����XML��
            strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
            If gobjSquare.objSquareCard.zlTransferAccountsCheck(Me, mlngModule, objCard.�ӿ����, _
                mCurCardPay.strˢ������, -1 * dblMoney, mCurCardPay.lngԭ�������, strXMLExpend) = False Then
                Call ShowThreeSwapErrMsg(0, strXMLExpend)
                Exit Function
            End If
        Else
            mrsBalance.Filter = "�����ID=" & objCard.�ӿ����
            If mrsBalance.EOF Then
                ShowMsgbox "δ���ҵ�" & objCard.���� & "ԭ֧��������Ϣ�������˻أ�"
                Exit Function
            End If
            mCurCardPay.lngԭ������� = Val(NVL(mrsBalance!�������))
            
            If objCard.�Ƿ�ȫ�� Then
                strSQL = "Select Nvl(Sum(ʵ�ս��), 0) As �ɿ�ϼ� From ���˿������¼ Where ������� = [1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(NVL(mrsBalance!�������)))
                If Val(NVL(rsTemp!�ɿ�ϼ�)) <> -1 * dblMoney Then
                    ShowMsgbox objCard.���� & "��֧�ֲ����ˣ���˲����˻أ���ѡ�������˿ʽ��" & _
                        "(ԭ֧����" & FormatEx(Val(NVL(rsTemp!�ɿ�ϼ�)), 6, , , 2) & _
                        "�����˿��" & FormatEx(-1 * dblMoney, 6, , , 2) & ")"
                    Exit Function
                End If
            End If
            'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
                ByVal lngCardTypeID As Long, bln���ѿ� As Boolean, ByVal strCardNo As String, _
                ByVal strBalanceIDs As String, _
                ByVal dblMoney As Double, ByVal strSwapNo As String, _
                ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
                '---------------------------------------------------------------------------------------------------------------------------------------------
                '����:�ʻ����˽���ǰ�ļ��
                '���:frmMain-���õ�������
                '       lngModule-���õ�ģ���
                '       lngCardTypeID-�����ID
                '       strCardNo-����
                '       strBalanceIDs   String  In  ����֧�����漰�Ľ���ID ��ʽ:�շ�����|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
                '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ��տ
                '                                           ����=7ʱ��IDΪ���˿������¼.�������
                '       dblMoney-�˿���
                '       strSwapNo-������ˮ��(�˿�ʱ���)
                '       strSwapMemo-����˵��(�˿�ʱ����)
                '       strXMLExpend    XML IN  ��ѡ����(��չ��):
                '        <TFDATA> //�˷�����
                '          <YCTF>1</YCTF> //�Ƿ��쳣����:1-�쳣����;0-�˷� �˽ڵ����û��
                '          <TFLIST> //�˷��б�
                '            <NO></NO> // �˷ѵ���
                '            <TFITEM> //�˷���
                '              <SerialNum></SerialNum> //���
                '              ��
                '            </TFITEM>
                '          </TFLIST>
                '          ....
                '        </TFDATA >
                '����:�˿�Ϸ�,����true,���򷵻�Flase
            mCurCardPay.strˢ������ = NVL(mrsBalance!���㿨��)
            mCurCardPay.str������ˮ�� = NVL(mrsBalance!������ˮ��)
            mCurCardPay.str����˵�� = NVL(mrsBalance!����˵��)
            strBalanceIDs = "7|" & mCurCardPay.lngԭ�������
            If gobjSquare.objSquareCard.zlReturncheck(Me, mlngModule, objCard.�ӿ����, _
                objCard.���ѿ�, mCurCardPay.strˢ������, strBalanceIDs, -1 * dblMoney, _
                mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, strXMLExpend) = False Then Exit Function
        
            If objCard.�Ƿ��˿��鿨 Then
               '����ˢ������
                'zlBrushCard(frmMain As Object, _
                ByVal lngModule As Long, _
                ByVal rsClassMoney As ADODB.Recordset, _
                ByVal lngCardTypeID As Long, _
                ByVal bln���ѿ� As Boolean, _
                ByVal strPatiName As String, ByVal strSex As String, _
                ByVal strOld As String, ByVal dbl��� As Double, _
                Optional ByRef strCardNo As String, _
                Optional ByRef strPassWord As String, _
                Optional ByRef bln�˷� As Boolean = False, _
                Optional ByRef blnShowPatiInfor As Boolean = False, _
                Optional ByRef bln���� As Boolean = False, _
                Optional ByVal bln�����ֹ As Boolean = True, _
                Optional ByRef varSquareBalance As Variant, _
                Optional ByVal blnתԤ�� As Boolean = False, _
                Optional ByVal blnAllPay As Boolean = False, _
                Optional ByVal strXmlIn As String = "") As Boolean
                '       strXmlIn-����������XML���,Ŀǰ��ʽ����:
                '       <IN>
                '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
                '       </IN>
                If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                    objCard.�ӿ����, False, "", "", "", -1 * dblMoney, mCurCardPay.strˢ������, mCurCardPay.strˢ������, _
                    True, False, False, True, Nothing, False, True, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
            End If
        End If
    End If
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapPay(ByVal objCard As Card, ByVal lng������� As Long, _
    ByVal dblMoney As Double, Optional ByVal lng��Ƭ���� As Long = 1) As Boolean
    '����:һ��֧ͨ��(�����ӿ�)
    '���:
    '   objCard-��ǰ��
    '   dblMoney-����֧�����
    '����:
    '����:ִ�гɹ�,����true,���򷵻�False
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSwapExtendInfor As String, strTemp As String
    Dim strXMLExpend As String
    
    On Error GoTo ErrHandler
    If objCard.�ӿ���� <= 0 Then ExecuteThreeSwapPay = True: Exit Function
    If dblMoney = 0 Then ExecuteThreeSwapPay = True: Exit Function
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    If dblMoney > 0 Then
        'zlPaymentMoney(ByVal frmMain As Object, _
        ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
        ByVal bln���ѿ� As Boolean, _
        ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        byval  strPrepayNos as string , _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, _
        ByRef strSwapMemo As String, _
        Optional ByRef strSwapExtendInfor As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '����:�ʻ��ۿ��
        '���:frmMain-���õ�������
        '        lngModule-����ģ���
        '        strBalanceIDs-����ID,����ö��ŷ���;���ѿ��տ�ʱΪ���˿������¼.�������
        '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
        '       strCardNo-����
        '       dblMoney-֧�����
        '       strSwapExtendInfor- ���ѿ��տ�ʱ������XML����
        '                            <IN>
        '                                <MZXSJE>��ֵ���۽��</MZXSJE>
        '                                <CZJKJE>��ֵ�ɿ���</CZJKJE>
        '                            </IN>
        '����:strSwapGlideNO-������ˮ��
        '       strSwapMemo-����˵��
        '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
        '����:�ۿ�ɹ�,����true,���򷵻�Flase
        '˵��:
        '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
        '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
        '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
        '---------------------------------------------------------------------------------------------------------------------------------------------
        strSwapExtendInfor = ""
        strSwapExtendInfor = strSwapExtendInfor & "<IN>"
            strSwapExtendInfor = strSwapExtendInfor & "<MZXSJE>" & IIf(mEditType = gEd_��ֵ, "0", lng��Ƭ���� * Val(txt(txt_���۶�).Text)) & "</MZXSJE>"
            strSwapExtendInfor = strSwapExtendInfor & "<CZJKJE>" & lng��Ƭ���� * Val(txt(txt_��ֵ�ɿ�).Text) & "</CZJKJE>"
        strSwapExtendInfor = strSwapExtendInfor & "</IN>"
        strTemp = strSwapExtendInfor
        If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, _
            mCurCardPay.strˢ������, lng�������, "", dblMoney, _
            mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, strSwapExtendInfor) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
        
        Call zlAddUpdateSwapSQL(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
            mCurCardPay.strˢ������, mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, cllUpdate, 1, 0, 1)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
                mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap, 0, 1)
            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        End If
    Else
        If objCard.�Ƿ�ת�ʼ����� Then
            'zlTransferAccountsMoney
            '������  ��������    ��/��   ��ע
            'frmMain Object  In  ���õ�������
            'lngModule   Long    In  HIS����ģ���
            'lngCardTypeID   Long    In  �����ID
            'strCardNo   String  In  ����
            'strBalanceID    String  In  ����ID ����֧���������,���ò����¼.������Ż���Ԥ����¼.������Ż��˿������¼.�������
            'dblMoney    Double  In  ת�ʽ��
            'strSwapGlideNO  String  Out ������ˮ��
            'strSwapMemo String  Out ����˵��
            'strSwapExtendInfor  String  In �˷�ҵ��ʱ�����뱾���˷ѵĳ���ID:
            '                               ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
            '                               �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ������տ�(IDΪ�������)
            '                           Out ������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
            'strXMLExpend String In   XML��:
            '                            <IN>
            '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��
            '                                       2-����ҵ��;3-�����˷�ҵ��4-�����˷�ҵ��5-���ѿ������˷�ҵ��
            '                            </IN>
            '                    Out  XML��:
            '                            <OUT>
            '                               <ERRMSG>������Ϣ</ERRMSG >
            '                            </OUT>
            '    Boolean ��������    True:���óɹ�,False:����ʧ��
            '˵��:
            '��. ��ҽ���������ʱ���е�����ת��ʱ���á�
            '��. һ����˵���ɹ�ת�ʺ󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
            '��. ��ת�ʳɹ��󣬷��ؽ�����ˮ�ź���ؽ���˵���������������������Ϣ�����Է�����չ��Ϣ�з���.
            '����XML��
            strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
            strSwapExtendInfor = "7|" & mCurCardPay.lngԭ�������: strTemp = strSwapExtendInfor
            If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.�ӿ����, _
                mCurCardPay.strˢ������, lng�������, -1 * dblMoney, _
                mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, strSwapExtendInfor, strXMLExpend) = False Then
                gcnOracle.RollbackTrans: Call ShowThreeSwapErrMsg(1, strXMLExpend)
                Exit Function
            End If
            gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
            
            Call zlAddUpdateSwapSQL(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
                mCurCardPay.strˢ������, mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, cllUpdate, 1, 0, 1)
            zlExecuteProcedureArrAy cllUpdate, Me.Caption
            If strTemp <> strSwapExtendInfor Then
                Call zlAddThreeSwapSQLToCollection(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
                    mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap, 0, 1)
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
            End If
        Else
            'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
                ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
                ByVal dblMoney As Double, _
                ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
                ByRef strSwapExtendInfor As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:�ʻ��ۿ���˽���
            '���:frmMain-���õ�������
            '       lngModule-���õ�ģ���
            '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
            '       strCardNo-����
            '       strBalanceIDs-����֧�����漰�Ľ���ID(����ԭ����ID):
            '                           ��ʽ:�շ�����(|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
            '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ��տ�
            '       dblMoney-�˿���
            '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
            '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
            '       strSwapExtendInfor-���룬�����˷ѵĳ���ID��
            '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
            '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-���ղ������,7-���ѿ��տ�
            '       strSwapExtendInfor-���������׵���չ��Ϣ
            '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
            strSwapExtendInfor = "7|" & lng�������
            If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, _
                mCurCardPay.strˢ������, "7|" & mCurCardPay.lngԭ�������, -1 * dblMoney, _
                mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, strSwapExtendInfor) = False Then
                gcnOracle.RollbackTrans: Exit Function
            End If
            gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
            
            Call zlAddUpdateSwapSQL(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
                mCurCardPay.strˢ������, mCurCardPay.str������ˮ��, mCurCardPay.str����˵��, cllUpdate, 1, 0, 1)
            zlExecuteProcedureArrAy cllUpdate, Me.Caption
            If strTemp <> strSwapExtendInfor Then
                Call zlAddThreeSwapSQLToCollection(False, lng�������, objCard.�ӿ����, objCard.���ѿ�, _
                    mCurCardPay.strˢ������, strSwapExtendInfor, cllThreeSwap, 0, 1)
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
            End If
        End If
    End If
    ExecuteThreeSwapPay = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowThreeSwapErrMsg(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '����:����ת�˼�������ҵ�������ʾ
    '����:
    '   bytType:0-ת�˼��,1-ת�˽���
    '   strXMLErrMsg:��ʽ����
    '            <OUT>
    '               <ERRMSG>������Ϣ</ERRMSG >
    '            </OUT>
    Dim strValue As String
    
    On Error GoTo errHandle
    '����������Ϣ
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '��ʾ������Ϣ
    If Trim(strValue) = "" Then
        If bytType = 0 Then
            strValue = vbCrLf & "���׼��ʧ�ܣ�"
        Else
            strValue = vbCrLf & "����ʧ�ܣ�"
        End If
    End If
    MsgBox strValue, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CheckThreeBalanceToCash(ByVal objCard As Card) As Boolean
    '���������ּ��
    Dim str����Ա As String
    
    On Error GoTo errHandle
    If Not (objCard.�ӿ���� > 0 And Not objCard.���ѿ�) Then CheckThreeBalanceToCash = True: Exit Function
    If objCard.�Ƿ����� Then CheckThreeBalanceToCash = True: Exit Function
    
    If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "�����˿�ǿ������") Then
        If MsgBox(objCard.���� & "��֧�����֣���ȷ��Ҫ����ǿ��������", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        str����Ա = zlDatabase.UserIdentifyByUser(Me, objCard.���� & "ǿ�����֣�Ȩ����֤��", _
            glngSys, mlngModule, "�����˿�ǿ������", , True)
        If str����Ա = "" Then Exit Function
    End If
    CheckThreeBalanceToCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
