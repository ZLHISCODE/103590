VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSquareAffirm 
   Caption         =   "�������ѽ���"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSquareAffirm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8114.99
   ScaleMode       =   0  'User
   ScaleWidth      =   11445
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPatientInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   60
      ScaleHeight     =   1455
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   60
      Width           =   8925
      Begin VB.CommandButton cmdYB 
         Caption         =   "ҽ��"
         Height          =   375
         Left            =   3435
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ���F6"
         Top             =   105
         Width           =   720
      End
      Begin VB.Label lbl 
         Caption         =   "Ԥ�����:99999999.99"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   6
         Top             =   585
         Width           =   4110
      End
      Begin VB.Label lbl 
         Caption         =   "δ�����:99999999.99"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4380
         TabIndex        =   7
         Top             =   585
         Width           =   4410
      End
      Begin VB.Label lbl 
         Caption         =   "��˹��.����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   1455
         TabIndex        =   2
         Top             =   135
         Width           =   2370
      End
      Begin VB.Label lbl 
         Caption         =   "�����:1810080001"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6000
         TabIndex        =   5
         Top             =   135
         Width           =   2760
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "��  ��:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Index           =   0
         Left            =   210
         TabIndex        =   1
         Top             =   135
         Width           =   1260
      End
      Begin VB.Label lbl 
         Caption         =   "�Ա�:����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4380
         TabIndex        =   4
         Top             =   135
         Width           =   1425
      End
      Begin VB.Label lbl 
         Caption         =   "ʣ����:99999999.99"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   60
         TabIndex        =   8
         Top             =   1035
         Width           =   4110
      End
      Begin VB.Label lbl 
         Caption         =   "�������:99999999.99"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   4380
         TabIndex        =   9
         Top             =   1035
         Width           =   4410
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   1470
         X2              =   4140
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   2
         X1              =   5160
         X2              =   5850
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   3
         X1              =   7140
         X2              =   8790
         Y1              =   465
         Y2              =   465
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   4
         X1              =   1470
         X2              =   4140
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   5
         X1              =   1470
         X2              =   4140
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   6
         X1              =   5805
         X2              =   8785
         Y1              =   915
         Y2              =   915
      End
      Begin VB.Line lineUnder 
         BorderColor     =   &H80000000&
         Index           =   7
         X1              =   5805
         X2              =   8785
         Y1              =   1365
         Y2              =   1365
      End
   End
   Begin VB.CommandButton cmdYBBalance 
      Caption         =   "ҽ������(&Y)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9210
      TabIndex        =   28
      ToolTipText     =   "�ȼ���F2"
      Top             =   345
      Width           =   2055
   End
   Begin VB.CommandButton cmdInsureSet 
      Caption         =   "��������(&I)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   31
      Top             =   3270
      Width           =   2055
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "��ӡ����(&P)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   32
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   29
      ToolTipText     =   "�ȼ���F2"
      Top             =   375
      Width           =   2055
   End
   Begin VB.PictureBox pic��� 
      BorderStyle     =   0  'None
      Height          =   810
      Left            =   9210
      ScaleHeight     =   810
      ScaleWidth      =   2040
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2370
      Visible         =   0   'False
      Width           =   2040
      Begin VB.Label lbl 
         Caption         =   "�������"
         Height          =   315
         Index           =   13
         Left            =   105
         TabIndex        =   26
         Top             =   60
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0.0111"
         Height          =   315
         Index           =   14
         Left            =   885
         TabIndex        =   27
         Top             =   420
         Width           =   1080
      End
   End
   Begin VB.PictureBox picʣ���Ը� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   60
      ScaleHeight     =   1365
      ScaleWidth      =   3360
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1590
      Width           =   3390
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Index           =   8
         Left            =   2235
         TabIndex        =   12
         Top             =   585
         Width           =   1005
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTitle 
         Height          =   450
         Left            =   15
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   30
         Width           =   3345
         _Version        =   589884
         _ExtentX        =   5900
         _ExtentY        =   794
         _StockProps     =   6
         Caption         =   "��ǰδ��"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox pic�Ը��ϼ� 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   60
      ScaleHeight     =   1320
      ScaleWidth      =   3360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3030
      Width           =   3390
      Begin XtremeSuiteControls.ShortcutCaption stcTitleTotal 
         Height          =   420
         Left            =   15
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   30
         Width           =   3345
         _Version        =   589884
         _ExtentX        =   5900
         _ExtentY        =   741
         _StockProps     =   6
         Caption         =   "�Ը��ϼ�"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Index           =   9
         Left            =   2220
         TabIndex        =   15
         Top             =   600
         Width           =   1005
      End
   End
   Begin MSCommLib.MSComm mscCom 
      Left            =   11160
      Top             =   6150
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picPayInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   3510
      ScaleHeight     =   2745
      ScaleWidth      =   5445
      TabIndex        =   16
      Top             =   1590
      Width           =   5475
      Begin VB.TextBox txtժҪ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1290
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   1320
         Width           =   3960
      End
      Begin VB.TextBox txt��� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   764
         Width           =   2100
      End
      Begin VB.ComboBox cbo֧����ʽ 
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   772
         Width           =   1395
      End
      Begin VB.TextBox txt��Ԥ�� 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   1290
         MaxLength       =   10
         TabIndex        =   18
         Top             =   210
         Width           =   3960
      End
      Begin zlIDKind.ucQRCodePayButton btQRCodePay 
         Height          =   450
         Left            =   4800
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "ɨ�븶����ʹ�ÿ����F3�����п���֧��"
         Top             =   744
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   794
         Appearance      =   2
         ToolTipString   =   "ɨ�븶����ʹ�ÿ����F3�����п���֧��"
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ժ  Ҫ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   285
         TabIndex        =   23
         Top             =   1365
         Width           =   960
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   285
         TabIndex        =   17
         Top             =   285
         Width           =   945
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   285
         TabIndex        =   19
         Top             =   832
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9210
      TabIndex        =   30
      Top             =   930
      Width           =   2055
   End
   Begin VB.PictureBox picBlanceAndFee 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   60
      ScaleHeight     =   2460
      ScaleWidth      =   11205
      TabIndex        =   33
      Top             =   4440
      Width           =   11235
      Begin VB.PictureBox picFee 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1470
         Left            =   300
         ScaleHeight     =   1470
         ScaleWidth      =   10275
         TabIndex        =   40
         Top             =   780
         Width           =   10275
         Begin VSFlex8Ctl.VSFlexGrid vsFee 
            Height          =   1125
            Left            =   300
            TabIndex        =   41
            Top             =   270
            Width           =   10065
            _cx             =   17754
            _cy             =   1984
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSquareAffirm.frx":0442
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
            AllowUserFreezing=   1
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picBlance 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   30
         ScaleHeight     =   2415
         ScaleWidth      =   10995
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   150
         Width           =   10995
         Begin VSFlex8Ctl.VSFlexGrid vsBalance 
            Height          =   2295
            Left            =   0
            TabIndex        =   39
            Top             =   420
            Width           =   10125
            _cx             =   17859
            _cy             =   4048
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
            BackColorSel    =   16761024
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   350
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSquareAffirm.frx":05A5
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
            Editable        =   2
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
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ҽ��֧��:99999999.99"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   17
            Left            =   6690
            TabIndex        =   38
            Top             =   90
            Width           =   2640
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���úϼ�:99999999.99"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   15
            Left            =   60
            TabIndex        =   36
            Top             =   90
            Width           =   2640
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�Ѹ��ϼ�:99999999.99"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Index           =   16
            Left            =   3240
            TabIndex        =   37
            Top             =   90
            Width           =   2640
         End
      End
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1125
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   2355
         _Version        =   589884
         _ExtentX        =   4154
         _ExtentY        =   1984
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   42
      Top             =   7080
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmSquareAffirm.frx":06BC
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11721
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Key             =   "�����ʻ���ʾ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmSquareAffirm.frx":0F50
            Key             =   "Calc"
            Object.ToolTipText     =   "������:ALT+?"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1693
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSquareAffirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------
'��α���
Private mfrMain As Object
Private mbytBillType As Byte '0-�������շѻ���ʵ�,1-�շѼ�¼;2-���ʼ�¼
Private mlngModule As Long, mlngPatiID As Long
Private mstrInNos As String, mstrҽ��IDs As String, mstrPrivs As String
Private mlngCardTypeID As Long, mbln���ѿ� As Boolean
Private mblnCliniqueRoomPay As Boolean  '���֧��
Private mblnʹ��Ԥ�� As Boolean '�Ƿ�����ʹ��Ԥ����,104381
'---------------------------------------------------------------------
'ģ�����
Private mlng����ID As Long, mblnOk As Boolean
Private mobjPayCards As Cards
Private mrsInfo As ADODB.Recordset
Private mblnFirst As Boolean
Private mstrTittle As String '�������

'---------------------------------------------------------------------
'ģ�����
Private mintFeePrecision  As Integer
Private mbytFeeMoneyPrecision  As Byte
Private Type Ty_Para
    int���Ʊ�ݸ�ʽ As Integer
    int�շ�Ʊ�ݸ�ʽ As Integer
    int��˴�ӡ��ʽ As Integer
    int�շѴ�ӡ��ʽ As Integer
    intҩƷ��λ As Integer
End Type
Private mbytCurType As Byte '1-�����շ�;2-�������
Private mPara As Ty_Para
Private mblnֻ��ҽ������ɹ������շ� As Boolean

Public mbln�����Զ����� As Boolean '���ʻ��۵���˺��Զ�����

'����ֵ
Private Enum Pg_Index
    Blance = 0
    FeeDetail
End Enum

Private Enum Lbl_Index
    ���� = 1
    �Ա� = 2
    ����� = 3
    Ԥ����� = 4
    δ����� = 5
    ʣ���� = 6
    ������� = 7
    ��ǰδ�� = 8
    �Ը��ϼ� = 9
    Ԥ��� = 10
    �ɿ� = 11
    ժҪ = 12
    ��� = 14
    ���úϼ� = 15
    �Ѹ��ϼ� = 16
    ҽ��֧�� = 17
End Enum

Private Enum Pan
    C2��ʾ��Ϣ = 2
    C3�����ʻ� = 3
End Enum
'----------------------------------------------------------------------------
'��������
Private mrs���㷽ʽ As ADODB.Recordset
Private Type TY_ChargeMoney
    dbl���úϼ� As Double
    dbl���γ�Ԥ��  As Double
    dblҽ��֧�� As Double
    dbl�Ѹ��ϼ� As Double
    dbl��ǰδ�� As Double
    dbl�������� As Double
    
    dblԤ����� As Double
    dbl������� As Double
    dbl����Ԥ�� As Double
    
    lng����ID As Long
    lng������� As Long
End Type
Private mCurCharge As TY_ChargeMoney
'------------------------------------------------------------------------------------------
'��֧�����
Private Type TY_PayMoney
    lng�����ID As Long
    bln���ѿ� As Boolean
    str���㷽ʽ As String
    str���� As String
    strˢ������ As String
    strˢ������ As String
    strQRCode As String
    str������ˮ�� As String
    str����˵�� As String
    bln���� As Boolean
    bln��������  As Boolean
    intҽ�ƿ����� As Integer
    bln֧Ʊ As Boolean
    bln���ƿ� As Boolean
    blnOneCard As Boolean '�Ƿ�һ��ͨ����
    int���� As Integer '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����;<0 ��ʾ������֧��
    strNo As String
    lngID As Long 'Ԥ��ID
    objCard As Card
    str֧������ As String '�ַ�������ʽ�����㷽ʽ|֧�����||...
End Type
Private mCurCardPay As TY_PayMoney '���ο�֧��
Private mcllSquareBalance As Collection '���ѿ�����
Private mobjThreeSwap As clsThreeSwap

Private mstr����IDs As String '���˼���ID,79868
Private mbytԤ��������鿨 As Byte 'Ԥ�������ˢ�����ƣ�0-������ˢ������,1-��������ʱ��Ҫˢ����֤,2-��������ʱ��������ģ������ˢ����֤
Private mdblBrushCardMoney As Double
'----------------------------------------------------------------------------
Private mstrCurNos As String
Private mrsFeeData As ADODB.Recordset   '��¼����ˢ�����ѵ�����
Private mobjBalanceBills As BalanceBills 'ע�⣺����˳������� mstrCurNos ��˳��һ��
Private mblnCommitData As Boolean
Private mblnSaveBill As Boolean
Private mblnCommitBill As Boolean
Private mblnYbBalanced As Boolean
'----------------------------------------------------------------------------
'ҽ�����
Private Type TY_Insure
    intInsure As Integer
    strYBPati As String 'New:�ջ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID
    dbl������� As Double '��ǰ���˸����ʻ����
    dbl����͸֧ As Double '�����ʻ�����͸֧���
    colBalance As Collection '��¼���ŵ��ݱ��ս���ԭʼֵ���޸�ֵ,Ԫ��:BalanceMoneys
    
    strAllNos As String 'ԭ��ȡ���ĵ��ݣ����ܲ��ֽ���ɹ�
End Type
Private mInsure As TY_Insure '���ο�֧��
Private mstr�����ʻ� As String '�Ƿ񽫸����ʻ����õ��շѿ���
Private mInsurePara As Ty_InsurePara

Private mclsExpenceSvr As Object 'zlPublicExpense.clsExpenceSvr
Private mobjDrugStuff As clsDrugStuff

Public Function zlSquareAffirm(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal strPrivs As String, _
    Optional ByVal lngPatiID As Long = 0, _
    Optional ByVal lngCardTypeID As Long = 0, _
    Optional ByVal bln���ѿ� As Boolean = False, _
    Optional ByVal blnCliniqueRoomPay As Boolean = False, _
    Optional ByVal bytBillType As Byte, _
    Optional ByVal strNOs As String = "", _
    Optional ByVal strҽ��IDs As String = "", _
    Optional ByRef strExpand As String = "", _
    Optional ByRef lng����ID As Long = 0, _
    Optional ByVal blnʹ��Ԥ�� As Boolean = True) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����: ����ȷ�Ͻӿ� , ��Ҫ��Ӧ���ڲ����ڸ����ѻ�����������ȷ��
    '���:frmMain-������ö���
    '       lngModule:���õ�ģ���
    '       strPrivs:Ȩ�޴�
    '       lngPatiID :����ID,���Բ���,�ڱ��ӿڴ�����ˢ��!
    '       lngCardTypeID   Long    In  �����ID(���ѿ�Ϊ���ѽӿ����):0Ϊ������;��ȷ�ϴ����д��� Ŀǰ , ֻ����Ԥ����ɿ���ʹ��,�����,֧����ʽȱʡΪ�÷�ʽ.
    '       bln���ѿ�   Boolean In  ȱʡΪFase,��ʾ�Ƿ����ѿ�����
    '       bytBillType:�������: 0-�������շѻ���ʵ�,1-�շѼ�¼;2-���ʼ�¼
    '       strNOs:��ʽΪ( ����1,����2),���BytBillType��������ʹ��.һ��ֻ��ʹ��һ������
    '                   ��:  A0001,A002,A003��;
    '       strҽ��IDs:��ʽΪ:ID1,ID2,...
    '       strCardNO-��������ˢ�Ŀ���
    '       blnCliniqueRoomPay-���֧��(���֧��������ˢ������),���֧��ʱ��ֻ����շ�����
    '       blnʹ��Ԥ��-�Ƿ�����ʹ��Ԥ����Ture������ʹ��Ԥ����Ҵ���Ԥ����ʱȱʡʹ��Ԥ���False��������ʹ��Ԥ�������Ҫ�����õ������ʻ�
    '����:
    '����:Boolean ����    �ɹ�,����true,����ķ���False
    '����:���˺�
    '����:2011-06-15 09:53:37
    '˵��:
    '      ���strNos��strҽ��IDs��û��,ֻ�Ƕ�ָ�����˵������շѻ��۵��շѺ�������ʻ��۽������.
    '      �������ID������,����Ҫ�ڴ������Ƚ���ˢ���ҵ����˺�,�ٽ�������ȷ��.
    '������:
    '    1.  ���;����;ҩ����.
    '    2.  ����������Ҫ��������ȷ�ϵĵط���Ӧ�õ��øýӿ�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln��������Ԥ��ʹ��Ԥ���� As Boolean
    On Error GoTo errHandle
    Set mfrMain = frmMain
    mlngModule = lngModule: mlngPatiID = lngPatiID: mstrPrivs = strPrivs
    mstrInNos = strNOs: mstrҽ��IDs = strҽ��IDs
    mbytBillType = bytBillType: mlngCardTypeID = lngCardTypeID
    mblnCliniqueRoomPay = blnCliniqueRoomPay
    
    strExpand = "": mlng����ID = 0
    mblnOk = False: mstr����IDs = ""
    
    
    bln��������Ԥ��ʹ��Ԥ���� = Val(zlDatabase.GetPara(323, glngSys)) <> 1
    If zlCheckCurPatiIsMzLg(lngPatiID) Then     '�������۲���ʹ��Ԥ����
       blnʹ��Ԥ�� = bln��������Ԥ��ʹ��Ԥ����
    End If
    mblnʹ��Ԥ�� = blnʹ��Ԥ��
    
    Call InitVariableData
    
    Set mrsFeeData = GetFeeData(lngPatiID)
    If mrsFeeData Is Nothing Then Exit Function
    If mrsFeeData.State <> 1 Then Exit Function
    If mrsFeeData.RecordCount = 0 Then zlSquareAffirm = True: Exit Function
    
    If CreateOneCardComLib(frmMain, lngModule) = False Then Exit Function
    If CreateExpenceSvr(mclsExpenceSvr, lngModule) = False Then Exit Function
    
    Set mobjDrugStuff = New clsDrugStuff
    If mobjDrugStuff.InitCommon(mlngModule, mstrPrivs, mblnCliniqueRoomPay) = False Then Exit Function
    
    Call zlInitPriceGrade '��ʼ���۸�ȼ�
    
    Call InitPara
    If GetPatient(mlngPatiID) = False Then Exit Function
    If InitThreeSwap(frmMain) = False Then Exit Function
    
    If mblnCliniqueRoomPay Then
        If CliniqueRoomPayValied = False Then Exit Function
        If ExecuteCliniqueRoomPay(frmMain) = False Then Exit Function
        lng����ID = mlng����ID
        zlSquareAffirm = True
        Exit Function
    End If
    
    On Error Resume Next
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lng����ID = mlng����ID
    zlSquareAffirm = mblnOk
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitVariableData()
    '��ʼ��ģ�����
    Dim tyInsureTmp As TY_Insure
    Dim tyChargeTmp As TY_ChargeMoney
    
    mblnYbBalanced = False
    mblnCommitData = False
    mblnSaveBill = False
    mblnCommitBill = False
    
    mInsure = tyInsureTmp
    mCurCharge = tyChargeTmp
End Sub

Private Function CreateLocalTypeObject(ByVal lngCardTypeID As Long) As Boolean
    '����:����ָ����������
    '���:
    '   lngCardTypeID-�����ID
    '����:�����ɹ�����true,���򷵻�False
    Dim objCard As Card, blnReturn As Boolean
    Dim tyTemp As TY_PayMoney
    
    On Error GoTo ErrHandler
    blnReturn = gobjOneCardComLib.zlGetCard(lngCardTypeID, False, objCard)
    If blnReturn = False Or objCard Is Nothing Then
        ShowMsgbox "δ�ҵ�ָ���������ʻ���֧�ֵĿ���𣬿��ܸ����δ���ã������ҽ�ƿ����ݡ�"
        Exit Function
    End If
    If objCard.���� = False Then
        ShowMsgbox objCard.���� & "δ���ã����顣"
        Exit Function
    End If
    If objCard.�Ƿ�����ʻ� = False Then
        ShowMsgbox objCard.���� & "δ���������ʻ��������ҽ�ƿ����ݡ�"
        Exit Function
    End If
    If objCard.���㷽ʽ = "" Then
        ShowMsgbox objCard.���� & "δ���ý��㷽ʽ�������ҽ�ƿ����ݡ�"
        Exit Function
    End If
    If objCard.�ӿڳ����� = "" Then
        ShowMsgbox objCard.���� & "δ���������ӿ���֧�ֵĲ����������ҽ�ƿ����ݡ�"
        Exit Function
    End If
    
    mCurCardPay = tyTemp
    With mCurCardPay
       .lng�����ID = objCard.�ӿ����
       .bln���ѿ� = objCard.���ѿ�
       .str���㷽ʽ = objCard.���㷽ʽ
       .str���� = objCard.����
    End With
    CreateLocalTypeObject = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Private Function CliniqueRoomPayValied() As Boolean
    '����:���֧�����
    '����:�Ϸ�����true,���򷵻�False
    
    On Error GoTo ErrHandler
    If mbytBillType <> 1 Then   'ֻ����շѵ�
        ShowMsgbox "���֧��ʱ����������Լ��ʵ��ݽ���֧����"
        Exit Function
    End If
    If mlngCardTypeID = 0 Then
        ShowMsgbox "���֧��ʱҪ��ָ��һ����Ч�������ʻ�֧�����"
        Exit Function
    End If
 
    '���󴴽�ʧ�ܵ�,������֧��
    If Not CreateLocalTypeObject(mlngCardTypeID) Then Exit Function
    CliniqueRoomPayValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteCliniqueRoomPay(frmMain As Object) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���֧��
    '����:���֧���ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2014-01-14 17:28:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim curMoney As Currency, tyTmp As TY_ChargeMoney
    Dim strPrintNo As String '��ʽ��'A001','A002',...
    
    On Error GoTo errHandle
    mCurCharge = tyTmp
    mbytCurType = 1
    mstrCurNos = GetCurFeeNos(mbytCurType)
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl���úϼ� = curMoney
    Call Cacl�����
    
    If isValied() = False Then Exit Function
    '��������
    If SaveCharge(strPrintNo) = False Then Exit Function
    
    Call PrintBill(strPrintNo)
    '��ҽһ��ͨд����85950
    Call WriteInforToCard(frmMain, mlngModule, mstrPrivs, 0, strPrintNo)
    
    ExecuteCliniqueRoomPay = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFeeData(ByVal lng����ID As Long) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������ȡ�ķ�������
    '����:��ȡ��������
    '����:���˺�
    '����:2011-09-14 20:09:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTableNos As String, strTableIDs As String
    Dim varPara() As Variant, strSubTable As String
    Dim rsTemp As ADODB.Recordset
    Dim strSfTable As String, strJzTable As String
    
    On Error GoTo ErrHandler
    If lng����ID = 0 Then Exit Function
    ReDim Preserve varPara(0 To 1) As Variant
    
    varPara(0) = lng����ID: varPara(1) = mbytBillType
    
    If mstrҽ��IDs <> "" Then
        If zlGetSubTable(0, mstrҽ��IDs, strTableIDs, varPara(), 2) = False Then Exit Function
    End If
    If mstrInNos <> "" Then
        If zlGetSubTable(1, mstrInNos, strTableNos, varPara(), UBound(varPara) + 1) = False Then Exit Function
    End If
 
    If mstrҽ��IDs <> "" And mstrInNos <> "" Then
        strSubTable = " With  ҽ��  As (" & strTableIDs & "),���� as (" & strTableNos & ")"
    ElseIf mstrҽ��IDs <> "" Then
        strSubTable = " With  ҽ��  As (" & strTableIDs & ") "
    ElseIf strTableNos <> "" Then
        strSubTable = " With   ���� as (" & strTableNos & ")"
    End If
    '110421:���ϴ�,2017/6/23,����ִ��ʱӦʹ�ü۸񸸺Ŷ����Ǵ�������
    strSfTable = "": strJzTable = ""
    If mbytBillType <= 1 Then
        strSfTable = "" & _
        "Select '�շ�' As ���, a.��¼����, a.ִ�в���ID, a.��ҩ����, a.����ID, " & vbNewLine & _
        "       a.NO, nvl(A.�۸񸸺�,A.���) As ���," & _
        "       b.����||'-'||Decode(Decode(J1.�������||':'||J1.ҽ������,'7:***',1,0), 1, '***', B.����) As ��Ŀ," & vbNewLine & _
        "       b.���, nvl(a.����,1)*a.���� As ����, b.���㵥λ, a.�շ�ϸĿID, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��," & vbNewLine & _
        "       a.�շ����, a.�Ǽ�ʱ��, a.�����־,a.���ʽ, a.���˿���ID, a.��������ID, a.�Ƿ���, a.������Ŀ��, a.ͳ����" & vbNewLine & _
        "From ������ü�¼ A,�շ���ĿĿ¼ B ,����ҽ����¼ J1" & IIf(mstrInNos <> "", " ,���� C", "") & vbNewLine & _
        "Where a.�շ�ϸĿID=b.ID And a.��¼����=1 And a.����ID=[1] And (a.��¼״̬=0 Or a.��¼״̬=1 And a.����ID Is Null) " & _
        "      And a.ҽ�����=J1.id(+) "
        If mstrҽ��IDs <> "" And mstrInNos <> "" Then
            '����:49593
            strSfTable = strSfTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=1  ))"
        ElseIf mstrҽ��IDs <> "" Then
            strSfTable = strSfTable & " And  A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=1)"
        ElseIf strTableNos <> "" Then
            strSfTable = strSfTable & " And A.NO= C.Column_Value  "
        End If
    End If
    If mbytBillType = 2 Or mbytBillType = 0 Then
        strJzTable = "" & _
        "Select '����' As ���,A.��¼����,A.ִ�в���ID,A.��ҩ����,A.����ID, " & vbNewLine & _
        "       a.NO, nvl(A.�۸񸸺�,A.���) As ���," & _
        "       b.����||'-'||Decode(Decode(J1.�������||':'||J1.ҽ������,'7:***',1,0), 1, '***', B.����) As ��Ŀ," & vbNewLine & _
        "       b.���, nvl(a.����,1)*a.���� As ����, b.���㵥λ, a.�շ�ϸĿID, a.��׼����, a.Ӧ�ս��, a.ʵ�ս��," & vbNewLine & _
        "       a.�շ����, a.�Ǽ�ʱ��, a.�����־, a.���ʽ, a.���˿���ID, a.��������ID, a.�Ƿ���, a.������Ŀ��, a.ͳ����" & vbNewLine & _
        "From ������ü�¼ A,�շ���ĿĿ¼ B,����ҽ����¼ J1" & IIf(mstrInNos <> "", " ,���� C", "") & vbNewLine & _
        "Where a.�շ�ϸĿID=B.ID And a.��¼����=2 And a.����ID=[1] And a.��¼״̬=0 " & _
        "      And a.ҽ�����=J1.id(+) "
        If mstrҽ��IDs <> "" And mstrInNos <> "" Then
            '����:49593
            strJzTable = strJzTable & " And (A.NO= C.Column_Value  or  A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=2  ))"
        ElseIf mstrҽ��IDs <> "" Then
            strJzTable = strJzTable & " And   A.NO in (Select Distinct NO From ������ü�¼ J,ҽ�� P Where J.����ID=[1] and J.ҽ�����=P.ID And J.��¼����=2  ) "
        ElseIf strTableNos <> "" Then
            strJzTable = strJzTable & " And A.NO= C.Column_Value "
        End If
        If strSfTable <> "" Then strJzTable = vbCrLf & " Union all   " & vbCrLf & strJzTable
    End If
    strSQL = strSubTable & vbCrLf & strSfTable & vbCrLf & strJzTable
    strSQL = "  Select * From (" & strSQL & ") Order by ��¼����,NO,���"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "��ȡ���˷�����Ϣ", varPara)
    Set GetFeeData = rsTemp
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function LoadFeeData(ByVal bytType As Byte, Optional ByVal strNOs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ط�������
    ' ����:
    '   bytType-1-�����շ�;2-����
    '   strNos - ��ʽ��A001,A002,...
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-15 14:33:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    
    On Error GoTo ErrHandler
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "��¼����=" & bytType
    If mrsFeeData.RecordCount = 0 Then Exit Function
    
    mrsFeeData.Sort = "NO,���"
    With vsFee
        .Redraw = flexRDNone
        .Clear 1: .Rows = 1
        i = 1
        Do While Not mrsFeeData.EOF
            If strNOs = "" Or InStr("," & strNOs & ",", "," & nvl(mrsFeeData!NO) & ",") > 0 Then
                If i > .Rows - 1 Then .Rows = .Rows + 1
                .RowData(i) = Val(nvl(mrsFeeData!���))
                .TextMatrix(i, .ColIndex("���")) = nvl(mrsFeeData!���)
                .Cell(flexcpData, i, .ColIndex("���")) = Val(nvl(mrsFeeData!��¼����))
                .TextMatrix(i, .ColIndex("���ݺ�")) = nvl(mrsFeeData!NO)
                .Cell(flexcpData, i, .ColIndex("���ݺ�")) = Trim(nvl(mrsFeeData!�շ����))
                .TextMatrix(i, .ColIndex("��Ŀ")) = nvl(mrsFeeData!��Ŀ)
                .TextMatrix(i, .ColIndex("���")) = nvl(mrsFeeData!���)
                .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(nvl(mrsFeeData!����)), 5)
                .TextMatrix(i, .ColIndex("��λ")) = nvl(mrsFeeData!���㵥λ)
                .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(nvl(mrsFeeData!��׼����)), mintFeePrecision, , True)
                .TextMatrix(i, .ColIndex("Ӧ�ս��")) = FormatEx(Val(nvl(mrsFeeData!Ӧ�ս��)), mbytFeeMoneyPrecision, , True)
                .TextMatrix(i, .ColIndex("ʵ�ս��")) = FormatEx(Val(nvl(mrsFeeData!ʵ�ս��)), mbytFeeMoneyPrecision, , True)
                .Cell(flexcpData, i, .ColIndex("ʵ�ս��")) = Val(nvl(mrsFeeData!ʵ�ս��))
                .TextMatrix(i, .ColIndex("�����־")) = Val(nvl(mrsFeeData!�����־))
                
                i = i + 1
            End If
            mrsFeeData.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    
    LoadFeeData = True
    Exit Function
ErrHandler:
    vsFee.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetButtonVisible()
    '���ð�ť����ʾ״̬
    
    cmdYB.Visible = mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "�����շ�") _
        And (mInsure.intInsure = 0 Or mInsure.intInsure <> 0 And Not mblnYbBalanced)
    'ҽ����ҽ��δ���н���ʱ,����ʾ
    cmdYBBalance.Visible = mInsure.intInsure <> 0 And Not mblnYbBalanced
    'ҽ�����н����˵�,���ҽ����,��ʾ����շ�
    cmdOK.Visible = (mInsure.intInsure = 0 Or mInsure.intInsure <> 0 And mblnYbBalanced)
    cmdInsureSet.Visible = mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "�����շ�") And mInsure.intInsure = 0
End Sub

Private Sub SetControlProperty()
    '���ÿؼ�����
    On Error GoTo ErrHandler
    Call SetButtonVisible
    Call Cacl�����
    
    lbl(Lbl_Index.�Ը��ϼ�).Caption = FormatEx(mCurCharge.dbl���úϼ� - mCurCharge.dblҽ��֧��, 6, , , 2)
    lbl(Lbl_Index.��ǰδ��).Caption = Format(mCurCharge.dbl��ǰδ��, "0.00")
    
    lbl(Lbl_Index.���úϼ�).Caption = "���úϼ�:" & FormatEx(mCurCharge.dbl���úϼ�, 6, , , 2)
    lbl(Lbl_Index.�Ѹ��ϼ�).Caption = "�Ѹ��ϼ�:" & Format(mCurCharge.dbl�Ѹ��ϼ�, "0.00")
    lbl(Lbl_Index.ҽ��֧��).Caption = "ҽ��֧��:" & Format(mCurCharge.dblҽ��֧��, "0.00")
    
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cacl�����()
    '��ʾ�����
    Dim dblMoney As Double
    
    On Error GoTo ErrHandler
    mCurCharge.dbl�������� = 0
    mCurCharge.dbl��ǰδ�� = RoundEx(mCurCharge.dbl���úϼ� - mCurCharge.dbl�Ѹ��ϼ�, 6)
    
    dblMoney = RoundEx(mCurCharge.dbl��ǰδ��, 2)
    mCurCharge.dbl�������� = RoundEx(mCurCharge.dbl��ǰδ�� - dblMoney, 6)
    mCurCharge.dbl��ǰδ�� = RoundEx(mCurCharge.dbl��ǰδ�� - mCurCharge.dbl��������, 6)
    
    If mblnCliniqueRoomPay Then Exit Sub
    
    pic���.Visible = RoundEx(mCurCharge.dbl��������, 6) <> 0
    lbl(Lbl_Index.���).Caption = FormatEx(mCurCharge.dbl��������, 6, , , 2)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ClearData()
    '����:�����������
    lbl(Lbl_Index.����).Caption = ""
    lbl(Lbl_Index.�Ա�).Caption = "�Ա�:"
    lbl(Lbl_Index.�����).Caption = "�����:"
    
    lbl(Lbl_Index.Ԥ�����).Caption = "Ԥ�����:0.00"
    lbl(Lbl_Index.δ�����).Caption = "δ�����:0.00"
    lbl(Lbl_Index.ʣ����).Caption = "ʣ����:0.00"
    lbl(Lbl_Index.�������).Caption = "�������:0.00"
    
    lbl(Lbl_Index.�������).Visible = False
    lineUnder(Lbl_Index.�������).Visible = False
    
    lbl(Lbl_Index.��ǰδ��).Caption = "0.00"
    lbl(Lbl_Index.�Ը��ϼ�).Caption = "0.00"
    lbl(Lbl_Index.���).Caption = "0.00"
    
    lbl(Lbl_Index.���úϼ�).Caption = "���úϼ�:0.00"
    lbl(Lbl_Index.�Ѹ��ϼ�).Caption = "�Ѹ��ϼ�:0.00"
    lbl(Lbl_Index.ҽ��֧��).Caption = "ҽ��֧��:0.00"
    
    txt��Ԥ��.Text = "0.00"
    txt���.Text = "0.00"
    txtժҪ.Text = ""
    
    vsFee.Clear 1: vsFee.Rows = 2
    vsBalance.Clear 1: vsBalance.Rows = 2
    vsBalance.RowData(1) = ""
    
    staThis.Panels(Pan.C2��ʾ��Ϣ).Text = ""
    staThis.Panels(Pan.C3�����ʻ�).Text = ""
    staThis.Panels(Pan.C3�����ʻ�).Visible = False
End Sub

Private Sub SetControlMove()
    '����:���ÿؼ�����
    Dim sngTop As Single, sngSplitHeight As Single, blnԤ�� As Boolean
    Dim sngHeght As Single
    
    sngTop = 200: sngSplitHeight = 80
    blnԤ�� = mCurCharge.dbl����Ԥ�� <> 0 Or cbo֧����ʽ.ListCount = 0
    If mbytCurType = 1 And cbo֧����ʽ.ListCount > 0 Then
        lbl(Lbl_Index.Ԥ���).Visible = blnԤ��: txt��Ԥ��.Visible = blnԤ��
        If blnԤ�� Then
            txt��Ԥ��.Top = sngTop: sngTop = txt��Ԥ��.Top + txt��Ԥ��.Height + sngSplitHeight
        End If
        cbo֧����ʽ.Top = sngTop: sngTop = cbo֧����ʽ.Top + cbo֧����ʽ.Height + sngSplitHeight
        txt���.Top = cbo֧����ʽ.Top: btQRCodePay.Top = txt���.Top - 20
        
        txt���.Width = txt��Ԥ��.Left + txt��Ԥ��.Width - txt���.Left - IIf(mbytCurType = 1 And btQRCodePay.Tag <> "", btQRCodePay.Width + 10, 0)
    
        txtժҪ.Top = sngTop: txtժҪ.Height = picPayInfo.ScaleHeight - txtժҪ.Top - sngSplitHeight
        
        lbl(Lbl_Index.Ԥ���).Top = txt��Ԥ��.Top + (txt��Ԥ��.Height - lbl(Lbl_Index.Ԥ���).Height) \ 2
        lbl(Lbl_Index.�ɿ�).Top = cbo֧����ʽ.Top + (cbo֧����ʽ.Height - lbl(Lbl_Index.�ɿ�).Height) \ 2
        lbl(Lbl_Index.ժҪ).Top = txtժҪ.Top + 45
        Exit Sub
    End If
    
    sngHeght = picPayInfo.ScaleHeight
    sngHeght = sngHeght - txt��Ԥ��.Height
    sngTop = sngHeght / 2
    txt��Ԥ��.Top = sngTop
    lbl(Lbl_Index.Ԥ���).Top = txt��Ԥ��.Top + (txt��Ԥ��.Height - lbl(Lbl_Index.Ԥ���).Height) \ 2
    
    lbl(Lbl_Index.Ԥ���).Visible = True: txt��Ԥ��.Visible = True
    lbl(Lbl_Index.�ɿ�).Visible = False: cbo֧����ʽ.Visible = False: txt���.Visible = False
    lbl(Lbl_Index.ժҪ).Visible = False: txtժҪ.Visible = False
End Sub

Private Sub cbo֧����ʽ_Click()
    Dim tyTemp As TY_PayMoney
    Dim objCard As Card
    
    On Error GoTo ErrHandler
    mCurCardPay = tyTemp
    '���ʲ�����
    If mbytCurType = 2 Then Exit Sub
    
    Call GetCurCard(objCard)
    If objCard Is Nothing Then Exit Sub
    
    With mCurCardPay
        .lng�����ID = objCard.�ӿ����
        .bln���ѿ� = objCard.���ѿ�
        .str���㷽ʽ = objCard.���㷽ʽ
        .str���� = objCard.����
        .bln���ƿ� = objCard.���ƿ�
     End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdInsureSet_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub cmdYB_Click()
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State <> 1 Then Exit Sub
    If mrsInfo.RecordCount = 0 Then Exit Sub
    
    Call MCPatientProcess(mrsInfo!����ID)
End Sub

Private Function YBIdentifyCancel() As Boolean
    'ȡ��ҽ�����������֤
    Dim lng����ID As Long
    
    YBIdentifyCancel = True
    If mInsure.intInsure = 0 Then Exit Function
    If mInsure.strYBPati = "" Then Exit Function
    If mblnYbBalanced Then Exit Function
    
    If UBound(Split(mInsure.strYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mInsure.strYBPati, ";")(8)) And Val(Split(mInsure.strYBPati, ";")(8)) <> 0 Then
            lng����ID = Val(CLng(Split(mInsure.strYBPati, ";")(8)))
        End If
    End If
    If lng����ID = 0 Then Exit Function
    
    YBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng����ID, mInsure.intInsure)
End Function

Private Sub MCPatientProcess(ByVal lngCur����ID As Long)
    Dim i As Long, blnTran As Boolean, strSQL As String
    Dim lng����ID As Long, lng����IDOut As Long, intInsure As Integer
    Dim cur͸֧�� As Currency, strҽ���� As String
    Dim varNos As Variant, curMoney As Currency
    
    On Error GoTo ErrHandler
    If mrsInfo Is Nothing Then Exit Sub
    If mrsInfo.State = 0 Then Exit Sub
    
    If gblnLED Then zl9LedVoice.Speak "#50"
    mInsure.dbl������� = 0: mInsure.dbl����͸֧ = 0
    lng����IDOut = lngCur����ID '����Identify�ӿ����޸ĸñ����󷵻���ֵ
    
    '���أ�0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);8����ID,24��������(1=��������),25������������
    mInsure.strYBPati = gclsInsure.Identify(id�����շ�, lng����IDOut, mInsure.intInsure)
    If mInsure.strYBPati = "" Then
        mInsure.intInsure = 0: Exit Sub
    End If
    
    '��ȡ������Ϣ
    If UBound(Split(mInsure.strYBPati, ";")) >= 8 Then
        If IsNumeric(Split(mInsure.strYBPati, ";")(8)) And Val(Split(mInsure.strYBPati, ";")(8)) <> 0 Then
            lng����ID = Val(CLng(Split(mInsure.strYBPati, ";")(8)))
            If lng����ID <> lngCur����ID Then
                ShowMsgbox "ҽ����֤�Ĳ����뵱ǰ���˲���ͬһ�����ˣ�"
                staThis.Panels(Pan.C2��ʾ��Ϣ) = "ҽ����֤�Ĳ����뵱ǰ���˲���ͬһ�����ˣ�"
                GoTo IdentifyCancel:
            End If
        End If
    End If

    mInsurePara = initInsurePara(mInsure.intInsure, lng����ID)  '��ʼ��ҽ������
    
    '���¼��ز�����Ϣ������ҽ���ӿ����б䶯
    Call GetPatient(mlngPatiID)
    Call LoadPatient
    Call ShowLedInfor
    
    lbl(Lbl_Index.����).ForeColor = vbRed
    If nvl(mrsInfo!��������) <> "" Then
        Call SetPatiColor(lbl(Lbl_Index.����), nvl(mrsInfo!��������), vbRed)
    End If
        
    '�����ʻ�
    strҽ���� = CStr(Split(mInsure.strYBPati, ";")(1))
    mInsure.dbl������� = gclsInsure.SelfBalance(lng����ID, strҽ����, 10, cur͸֧��, mInsure.intInsure)
    staThis.Panels(Pan.C3�����ʻ�).Text = "�����ʻ����:" & Format(mInsure.dbl�������, "0.00")
    staThis.Panels(Pan.C3�����ʻ�).Visible = True
    mInsure.dbl����͸֧ = cur͸֧��
        
    '��������ȡ�Ļ��۵�����ر�������
    varNos = Split(mstrCurNos, ",")
    gcnOracle.BeginTrans: blnTran = True
    For i = 0 To UBound(varNos)
        strSQL = "zl_���ﻮ�ۼ�¼_Update(" & mInsure.intInsure & "," & lng����ID & ",'" & varNos(i) & "',0)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    
    '����Ԥ����������
    Set mInsure.colBalance = New Collection
    For i = 0 To UBound(varNos)
        mInsure.colBalance.Add New BalanceMoneys
    Next
    
    Set mrsFeeData = GetFeeData(lng����ID) '���¶�ȡ������Ϣ
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl���úϼ� = curMoney
    staThis.Panels(Pan.C2��ʾ��Ϣ) = ""
    'ֱ�ӽ���Ԥ����
    If ����Ԥ����() = False Then GoTo IdentifyCancel:
    
    If mInsurePara.����Ԥ���� Then
        Call InsureLedSpeak
    End If
    
    tbPage.Item(Pg_Index.Blance).Selected = True
    cmdYBBalance.Enabled = True
    Call SetControlProperty
    Call SetDefaultPrepayMoney
    Call SetCtlEnable(False)
    
    zlControl.ControlSetFocus vsBalance
    
    Exit Sub
IdentifyCancel:
    'ȡ��ҽ����֤
    Call YBIdentifyCancel
    mInsure.intInsure = 0: mInsure.strYBPati = ""
    
    Call SetPatiColor(lbl(Lbl_Index.����), nvl(mrsInfo!��������), &HFF0000)
    staThis.Panels(Pan.C3�����ʻ�).Text = ""
    staThis.Panels(Pan.C3�����ʻ�).Visible = False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Function GetBalanceBills(ByVal bytType As Byte, ByVal strNOs As String, _
    Optional ByRef curʵ�պϼ� As Currency) As BalanceBills
    '��ȡ���õ�����Ϣ
    '��Σ�
    '   bytType 1-�����շ�;2-�������
    Dim objBalanceBill As BalanceBill, objBalanceBills As BalanceBills
    Dim varNos As Variant, strNo As String
    Dim curʵ�ս�� As Currency
    Dim p As Integer, i As Integer
    
    On Error GoTo ErrHandler
    Set objBalanceBills = New BalanceBills
    curʵ�պϼ� = 0
    varNos = Split(strNOs, ",")
    For p = 1 To UBound(varNos) + 1
        strNo = varNos(p - 1)
        Set objBalanceBill = New BalanceBill
        objBalanceBill.NO = strNo
        
        mrsFeeData.Filter = "��¼����=" & bytType & " And No='" & strNo & "'"
        For i = 1 To mrsFeeData.RecordCount
            curʵ�ս�� = Val(nvl(mrsFeeData!ʵ�ս��))
            objBalanceBill.ʵ�պϼ� = objBalanceBill.ʵ�պϼ� + curʵ�ս��
            
            'ͳ�Ʊ��ս��
            If nvl(mrsFeeData!ͳ����, 0) = 0 Or Val(nvl(mrsFeeData!������Ŀ��)) = 0 Then
                '��ԭʼ���Ϊ׼,���ֱܷҴ���
                objBalanceBill.ȫ�Ը� = objBalanceBill.ȫ�Ը� + curʵ�ս��
            Else
                objBalanceBill.����ͳ�� = objBalanceBill.����ͳ�� + Val(nvl(mrsFeeData!ͳ����))
                '��ԭʼ���Ϊ׼,���ֱܷҴ���
                objBalanceBill.���Ը� = objBalanceBill.���Ը� + curʵ�ս�� - Val(nvl(mrsFeeData!ͳ����))
            End If
            
            curʵ�պϼ� = curʵ�պϼ� + curʵ�ս��
            mrsFeeData.MoveNext
        Next
        
        objBalanceBills.AddItem objBalanceBill, "K" & strNo
    Next
    Set GetBalanceBills = objBalanceBills
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ����Ԥ����() As Boolean
    '���ܣ�����Ԥ����
    Dim bytMode As Byte
    Dim i As Integer, j As Integer, k As Integer, p As Integer
    Dim dbl�ϼ� As Double, cur����Ԥ�� As Currency, cur���ø��� As Currency, cur����֧�� As Currency
    Dim objItem As BalanceMoney, strNone As String
    Dim strErrMsg As String
    Dim curʵ�պϼ� As Currency
    
    On Error GoTo ErrHandler
    '��ʼ�����������
    vsBalance.Clear 1: vsBalance.Rows = 2
    vsBalance.RowData(1) = ""
    If mInsure.intInsure = 0 Then Exit Function
    
    If mInsurePara.����Ԥ���� = False Then
        If mstr�����ʻ� = "" Then ����Ԥ���� = True: Exit Function
        
        '���㵱ǰ���ݸ����ʻ�֧�����:��֧��Ԥ����ʱ
        If mInsurePara.�൥�ݷֵ��ݽ��� Then
            For p = 1 To mobjBalanceBills.Count
                With mobjBalanceBills(p)
                    cur����Ԥ�� = .����ͳ�� + IIf(mInsurePara.���Ը�, .���Ը�, 0) + IIf(mInsurePara.ȫ�Ը�, .ȫ�Ը�, 0)
                End With
                'ͳ�Ƴ���֮ǰ���ݸ���֧����ĸ������
                cur���ø��� = mInsure.dbl�������
                For i = 1 To p - 1
                    cur���ø��� = cur���ø��� - GetMedicareSum(mInsure.colBalance, mstr�����ʻ�, i)
                Next
                
                cur����֧�� = Get���ʱ������(mobjBalanceBills(p).ʵ�պϼ�, cur����Ԥ��, cur���ø���, mInsure.dbl����͸֧)
                Call SetBalanceVal(mInsure.colBalance, p, mstr�����ʻ� & "|" & cur����֧��)
            Next
        Else
            cur����Ԥ�� = 0: curʵ�պϼ� = 0
            For i = 1 To mobjBalanceBills.Count
                With mobjBalanceBills(i)
                    cur����Ԥ�� = cur����Ԥ�� + .����ͳ�� + IIf(mInsurePara.���Ը�, .���Ը�, 0) + IIf(mInsurePara.ȫ�Ը�, .ȫ�Ը�, 0)
                    curʵ�պϼ� = curʵ�պϼ� + mobjBalanceBills(i).ʵ�պϼ�
                End With
            Next
            cur���ø��� = mInsure.dbl�������
            
            cur����֧�� = Get���ʱ������(curʵ�պϼ�, cur����Ԥ��, cur���ø���, mInsure.dbl����͸֧)
            Call SetBalanceVal(mInsure.colBalance, 1, mstr�����ʻ� & "|" & cur����֧��)
        End If
    Else
    
        If mInsurePara.ʵʱ��� Then
            '�������ڻ��۵��Ŵ�2������ϸ�ͻ��ܵļ�飬���ǣ���������ԭ��������ʵ�ս����������ͨ������ܸı䣬�������ٴμ����ϸ
            '1.���뵥�ݣ�2.�޸ĵ��ݣ�3.������ҩ�䷽��4.�޸���ҩ�����������еĸ���ͬʱ�仯��5.��������Զ���������Լ�������ܼ����ۿ�
            '6.�޸ĵ��ۣ�7.����ִ�п��ң�ҩƷ�۸����㣬8.�����ѱ�ʵ�ս������,9.�����������֤ҽ�����,�����ȵ�
            If gclsInsure.CheckItem(mInsure.intInsure, 0, 2, MakeDetailRecord(mobjBalanceBills)) = False Then
                staThis.Panels(Pan.C2��ʾ��Ϣ).Text = "������Ŀ���ʧ�ܣ�"
                Exit Function
            End If
        End If
    
        If mInsurePara.�൥�ݷֵ��ݽ��� Then
            bytMode = 2
        ElseIf mInsurePara.һ�ν���ֵ����˷� Then
            bytMode = 1
        Else
            bytMode = 0
        End If
        
        If ZlExecuteInsurePreSwap(bytMode, mobjBalanceBills, mInsure.intInsure, mInsure.colBalance, strErrMsg) = False Then
            staThis.Panels(Pan.C2��ʾ��Ϣ).Text = strErrMsg
            Exit Function
        End If
        strNone = CheckInsureBalanceValid(mrs���㷽ʽ, mInsure.colBalance)
        If strNone <> "" Then
            ShowMsgbox "��ǰ���ս���ʹ�õĽ��㷽ʽ" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                "������δ���ã����ȵ����㷽ʽ������������Щ���㷽ʽ��"
            Exit Function
        End If
    End If
    
    'ȫ��Ԥ�����Ĵ���
    '-----------------------------------------------------------
    '��ʾԤ��ı����
    For p = 1 To mInsure.colBalance.Count
        For Each objItem In mInsure.colBalance(p)
            With vsBalance
                '��λ��ƥ���л����
                k = -1
                For j = 1 To .Rows - 1
                    If .TextMatrix(j, .ColIndex("֧����ʽ")) = objItem.���㷽ʽ Then
                        k = j: Exit For '��¼����д��ƥ����
                    ElseIf .TextMatrix(j, .ColIndex("֧����ʽ")) = "" Then
                        If k = -1 Then k = j '��¼��һ���ÿ���
                    End If
                Next
                If j > .Rows - 1 And k = -1 Then
                    .Rows = .Rows + 1
                    k = .Rows - 1
                End If
                
                '���ܸ��ֽ��㷽ʽ�Ľ��
                .TextMatrix(k, .ColIndex("֧����ʽ")) = objItem.���㷽ʽ
                .TextMatrix(k, .ColIndex("֧�����")) = Format(Val(.TextMatrix(k, .ColIndex("֧�����"))) + objItem.ԭʼ���, "0.00")
                dbl�ϼ� = dbl�ϼ� + Val(Format(objItem.ԭʼ���, "0.00"))
                If .RowData(k) = 0 Then
                    '���ŵ�����,ֻҪ��һ�������޸�,����ܵ������޸�
                    .RowData(k) = IIf(objItem.�����޸�, 1, 0)
                End If
            End With
        Next
    Next
    mCurCharge.dblҽ��֧�� = dbl�ϼ�
    mCurCharge.dbl�Ѹ��ϼ� = dbl�ϼ�
    
    ����Ԥ���� = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InsureLedSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ҽ��Ԥ��Led����
    '����:���˺�
    '����:2011-12-15 13:40:46
    '����:44425
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double
    
    If Not gblnLED Then Exit Sub
    dbl���ʺϼ� = GetMedicareSum(mInsure.colBalance, mstr�����ʻ�)
    zl9LedVoice.DisplayBank "ҽ������:", "�ʻ����" & Format(mInsure.dbl�������, "0.00"), _
        "�ʻ�֧��" & Format(dbl���ʺϼ�, "0.00"), "ͳ��֧��" & Format(GetMedicareSum(mInsure.colBalance) - dbl���ʺϼ�, "0.00")
    zl9LedVoice.Speak "#21 " & Format(mCurCharge.dbl���úϼ�, "0.00")
End Sub

Private Sub LedDisplayBank(Optional ByVal blnSpeak As Boolean = True)
    '����:��ʾ������Ϣ
    '����:52117
    Dim i As Long
    Dim strҽ�� As String, str�������� As String, str��һ��ͨ As String, str��ͨ���� As String
    Dim varPara  As Variant, str���㷽ʽ As String
    Dim strTemp As String
    
    If Not gblnLED Then Exit Sub
    With vsBalance
        For i = 1 To .Rows - 1
            'ҽ������
            If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then
                strTemp = .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("֧�����"))), "0.00")
                Select Case .RowData(i)
                Case Enum_BalanceType.ҽ��
                    strҽ�� = strҽ�� & "||" & strTemp
                Case Enum_BalanceType.һ��ͨ
                    str�������� = str�������� & "||" & strTemp
                Case Enum_BalanceType.��һ��ͨ
                    str��һ��ͨ = str��һ��ͨ & "||" & strTemp
                Case Else
                    str��ͨ���� = str��ͨ���� & "||" & strTemp
                End Select
            End If
        Next
    End With
     
    str���㷽ʽ = ""
    If strҽ�� <> "" Then str���㷽ʽ = str���㷽ʽ & "||ҽ������:||�ʻ����:" & Format(mInsure.dbl�������, "0.00") & strҽ��
    If str�������� <> "" Then str���㷽ʽ = str���㷽ʽ & "||һ��ͨ����:" & str��������
    If str��һ��ͨ <> "" Then str���㷽ʽ = str���㷽ʽ & "||һ��ͨ����(��):" & str��һ��ͨ
    If str��ͨ���� <> "" Then str���㷽ʽ = str���㷽ʽ & "" & str��ͨ����
    If str���㷽ʽ = "" Then Exit Sub
    
    str���㷽ʽ = Mid(str���㷽ʽ, 3)
    varPara = Split(str���㷽ʽ, "||")
    'Ŀǰ���ֻ����ʾ10������ֵ
    Select Case UBound(varPara)
    Case 0
          zl9LedVoice.DisplayBank varPara(0)
    Case 1
          zl9LedVoice.DisplayBank varPara(0), varPara(1)
    Case 2
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2)
    Case 3
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3)
    Case 4
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4)
    Case 5
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5)
    Case 6
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6)
    Case 7
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7)
    Case 8
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8)
    Case 9
          zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9)
    Case Else
        str���㷽ʽ = ""
        For i = 10 To UBound(varPara)
            str���㷽ʽ = str���㷽ʽ & ";" & varPara(i)
        Next
        If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str���㷽ʽ
    End Select

    If blnSpeak Then zl9LedVoice.Speak "#21 " & Format(mCurCharge.dbl��ǰδ��, "0.00")
End Sub

Private Sub cmdYBBalance_Click()
    Dim blnSpeak As Boolean, dblOldδ֧�� As Currency
    
    On Error GoTo ErrHandler
    dblOldδ֧�� = mCurCharge.dbl��ǰδ��
    '�������ݱ���
    If SaveFeeBill() = False Then Exit Sub
    '����ҽ������
    If ExecuteInsureSwap() = False Then
        Call SetButtonVisible
        Exit Sub
    End If
    
    Call LoadBalancedData(mCurCharge.lng����ID)
    Call SetControlProperty
    Call SetButtonVisible
    Call SetCtlEnable
     
    Call SetDefaultPrepayMoney
    Call SetBeginFocus '��궨λ
    
    blnSpeak = RoundEx(dblOldδ֧��, 6) <> RoundEx(mCurCharge.dbl��ǰδ��, 6)
    Call LedDisplayBank(blnSpeak)
    If RoundEx(mCurCharge.dbl��ǰδ��, 6) = 0 Then
        'ҽ��ȫ������,ֱ��ȷ�����
        If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
    End If
    
    If RoundEx(mCurCharge.dbl��ǰδ��, 6) < 0 Then
        'ҽ�������������˷����ܽ��ʱ����Ҫ�˿������
        MsgBox "    ����ҽ�������������˷����ܽ��޷���ɽ��㡣" & vbCrLf & _
            "�뵽�շѴ��ڽ��д���", vbExclamation + vbOKOnly, gstrSysName
        Unload Me: Exit Sub
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveFeeBill() As Boolean
    '������õ�������
    '˵��:
    '   ���ô˹���ʱ,����Ҫ��ʼ�����쳣ʱ,���ݻ��ˣ�����ɹ�ʱ,δ�ύ����
    Dim objBalanceBill As BalanceBill
    Dim blnTrans As Boolean, strSQL As String
    Dim str����ʱ�� As String, str��ҩ���� As String
    Dim cllDept As Collection, int������Դ As Integer
    Dim varNos As Variant, strNo As String, p As Integer
    Dim strErrMsg As String, i As Integer
    
    On Error GoTo ErrHandler
    If (mblnSaveBill And mblnCommitBill) Or mblnCommitData Then
        gcnOracle.BeginTrans
        SaveFeeBill = True: Exit Function
    End If
    
    '���۵������շѼ��
    varNos = Split(mstrCurNos, ",")
    For i = 0 To UBound(varNos)
        If mclsExpenceSvr.zlPriceChargeCheck(varNos(i), mlngPatiID, strErrMsg) = False Then
            MsgBox IIf(strErrMsg = "", "ִ�л����շѼ�����", strErrMsg), vbInformation, gstrSysName
            Exit Function
        End If
    Next
    
    mCurCharge.lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    mCurCharge.lng������� = -1 * mCurCharge.lng����ID
    str����ʱ�� = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    int������Դ = IIf(Val(nvl(mrsInfo!��Ժ)) = 1, 2, 1)
    
    For p = 1 To UBound(varNos) + 1
        strNo = varNos(p - 1)
        mrsFeeData.Filter = "��¼����=" & mbytCurType & " And No='" & strNo & "'"
        If mrsFeeData.RecordCount <> 0 Then
            '��ҩ����
            Set cllDept = New Collection
            Do While Not mrsFeeData.EOF
                If InStr(",5,6,7,", nvl(mrsFeeData!�շ����)) > 0 Then
                    cllDept.Add Array(nvl(mrsFeeData!�շ����), Val(nvl(mrsFeeData!ִ�в���ID)), nvl(mrsFeeData!��ҩ����))
                End If
                mrsFeeData.MoveNext
            Loop
            str��ҩ���� = GetPayDrugWindow(mlngPatiID, CDate(str����ʱ��), cllDept)
            
            mrsFeeData.MoveFirst
            'Zl_���˻����շ�_Insert
            strSQL = "Zl_���˻����շ�_Insert("
            '  No_In         ������ü�¼.NO%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "" & ZVal(mlngPatiID) & ","
            '  ������Դ_In   Number,
            strSQL = strSQL & "" & int������Դ & ","
            '  ���ʽ_In   ������ü�¼.���ʽ%Type,
            If nvl(mrsInfo!���ʽ����) <> "" Then
               strSQL = strSQL & "'" & nvl(mrsInfo!���ʽ����) & "',"
            Else
               strSQL = strSQL & "'" & nvl(mrsFeeData!���ʽ) & "',"
            End If
            '  ����_In       ������ü�¼.����%Type,
            strSQL = strSQL & "'" & nvl(mrsInfo!����) & "',"
            '  �Ա�_In       ������ü�¼.�Ա�%Type,
            strSQL = strSQL & "'" & nvl(mrsInfo!�Ա�) & "',"
            '  ����_In       ������ü�¼.����%Type,
            strSQL = strSQL & "'" & nvl(mrsInfo!����) & "',"
            '  ���˿���id_In ������ü�¼.���˿���id%Type,
            strSQL = strSQL & "" & ZVal(nvl(mrsFeeData!���˿���ID)) & ","
            '  ��������id_In ������ü�¼.��������id%Type,
            strSQL = strSQL & "" & ZVal(nvl(mrsFeeData!��������ID)) & ","
            '  ������_In     ������ü�¼.������%Type,
            strSQL = strSQL & "NULL,"    ' �����ڲ�����,����ԭ���Ĳ���
            '  ����id_In     ������ü�¼.����id%Type,
            strSQL = strSQL & "" & mCurCharge.lng����ID & ","
            '  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
            strSQL = strSQL & "To_Date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'),"
            '  ����Ա���_In ������ü�¼.����Ա���%Type,
            strSQL = strSQL & "'" & UserInfo.��� & "',"
            '  ����Ա����_In ������ü�¼.����Ա����%Type,
            strSQL = strSQL & "'" & UserInfo.���� & "',"
            '  ��ҩ����_In   ������ü�¼.��ҩ����%Type := Null,
            strSQL = strSQL & "'" & str��ҩ���� & "',"
            '  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
            strSQL = strSQL & "" & Val(nvl(mrsFeeData!�Ƿ���)) & ","
            '  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
            strSQL = strSQL & "To_Date('" & str����ʱ�� & "','yyyy-mm-dd hh24:mi:ss'))"
            
            mobjBalanceBills("K" & strNo).�����շ�SQL = strSQL
            If mInsure.intInsure <> 0 Then
                Set mobjBalanceBills("K" & strNo).Ԥ���� = mInsure.colBalance(p)
            End If
        End If
    Next
    
    gcnOracle.BeginTrans: blnTrans = True
    For Each objBalanceBill In mobjBalanceBills
        zlDatabase.ExecuteProcedure objBalanceBill.�����շ�SQL, Me.Caption
    Next
    
    mblnSaveBill = True
    SaveFeeBill = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteInsureSwap() As Boolean
    'ҽ������
    Dim bytMode As Byte, blnCommit As Boolean
    Dim strErrMsg As String
    Dim strSavedNos As String, lngSavedBillCount As Long, blnYbBalanced As Boolean
    
    On Error GoTo ErrHandler
    If mInsure.intInsure = 0 Then ExecuteInsureSwap = True: Exit Function
    
    If mInsurePara.�൥�ݷֵ��ݽ��� Then
        bytMode = 2
    ElseIf mInsurePara.һ�ν���ֵ����˷� Then
        bytMode = 1
    Else
        bytMode = 0
    End If
    
    mInsure.strAllNos = ""
    If zlExecuteInsureSwap(bytMode, mlngPatiID, mInsure.intInsure, mstr�����ʻ�, _
        mblnֻ��ҽ������ɹ������շ�, mCurCharge.lng����ID, mCurCharge.lng�������, _
        mobjBalanceBills, blnCommit, strSavedNos, lngSavedBillCount, blnYbBalanced, strErrMsg) = False Then
        If blnCommit = False Then
            If strErrMsg <> "" Then ShowMsgbox strErrMsg
            Exit Function
        End If
        
        mblnCommitBill = True
        '���¼�������
        If blnYbBalanced Then
            mInsure.strAllNos = mstrCurNos
            mstrCurNos = strSavedNos
            Call LoadFeeData(mbytCurType, strSavedNos)
            
            mblnYbBalanced = True 'ҽ���Ѿ�����
            ExecuteInsureSwap = True
        End If
    Else
        mblnCommitBill = True
        mblnYbBalanced = True 'ҽ���Ѿ�����
        ExecuteInsureSwap = True
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ClearBanalce()
    '�����������
    With mCurCharge
        .dbl���úϼ� = 0
        .dblҽ��֧�� = 0
        .dbl�Ѹ��ϼ� = 0
        .dbl��ǰδ�� = 0
        .dbl���γ�Ԥ�� = 0
        .dbl�������� = 0
    End With
End Sub

Private Sub AddNewRow()
    '����������һ��
    With vsBalance
        If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
            .Rows = .Rows + 1
            .RowPosition(.Rows - 1) = 1
        End If
    End With
End Sub

Private Function LoadBalancedData(ByVal lng����ID As Long) As Boolean

    '�����ѽ���ɹ��Ľ�������
    Dim strSQL As String, rsBalance As ADODB.Recordset
    Dim bln���ѿ� As Boolean, lng�����ID As Long
    Dim rsTypes As ADODB.Recordset
    Dim bln���� As String, str������� As String
    On Error GoTo ErrHandler
    Call ClearBanalce
    vsBalance.Clear 1: vsBalance.Rows = 2
    vsBalance.RowData(1) = ""
    
    If gobjOneCardComLib Is Nothing Then Call CreateOneCardComLib(Me, mlngModule, gcnOracle)
    If Not gobjOneCardComLib Is Nothing Then
       Call gobjOneCardComLib.zlGetOneCardTypes(rsTypes)
    End If
    
    
    
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    Set rsBalance = GetChargeBalance(lng����ID)
    Do While Not rsBalance.EOF
        Select Case nvl(rsBalance!����)
        Case Enum_BalanceType.Ԥ���
            mCurCharge.dbl���γ�Ԥ�� = RoundEx(mCurCharge.dbl���γ�Ԥ�� + Val(nvl(rsBalance!��Ԥ��)), 6)
            mCurCharge.dbl�Ѹ��ϼ� = RoundEx(mCurCharge.dbl�Ѹ��ϼ� + Val(nvl(rsBalance!��Ԥ��)), 6)
        Case Else
            If nvl(rsBalance!����) = Enum_BalanceType.ҽ�� Then
                mCurCharge.dblҽ��֧�� = RoundEx(mCurCharge.dblҽ��֧�� + nvl(rsBalance!��Ԥ��, 0), 6)
            End If
            
            If Val(nvl(rsBalance!У�Ա�־)) = 2 Then
                bln���ѿ� = nvl(rsBalance!����) = Enum_BalanceType.���ѿ�
                If bln���ѿ� Then
                    lng�����ID = Val(nvl(rsBalance!���㿨���))
                Else
                    lng�����ID = Val(nvl(rsBalance!�����ID))
                End If
                
                With vsBalance
                    Call AddNewRow
                    .RowData(1) = nvl(rsBalance!����)
                    .TextMatrix(1, .ColIndex("֧����ʽ")) = nvl(rsBalance!���㷽ʽ)
                    str������� = nvl(rsBalance!���������, nvl(rsBalance!���㷽ʽ))
                    bln���� = Val(nvl(rsBalance!�Ƿ�����)) = 1
                    If Not bln���ѿ� And Not rsTypes Is Nothing Then
                        rsTypes.Filter = "ID=" & lng�����ID
                        If Not rsTypes.EOF Then
                            bln���� = Val(nvl(rsTypes!�Ƿ�����)) = 1
                            str������� = nvl(rsTypes!����)
                        End If
                    End If
                    'ҽ�ƿ����ID|���ѿ�(1,0)|�ӿ�����
                    .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = lng�����ID & "|" & IIf(bln���ѿ�, 1, 0) & "|" & str�������
                    .TextMatrix(1, .ColIndex("֧�����")) = Format(Val(nvl(rsBalance!��Ԥ��)), "0.00")
                    .TextMatrix(1, .ColIndex("��ע")) = nvl(rsBalance!ժҪ)
                    .TextMatrix(1, .ColIndex("������ˮ��")) = nvl(rsBalance!������ˮ��)
                    .TextMatrix(1, .ColIndex("����˵��")) = nvl(rsBalance!����˵��)
                    
                    .TextMatrix(1, .ColIndex("����")) = IIf(bln����, String(Len(nvl(rsBalance!����)), "*"), nvl(rsBalance!����))
                    .Cell(flexcpData, 1, .ColIndex("����")) = nvl(rsBalance!����)
                    .TextMatrix(1, .ColIndex("����״̬")) = 1
                    .Cell(flexcpBackColor, 1, 0, 1, .Cols - 1) = Me.BackColor
                End With
                mCurCharge.dbl�Ѹ��ϼ� = RoundEx(mCurCharge.dbl�Ѹ��ϼ� + Val(nvl(rsBalance!��Ԥ��)), 6)
            End If
        End Select
        
        rsBalance.MoveNext
    Loop
                   
    strSQL = "Select Sum(b.ʵ�ս��) As ʵ�պϼ� From ������ü�¼ B Where b.����id = [1]"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    mCurCharge.dbl���úϼ� = Val(rsBalance!ʵ�պϼ�)
        
    If mCurCharge.dbl���γ�Ԥ�� <> 0 Then
        txt��Ԥ��.Text = Format(mCurCharge.dbl���γ�Ԥ��, "0.00")
        txt��Ԥ��.Tag = "1"
        txt��Ԥ��.BackColor = Me.BackColor
        txt��Ԥ��.Enabled = False
    End If
    LoadBalancedData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    If mblnCliniqueRoomPay Then Exit Sub
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "�����շ�") Then '136681
        mInsure.intInsure = GetCustomPatiInsure(mrsInfo!����ID)
        If mInsure.intInsure <> 0 Then
            Call MCPatientProcess(mrsInfo!����ID)
        End If
    End If
    
    Call SetBeginFocus '��궨λ
    '78773:���ϴ�,2014-10-29,LED��ʾһ��֧ͨ����Ϣ
    Call ShowLedInfor
End Sub

Private Sub SetBeginFocus()
    '���ÿ�ʼʱ�Ľ���λ��
    If Val(txt��Ԥ��.Text) <> 0 And mblnʹ��Ԥ�� Or cbo֧����ʽ.ListCount = 0 Or mbytCurType = 2 Then
        zlControl.ControlSetFocus txt��Ԥ��: zlControl.TxtSelAll txt��Ԥ��
    Else
        zlControl.ControlSetFocus txt���: zlControl.TxtSelAll txt���
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF6
        If cmdYB.Visible And cmdYB.Enabled Then Call cmdYB_Click
    Case vbKeyF2
        'ǿ�����
        If mInsure.intInsure <> 0 And mblnYbBalanced = False Then
            If cmdYBBalance.Visible And cmdYBBalance.Enabled Then Call cmdYBBalance_Click
        Else
            If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
        End If
    Case vbKeyF4
        If Me.ActiveControl Is txt��� And txt���.Enabled Then
            If cbo֧����ʽ.Visible = False Or cbo֧����ʽ.Enabled = False Then Exit Sub
            If Shift = vbShiftMask Then
                If cbo֧����ʽ.ListIndex - 1 < 0 Then
                    cbo֧����ʽ.ListIndex = cbo֧����ʽ.ListCount - 1
                Else
                    cbo֧����ʽ.ListIndex = cbo֧����ʽ.ListIndex - 1
                End If
            Else
                If cbo֧����ʽ.ListIndex + 1 > cbo֧����ʽ.ListCount - 1 Then
                    cbo֧����ʽ.ListIndex = 0
                Else
                    cbo֧����ʽ.ListIndex = cbo֧����ʽ.ListIndex + 1
                End If
            End If
        End If
    Case vbKeyF3    'ɨ�븶���
        If btQRCodePay.Visible And btQRCodePay.Enabled Then Call btQRCodePay.zlReReadQRCode
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If YBIdentifyCancel() = False Then 'ȡ��ҽ�����������֤,���ؼ�ʱ���˳�
        Cancel = 1: Exit Sub
    End If
    
    '78773:���ϴ�,2014-10-29,LED��ʾһ��֧ͨ����Ϣ
    If gblnLED Then zl9LedVoice.DisplayPatient ""
    If Not mobjThreeSwap Is Nothing Then Set mobjThreeSwap = Nothing
    If Not mobjDrugStuff Is Nothing Then Set mobjDrugStuff = Nothing
    
    Set mobjPayCards = Nothing
    Set mrsInfo = Nothing
    Set mrs���㷽ʽ = Nothing
    Set mrsFeeData = Nothing
    SaveWinState Me, App.ProductName, mstrTittle
End Sub

Private Sub lbl_Change(Index As Integer)
    Select Case Index
    Case Lbl_Index.Ԥ���
        lbl(Index).Tag = ""
    End Select
End Sub

Private Sub picBlance_Resize()
    On Error Resume Next
    With picBlance
        vsBalance.Left = .ScaleLeft
        vsBalance.Height = .ScaleHeight - vsBalance.Top
        vsBalance.Width = .ScaleWidth - vsBalance.Left
    End With
End Sub

Private Sub picBlanceAndFee_Resize()
    On Error Resume Next
    With picBlanceAndFee
        tbPage.Left = .ScaleLeft + 30
        tbPage.Top = .ScaleTop + 10
        tbPage.Height = .ScaleHeight - tbPage.Top - 40
        tbPage.Width = .ScaleWidth - tbPage.Left - 40
    End With
    zlControl.PicShowFlat picBlanceAndFee, -1, , 1
End Sub

Private Sub picFee_Resize()
    On Error Resume Next
    With picFee
        vsFee.Left = .ScaleLeft
        vsFee.Top = .ScaleTop
        vsFee.Height = .ScaleHeight - vsFee.Top
        vsFee.Width = .ScaleWidth - vsFee.Left
    End With
End Sub

Private Function LoadԤ�����(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ԥ�����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-21 10:47:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim dbl������� As Double, dbl������� As Double, dbl������� As Double
    On Error GoTo errHandle
    
    '79868,�����˼��������벡��ʣ���
    '��ü�¼�����ֻ��������һ���ǲ��˱��˵ģ�һ���ǲ��˼�����
    Set rsTemp = GetMoneyInfo(lng����ID, , , 1, , , True)
    With mCurCharge
        .dblԤ����� = 0
        .dbl������� = 0
        Do While Not rsTemp.EOF
            .dblԤ����� = .dblԤ����� + Val(nvl(rsTemp!Ԥ�����))
            .dbl������� = .dbl������� + Val(nvl(rsTemp!�������))
            If nvl(rsTemp!����, 0) = 0 Then
                dbl������� = Val(nvl(rsTemp!Ԥ�����))
                dbl������� = Val(nvl(rsTemp!�������))
            Else
                dbl������� = Val(nvl(rsTemp!Ԥ�����)) - Val(nvl(rsTemp!�������))
            End If
            rsTemp.MoveNext
        Loop
        .dbl����Ԥ�� = .dblԤ����� - .dbl�������
        If .dbl����Ԥ�� < 0 Then .dbl����Ԥ�� = 0
    End With
    If mblnʹ��Ԥ�� = False And mbytCurType = 1 Then
        mCurCharge.dbl����Ԥ�� = 0: dbl������� = 0
    End If
    
    lbl(Lbl_Index.Ԥ�����).Caption = "Ԥ�����:" & Format(dbl�������, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.δ�����).Caption = "δ�����:" & Format(dbl�������, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.ʣ����).Caption = "ʣ����:" & Format(dbl������� - dbl�������, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.�������).Caption = "�������:" & Format(dbl�������, "###0.00;-###0.00;0.00;0.00")
    lbl(Lbl_Index.�������).Visible = RoundEx(dbl�������, 6) <> 0
    lineUnder((Lbl_Index.�������)).Visible = RoundEx(dbl�������, 6) <> 0
    LoadԤ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô����С
    '����:���˺�
    '����:2011-09-15 11:26:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '��С����ߴ�
    With gWinRect
        .MaxW = Me.Width
        .MaxH = Screen.Height * Screen.TwipsPerPixelY
        .MinH = Me.Height
        .MinW = Me.Width
    End With
    glngOld = GetWindowLong(hwnd, GWL_WNDPROC)
    Call SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SetWindowResizeWndMessage)
End Sub

Private Sub Form_Load()
    Dim curMoney As Currency
    
    mblnFirst = True
    If mblnCliniqueRoomPay Then Exit Sub
    
    mstrTittle = "�������ѽ���"
    
    If mbytBillType = 0 Then
        mrsFeeData.Filter = "��¼����=1"
        mbytCurType = IIf(mrsFeeData.RecordCount = 0, 2, 1)
    Else
        mbytCurType = mbytBillType
    End If
    Call InitFace
    If LoadPatient() = False Then Unload Me: Exit Sub
    
    mstrCurNos = GetCurFeeNos(mbytCurType)
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl���úϼ� = curMoney
    
    mstr�����ʻ� = ""
    Set mrs���㷽ʽ = Get���㷽ʽ()
    If Not mrs���㷽ʽ.EOF Then
        mrs���㷽ʽ.Filter = "����=3"
        If Not mrs���㷽ʽ.EOF Then mstr�����ʻ� = nvl(mrs���㷽ʽ!����)
    End If
    If LoadԤ�����(mrsInfo!����ID) = False Then Unload Me: Exit Sub
    If Load֧����ʽ() = False Then Unload Me: Exit Sub
    If LoadFeeData(mbytCurType) = False Then Unload Me: Exit Sub
    
    Call SetCtlEnable
    Call SetControlMove
    Call SetControlProperty
    Call SetDefaultPrepayMoney
End Sub

Public Function Get���㷽ʽ() As ADODB.Recordset
    '��ȡ���н��㷽ʽ���ݣ����ֽ��㳡�ϣ�Ҳ����������
    Dim strSQL As String
    
    On Error GoTo ErrHandler
    strSQL = _
        "Select b.����, b.����, b.ȱʡ��־ As ȱʡ, Nvl(b.����, 1) As ����, Nvl(b.Ӧ����, 0) As Ӧ����" & vbNewLine & _
        "From ���㷽ʽ B" & vbNewLine & _
        "Where b.���� <> 9"
    Set Get���㷽ʽ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitFace()
    '��ʼ������
    If mblnFirst Then
        RestoreWinState Me, App.ProductName, mstrTittle
        If Not IsDesinMode Then Call SetWindowsSize
        
        zlControl.CboSetWidth cbo֧����ʽ.hwnd, cbo֧����ʽ.Width * 2
        zlControl.PicShowFlat picPatientInfo, -1, , 1
        zlControl.PicShowFlat picʣ���Ը�, -1, , 1
        zlControl.PicShowFlat pic�Ը��ϼ�, -1, , 1
        zlControl.PicShowFlat picPayInfo, -1, , 1
    End If
    
    Call InitPage
    picBlance.Visible = (mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "�����շ�"))
    
    Call ClearData
End Sub

Private Sub InitPage()
    '����:��ʼ��ҳ��ؼ�
    Dim objItem As TabControlItem
    
    On Error GoTo ErrHandler
    tbPage.RemoveAll
    If mbytCurType = 1 And zlstr.IsHavePrivs(mstrPrivs, "�����շ�") Then
        Set objItem = tbPage.InsertItem(Pg_Index.Blance, "������Ϣ", picBlance.hwnd, 0)
        objItem.Tag = Pg_Index.Blance
    End If
    Set objItem = tbPage.InsertItem(Pg_Index.FeeDetail, "������ϸ", picFee.hwnd, 0)
    objItem.Tag = Pg_Index.FeeDetail
    
    With tbPage
        .Item(.ItemCount - 1).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
ErrHandler:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitFactPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��Ʊ��صĲ���
    '����:���˺�
    '����:2011-08-11 00:24:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mPara
        .int�շ�Ʊ�ݸ�ʽ = Val(zlDatabase.GetPara("�շ��վݸ�ʽ", glngSys, 1151))
        .int�շѴ�ӡ��ʽ = Val(zlDatabase.GetPara("�շѴ�ӡ��ʽ", glngSys, 1151))
        .int���Ʊ�ݸ�ʽ = Val(zlDatabase.GetPara("����վݸ�ʽ", glngSys, 1151))
        .int��˴�ӡ��ʽ = Val(zlDatabase.GetPara("��˴�ӡ��ʽ", glngSys, 1151))
        .intҩƷ��λ = Val(zlDatabase.SetPara("ҩƷ��λ", glngSys, 1151))
    End With
End Sub

Private Sub InitPara()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������ֵ
    '����:���˺�
    '����:2011-06-20 16:48:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    
    Call InitFactPara
    '���ﲡ������ʱ��Ҫˢ����֤
    strValue = zlDatabase.GetPara(28, glngSys, , "1|0")
    If InStr(strValue, "|") = 0 Then strValue = "1|0"
    mdblBrushCardMoney = Val(Split(strValue, "|")(0))
    If mdblBrushCardMoney < 0 Then
        mbytԤ��������鿨 = 3
        mdblBrushCardMoney = -1 * mdblBrushCardMoney
    Else
        mbytԤ��������鿨 = mdblBrushCardMoney
    End If
    
    '���õ��۱���λ��
    mintFeePrecision = Val(zlDatabase.GetPara(157, glngSys, , "5"))
    '���ý��С����λ��
    mbytFeeMoneyPrecision = Val(zlDatabase.GetPara(9, glngSys, , 2))
    
    '�Զ�����
    mbln�����Զ����� = zlDatabase.GetPara(92, glngSys) = "1"

    mblnֻ��ҽ������ɹ������շ� = Val(zlDatabase.GetPara(317, glngSys, , "0")) = 1
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    picBlanceAndFee.Width = Me.ScaleWidth - picBlanceAndFee.Left * 2
    picBlanceAndFee.Height = Me.ScaleHeight - staThis.Height - picBlanceAndFee.Top
End Sub

Private Function GetPatient(ByVal lngPatiID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '���:blnCard=��ʾ�Ƿ���￨ˢ��
    '����:
    '����:���˶�ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-20 16:04:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    On Error GoTo errH
    '��ȡ������Ϣ
    strSQL = "Select Decode(Sign(a.����ʱ�� - a.�Ǽ�ʱ��), 0, 1, 0) As ����, a.����id, a.��������, a.Ic����, a.���￨��," & vbNewLine & _
            "        a.�����, a.סԺ��, a.����, a.����֤��, a.�Ա�, a.����, a.��������, a.�ѱ�," & vbNewLine & _
            "        a.ҽ�Ƹ��ʽ, m.���� As ���ʽ����, a.��Ժ, Decode(B1.��������, Null, 0, 1, 1, 0) As ����," & vbNewLine & _
            "        B1.��Ժ����, a.����, c.���� As ��������" & vbNewLine & _
            " From ������Ϣ A, ������ҳ B1, ������� C, ҽ�Ƹ��ʽ M" & vbNewLine & _
            " Where a.����id = B1.����id(+) And a.��ҳid = B1.��ҳid(+) And a.���� = c.���(+)" & vbNewLine & _
            "       And a.ҽ�Ƹ��ʽ = m.����(+) And a.ͣ��ʱ�� Is Null And a.����id = [1]"
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, "�������ѽ���-��ȡ������Ϣ", lngPatiID)
    If mrsInfo.EOF Then
        ShowMsgbox "������Ϣδ�ҵ������飡"
        Set mrsInfo = New ADODB.Recordset
        Exit Function
    End If
    GetPatient = True
    Exit Function
errH:
    Set mrsInfo = New ADODB.Recordset
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadPatient() As Boolean
    '���ز�����Ϣ
    If mrsInfo Is Nothing Then Exit Function
    If mrsInfo.State <> 1 Then Exit Function
    If mrsInfo.RecordCount = 0 Then Exit Function
    
    lbl(Lbl_Index.����).Caption = nvl(mrsInfo!����)
    lbl(Lbl_Index.�Ա�).Caption = "�Ա�:" & nvl(mrsInfo!�Ա�)
    lbl(Lbl_Index.�����).Caption = "�����:" & nvl(mrsInfo!�����)
    '74309:���ϴ���2014-7-7������������ʾ��ɫ����
    Call SetPatiColor(lbl(Lbl_Index.����), nvl(mrsInfo!��������), &HFF0000)
    LoadPatient = True
End Function

Private Function Load֧����ʽ() As Boolean
    '������Ч��֧����ʽ�������õ�������
    Dim i As Long, objCards As Cards, lngKey As Long
    Dim strRQCardTypeIDs As String
    
    Set mobjPayCards = New Cards
     
    ' zlGetCards(ByVal BytType As Byte)
    'bytType-  0-����ҽ�ƿ�;1-���õ�ҽ�ƿ�, 2-���д��������˻���������3-���õ������˻���ҽ�ƿ�
    Set objCards = gobjOneCardComLib.zlGetCards(3)

    With cbo֧����ʽ
        .Clear
        For i = 1 To objCards.Count
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
            
            .AddItem objCards(i).����
            .ItemData(.NewIndex) = i
        Next
    End With
    If cbo֧����ʽ.ListCount > 0 Then cbo֧����ʽ.ListIndex = 0
    
    If mbytCurType = 1 Then
        strRQCardTypeIDs = mobjThreeSwap.GetRQCardTypeIDsFromCards(mobjPayCards)
        If Not btQRCodePay.zlInit(Me, strRQCardTypeIDs, glngSys, mlngModule, gcnOracle, gstrDBUser) Then strRQCardTypeIDs = ""
        btQRCodePay.Tag = strRQCardTypeIDs  '�洢��Ч�Ŀ����IDs
        btQRCodePay.Visible = btQRCodePay.Tag <> ""
    Else
        btQRCodePay.Visible = False
    End If
    
    Load֧����ʽ = True
End Function

Private Sub SetCtlEnable(Optional ByVal blnEdit As Boolean = True)
    '���ÿؼ��Ŀ���״̬
    
    If blnEdit Then blnEdit = (mInsure.intInsure = 0 Or mInsure.intInsure <> 0 And mblnYbBalanced)
    picPayInfo.Enabled = blnEdit
    txt��Ԥ��.Enabled = blnEdit And UsedPrepayMoney() = False And mCurCharge.dbl����Ԥ�� > 0
    btQRCodePay.Enabled = blnEdit And btQRCodePay.Tag <> "" 'Tag:�洢������Ч֧�ֵ�ɨ�븶�Ŀ����Ids
    txt���.Enabled = blnEdit
    txtժҪ.Enabled = blnEdit
    vsBalance.Editable = IIf(mInsure.intInsure <> 0 And mblnYbBalanced = False, flexEDKbdMouse, flexEDNone)
    
    '������ʾ��ɫ
    txt��Ԥ��.BackColor = IIf(txt��Ԥ��.Enabled, &H80000005, Me.BackColor)
    cbo֧����ʽ.BackColor = IIf(txt��Ԥ��.Enabled, &H80000005, Me.BackColor)
    txt���.BackColor = IIf(txt���.Enabled, &H80000005, Me.BackColor)
    txtժҪ.BackColor = IIf(txtժҪ.Enabled, &H80000005, Me.BackColor)
End Sub

Private Function UsedPrepayMoney() As Boolean
    '�ж��Ƿ���ʹ��Ԥ����
    Dim i As Integer
    
    On Error GoTo ErrHandler
    For i = 1 To vsBalance.Rows - 1
        If vsBalance.RowData(i) = 1 Then
            UsedPrepayMoney = True: Exit Function
        End If
    Next
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "Calc" Then Call ShowWindowsCalculator
End Sub

Private Sub txt��Ԥ��_Change()
    txt��Ԥ��.Tag = ""
    txt���.Text = "0.00"
End Sub

Private Sub txt��Ԥ��_GotFocus()
    zlControl.TxtSelAll txt��Ԥ��
End Sub

Private Sub txt��Ԥ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Val(txt��Ԥ��.Text) = 0 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If BrushcardStrikePrepay(False) = False Then
        zlControl.ControlSetFocus txt��Ԥ��: zlControl.TxtSelAll txt��Ԥ��
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt��Ԥ��_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt��Ԥ��, KeyAscii, m���ʽ)
End Sub

Private Function CheckPrepayValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ��������Ƿ���Ч
    '����:��Ч,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-14 22:30:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If txt��Ԥ��.Text = "" Then
        txt��Ԥ��.Text = "0.00"
    ElseIf Not IsNumeric(txt��Ԥ��.Text) And txt��Ԥ��.Text <> "" Then
        MsgBox "��Ч��ֵ��", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf Val(txt��Ԥ��.Text) < 0 Then
        MsgBox "Ԥ��������Ϊ����", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf Val(txt��Ԥ��.Text) > 0 And RoundEx(mCurCharge.dbl��ǰδ��, 6) < 0 Then
        MsgBox "��ǰδ�����Ϊ��ʱ����ʹ��Ԥ��", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf RoundEx(Val(txt��Ԥ��.Text), 6) > RoundEx(mCurCharge.dbl����Ԥ��, 6) Then
        MsgBox "Ԥ�������ܳ������˵�Ԥ�����:" & FormatEx(mCurCharge.dbl����Ԥ��, 6, , , 2) & " ��", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    ElseIf RoundEx(Val(txt��Ԥ��.Text), 6) > RoundEx(mCurCharge.dbl��ǰδ��, 6) Then
        MsgBox "Ԥ�������ܴ���δ�����:" & Format(mCurCharge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        GoTo InvalidDataHandler:
    Else
        txt��Ԥ��.Text = Format(Val(txt��Ԥ��.Text), "0.00")
    End If
    CheckPrepayValied = True
    Exit Function
InvalidDataHandler:
    Call SetDefaultPrepayMoney
    zlControl.ControlSetFocus txt��Ԥ��: zlControl.TxtSelAll txt��Ԥ��
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt��Ԥ��_LostFocus()
    If txt��Ԥ��.Tag = "1" Then Exit Sub
    Call SetControlProperty
End Sub

Private Sub txt��Ԥ��_Validate(Cancel As Boolean)
    If txt��Ԥ��.Tag = "1" Then Exit Sub
    If CheckPrepayValied = False Then Cancel = True: Exit Sub
End Sub

Private Sub txt���_GotFocus()
    txt���.Text = Format(mCurCharge.dbl��ǰδ�� - Val(txt��Ԥ��.Text), "0.00")
    zlControl.TxtSelAll txt���
End Sub

Private Sub txt���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Val(txt���.Text) = 0 Then txt���.Text = "0.00"
    txt���.Text = Format(Val(txt���.Text), "0.00")
    If cmdOK.Visible And cmdOK.Enabled Then Call cmdOK_Click
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt���, KeyAscii, m���ʽ)
End Sub

Private Sub txt���_Validate(Cancel As Boolean)
    txt���.Text = Format(Val(txt���.Text), "0.00")
End Sub

Private Sub SetDefaultPrepayMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ȱʡԤ�����
    '����:���˺�
    '����:2011-08-13 17:21:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt���.Text = "0.00"
    txt��Ԥ��.Text = "0.00"
    With mCurCharge
        If mbytCurType = 2 Then
            txt��Ԥ��.Text = Format(.dbl��ǰδ��, "###0.00;###0.00;0.00;0.00")
            Exit Sub
        End If
        If RoundEx(.dbl����Ԥ��, 6) <> 0 Then
            If RoundEx(.dbl����Ԥ��, 6) > RoundEx(.dbl��ǰδ��, 6) Then
                txt��Ԥ��.Text = Format(.dbl��ǰδ��, "###0.00;###0.00;0.00;0.00")
            Else
                txt��Ԥ��.Text = Format(.dbl����Ԥ��, "###0.00;###0.00;0.00;0.00")
            End If
        End If
    End With
End Sub

Private Function CheckThreeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������׽�������Ƿ�Ϸ�
    '����:�Ϸ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-09-15 00:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If txt���.Visible = False Or txt���.Enabled = False Then CheckThreeValied = True: Exit Function
    
    If Val(txt���) = 0 Then
        ShowMsgbox "δ���뽻�׽����飡"
        zlControl.ControlSetFocus txt���: zlControl.TxtSelAll txt���
        Exit Function
    End If
    If Not IsNumeric(txt���.Text) And txt���.Text <> "" Then
        ShowMsgbox "��Ч��ֵ��"
        zlControl.ControlSetFocus txt���: zlControl.TxtSelAll txt���
        Exit Function
    ElseIf Val(txt���.Text) < 0 Then
        ShowMsgbox "���׽���Ϊ����"
    ElseIf RoundEx(Val(txt���.Text) + Val(txt��Ԥ��.Text), 2) > RoundEx(mCurCharge.dbl��ǰδ��, 2) And Val(txt���.Text) <> 0 Then
        ShowMsgbox "���׽��ܴ��ڱ���δ�����:" & Format(mCurCharge.dbl��ǰδ�� - Val(txt��Ԥ��.Text), "0.00") & " ��"
        zlControl.ControlSetFocus txt���: zlControl.TxtSelAll txt���
        Exit Function
    End If
    CheckThreeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetClassMoney(ByVal bytType As Byte, ByVal strNOs As String, _
    ByRef rsClass As ADODB.Recordset) As Boolean
    '��ȡ������ܽ��
    '��Σ�
    '   bytType 1-�����շ�;2-�������
    Dim i As Integer
    Dim varNos As Variant
    
    On Error GoTo ErrHandler
    Set rsClass = New ADODB.Recordset
    rsClass.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    rsClass.Fields.Append "���", adDouble, , adFldIsNullable
    rsClass.CursorLocation = adUseClient
    rsClass.LockType = adLockOptimistic
    rsClass.CursorType = adOpenStatic
    rsClass.Open
    
    varNos = Split(strNOs, ",")
    For i = 0 To UBound(varNos)
        mrsFeeData.Filter = "��¼����=" & bytType & " And No='" & varNos(i) & "'"
        Do While Not mrsFeeData.EOF
            rsClass.Find "�շ����='" & nvl(mrsFeeData!�շ����) & "'", , adSearchForward, 1
            If rsClass.EOF Then rsClass.AddNew
            rsClass!�շ���� = nvl(mrsFeeData!�շ����)
            rsClass!��� = RoundEx(Val(nvl(rsClass!���)) + Val(nvl(mrsFeeData!ʵ�ս��)), 6)
            rsClass.Update
            
            mrsFeeData.MoveNext
        Loop
    Next
    GetClassMoney = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function BrushCardThreeSwapCheck(ByVal strNOs As String, _
    ByVal dblMoney As Double, ByVal str������Դ As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����֤
    '���:strNos -����֧���ĵ��ݺ�
    '       dblMoney-֧�����ܽ��
    '����:����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsClassMoney As ADODB.Recordset, cllSquareBalance As Collection
    
    On Error GoTo errHandle
    mCurCardPay.str֧������ = ""
    If mbytCurType = 2 Then BrushCardThreeSwapCheck = True: Exit Function
    If mCurCardPay.lng�����ID = 0 Then BrushCardThreeSwapCheck = True: Exit Function
    
    If mblnCliniqueRoomPay = False Then
        If CheckThreeValied() = False Then Exit Function
    End If
    
    If mCurCardPay.bln���ѿ� Then
        If GetClassMoney(mbytCurType, strNOs, rsClassMoney) = False Then Exit Function
        '�������ѿ���ˢ����Ϣ
        Set cllSquareBalance = mcllSquareBalance
    End If
    
    If mobjThreeSwap.CheckPayValid(mCurCardPay.lng�����ID, mCurCardPay.bln���ѿ�, mCurCardPay.str���㷽ʽ, _
        dblMoney, strNOs, mCurCardPay.strˢ������, mCurCardPay.strˢ������, , mCurCardPay.str֧������, _
        rsClassMoney, str������Դ, cllSquareBalance, mCurCardPay.strQRCode) = False Then Exit Function
    
    If mCurCardPay.bln���ѿ� Then Set mcllSquareBalance = cllSquareBalance
    
    BrushCardThreeSwapCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetCurFeeNos(ByVal bytType As Byte) As String
    '��ȡ���ݺ�
    '��Σ�
    '   bytType 1-�����շ�;2-�������
    '����:���ݺ�,����֮���ö��ŷ���,��:A0001,A0002....
    Dim strNOs As String
    
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "��¼����=" & bytType
    If mrsFeeData.RecordCount = 0 Then Exit Function
    
    mrsFeeData.Sort = "NO"
    With mrsFeeData
        Do While Not .EOF
            If InStr(strNOs & ",", "," & nvl(!NO) & ",") = 0 Then
                strNOs = strNOs & "," & nvl(!NO)
            End If
            .MoveNext
        Loop
    End With
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    GetCurFeeNos = strNOs
End Function

Private Function Get������Դ(ByVal bytType As Byte) As String
    '��ȡ���ݺ�
    '��Σ�
    '   bytType 1-�����շ�;2-�������
    '����:
    Dim str������Դ As String
    
    If mrsFeeData Is Nothing Then Exit Function
    mrsFeeData.Filter = "��¼����=" & bytType
    If mrsFeeData.RecordCount = 0 Then Exit Function
    
    With mrsFeeData
        Do While Not .EOF
            If InStr(str������Դ, Decode(Val(!�����־), 4, 3, 2, 2, 1)) = 0 Then
                str������Դ = str������Դ & "," & Decode(Val(!�����־), 4, 3, 2, 2, 1)
            End If
            .MoveNext
        Loop
    End With
    If str������Դ <> "" Then str������Դ = Mid(str������Դ, 2)
    Get������Դ = str������Դ
End Function

Private Function GetSelectNOsAndSerialNum(ByRef strNOs As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡѡ��ĵ��ݺź͵����е����
    '����:���ݺ�,����֮���ö��ŷ���,��:A0001:1,2|A0002:1,2,3|....
    '����:���˺�
    '����:2011-06-23 10:01:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strNo As String
    Dim str��� As String, strData As String
    Dim j As Long
    
    With vsFee
        strData = "": strNOs = ""
        For i = 1 To .Rows - 1
            strNo = Trim(.TextMatrix(i, .ColIndex("���ݺ�")))
            If InStr(1, strNOs & ",", "," & strNo & ",") = 0 Then
                str��� = ""
                For j = 1 To .Rows - 1
                    If strNo = Trim(.TextMatrix(j, .ColIndex("���ݺ�"))) Then
                        str��� = str��� & "," & .RowData(j)
                    End If
                Next
                If str��� <> "" Then str��� = Mid(str���, 2)
                strNOs = strNOs & "," & strNo
                strData = strData & "|" & strNo & ":" & str���
            End If
        Next
    End With
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    If strData <> "" Then strData = Mid(strData, 2)
    GetSelectNOsAndSerialNum = strData
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݺϷ��Լ��
    '����:���ݺϷ�������true,���򷵻�False
    '����:���˺�
    '����:2011-06-22 15:28:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��Ԥ��  As Double, dblThreeMoney  As Double
    Dim str������Դ As String
    
    If mrsInfo Is Nothing Then
        ShowMsgbox "������Ϣ����ȷ�������飡"
        zlControl.ControlSetFocus cmdCancel: Exit Function
    End If
    
    If mrsInfo.State <> 1 Then
        ShowMsgbox "������Ϣ����ȷ�������飡"
        zlControl.ControlSetFocus cmdCancel: Exit Function
    End If
    
    If mbytCurType = 1 Then
        dbl��Ԥ�� = 0: dblThreeMoney = 0
        If mblnCliniqueRoomPay = False Then '�����֧��ʱ����Ҫ�����ص����ݺϷ���
            If Not CheckTextLength("ժҪ", txtժҪ) Then Exit Function
            
            If txt��Ԥ��.Visible And txt��Ԥ��.Enabled Then dbl��Ԥ�� = Val(txt��Ԥ��.Text)
            If txt���.Visible And txt���.Enabled Then dblThreeMoney = Val(txt���.Text)
        
            If cbo֧����ʽ.ListIndex >= 0 Then
                If mCurCardPay.str���㷽ʽ = "" Then
                    ShowMsgbox mCurCardPay.str���� & " δ���ý��㷽ʽ������ϵͳ����Ա��ϵ��"
                    Exit Function
                End If
            ElseIf RoundEx(dblThreeMoney, 6) <> 0 Then
                ShowMsgbox "δѡ��֧����ʽ��"
                Exit Function
            End If
            
            If RoundEx(dbl��Ԥ�� + dblThreeMoney, 6) <> RoundEx(mCurCharge.dbl��ǰδ��, 6) Then
                If Val(txt���.Text) = 0 And txt��Ԥ��.Visible Then
                    ShowMsgbox "���˵�Ԥ������㣬���ֵ��"
                    zlControl.ControlSetFocus txt��Ԥ��: zlControl.TxtSelAll txt��Ԥ��
                ElseIf txt��Ԥ��.Visible = False Then
                    ShowMsgbox "����" & cbo֧����ʽ.Text & "֧�����(" & _
                        Format(dblThreeMoney, "0.00") & ")�뱾��δ�����(" & _
                        Format(mCurCharge.dbl��ǰδ��, "0.00") & ")���ȣ����飡"
                    zlControl.ControlSetFocus txt���: zlControl.TxtSelAll txt���
                Else
                    ShowMsgbox "����֧�����ϼ�(Ԥ���+" & cbo֧����ʽ.Text & ":" & _
                        Format(dbl��Ԥ�� + dblThreeMoney, "0.00") & ")�뱾��δ�����(" & _
                        Format(mCurCharge.dbl��ǰδ��, "0.00") & ")���ȣ����飡"
                    zlControl.ControlSetFocus txt���: zlControl.TxtSelAll txt���
                End If
                Exit Function
            End If
            
            If RoundEx(dbl��Ԥ��, 6) > 0 And Val(txt��Ԥ��.Tag) = 0 Then
                '֤��û����֤������Ҫ����������֤
                If BrushcardStrikePrepay(True) = False Then Exit Function
            End If
        Else
            dblThreeMoney = mCurCharge.dbl��ǰδ��
        End If
        
        str������Դ = Get������Դ(mbytCurType)
        If RoundEx(dblThreeMoney, 6) <> 0 Then
            If BrushCardThreeSwapCheck(mstrCurNos, dblThreeMoney, str������Դ) = False Then Exit Function
        End If
    Else
        If Val(txt��Ԥ��.Tag) = 0 Then
            '֤��û����֤������Ҫ����������֤
            If BrushcardStrikePrepay(True) = False Then Exit Function
        End If
    End If
    isValied = True
End Function

Private Function BrushcardStrikePrepay(ByVal blnOKClick As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤ˢ����Ԥ��
    '����:�����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-14 14:35:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double
    
    On Error GoTo ErrHandler
    If Val(txt��Ԥ��.Tag) = 1 Then BrushcardStrikePrepay = True: Exit Function
    If Val(txt��Ԥ��.Text) = 0 And mbytCurType = 1 Then BrushcardStrikePrepay = True: Exit Function
    
    If mbytCurType <> 2 Then
        If CheckPrepayValied() = False Then Exit Function
    End If
    dblMoney = Val(txt��Ԥ��.Text)
    
    gblnNotCloseWindows = True
    'ˢ��ȷ��
    If zlPatiIdentify(mlngModule, Me, mlngPatiID, dblMoney, False, 1, mlngCardTypeID, True, , mstr����IDs) Then
        gblnNotCloseWindows = False
                    
        txt��Ԥ��.Tag = "1" '�������֤
        
        '��������Ԥ����¼��������ID����
        If mobjThreeSwap.UpgradeHistoryData(mlngModule, mlngPatiID, mstr����IDs) = False Then Exit Function
        '59412
        If blnOKClick Then BrushcardStrikePrepay = True: Exit Function
        
        If RoundEx(dblMoney, 6) = RoundEx(mCurCharge.dbl��ǰδ��, 6) Or mbytCurType = 2 Then
            '���ʱ,��������
            Call cmdOK_Click
            If mblnOk Then BrushcardStrikePrepay = True: Exit Function
        ElseIf mbytCurType = 1 And cbo֧����ʽ.ListCount = 0 Then
            ShowMsgbox "���˵�Ԥ������㣬���ֵ��"
            Exit Function
        End If
        
        Call SetControlProperty
        BrushcardStrikePrepay = True
    Else
        gblnNotCloseWindows = False
        BrushcardStrikePrepay = False
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Dim ty_Tmp As TY_Insure
    
    On Error GoTo ErrHandler
    If mbytCurType = 1 And mInsure.intInsure <> 0 And mInsure.strYBPati <> "" Then
        If mblnCommitBill Then
            If MsgBox("    ��ǰ���ڶ�ҽ�������շѣ��˳��󱾴ν��㽫����Ϊ�쳣״̬��" & vbCrLf & _
                "��Ҫ���շѴ��ڽ��д���ȷʵҪ�˳���", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        Else
            'ȡ��ҽ����֤
            If YBIdentifyCancel() = False Then Exit Sub
            mInsure = ty_Tmp
            mCurCharge.dblҽ��֧�� = 0
            mCurCharge.dbl�Ѹ��ϼ� = 0
            
            vsBalance.Clear 1: vsBalance.Rows = 2
            vsBalance.RowData(1) = ""
            tbPage.Item(Pg_Index.FeeDetail).Selected = True
            cmdYBBalance.Enabled = False
            
            Call SetPatiColor(lbl(Lbl_Index.����), nvl(mrsInfo!��������), &HFF0000)
            staThis.Panels(Pan.C3�����ʻ�).Text = ""
            staThis.Panels(Pan.C3�����ʻ�).Visible = False
            
            Call SetControlProperty
            Call SetCtlEnable
            Call SetDefaultPrepayMoney
            Call SetBeginFocus '��궨λ
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

Private Sub PrintBill(ByVal strPrintNo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ӡƱ��
    '��Σ�
    '   strPrintNO ��ʽ��'A001','A002',...
    '����:���˺�
    '����:2014-01-20 11:01:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, strFormat As String
    Dim frmMain As Object
    
    If mblnCliniqueRoomPay Then
        Set frmMain = mfrMain
    Else
        Set frmMain = Me
    End If
    Select Case mbytCurType
    Case 1
        blnPrint = mPara.int�շѴ�ӡ��ʽ = 1
        If mPara.int�շѴ�ӡ��ʽ = 2 Then
            If MsgBox("���Ƿ����Ҫ��ӡ�嵥��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int�շ�Ʊ�ݸ�ʽ = 0, "", "ReportFormat=" & mPara.int�շ�Ʊ�ݸ�ʽ)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & strPrintNo, "ҩƷ��λ=" & mPara.intҩƷ��λ, "PrintEmpty=0", strFormat, 2)
        End If
    Case 2
        blnPrint = mPara.int��˴�ӡ��ʽ = 1
        If mPara.int��˴�ӡ��ʽ = 2 Then
            If MsgBox("���Ƿ����Ҫ��ӡ�嵥��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                blnPrint = True
            End If
        End If
        If blnPrint Then
            strFormat = IIf(mPara.int���Ʊ�ݸ�ʽ = 0, "", "ReportFormat=" & mPara.int���Ʊ�ݸ�ʽ)
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1151", frmMain, "NO=" & strPrintNo, "ҩƷ��λ=" & mPara.intҩƷ��λ, "PrintEmpty=0", strFormat, 2)
        End If
    End Select
End Sub

Private Sub cmdOK_Click()
    Dim blnPartialSaved As Boolean '���ֱ���ɹ�
    Dim curMoney As Currency, bln�����շ� As Boolean
    Dim strPrintNo As String '��ʽ��'A001','A002',...
    
    On Error GoTo errHandle
    '����У��
    If isValied = False Then Exit Sub
    If SaveData(strPrintNo, blnPartialSaved) = False Then Exit Sub
    If blnPartialSaved Then Unload Me: Exit Sub
    
    '��ӡƱ��
    Call PrintBill(strPrintNo)
    
    '��ҽһ��ͨд����85950
    If mbytCurType = 1 Then
        Call WriteInforToCard(Me, mlngModule, mstrPrivs, 0, strPrintNo)
    End If
    
    bln�����շ� = False
    If mbytCurType = 1 And mInsure.strAllNos <> "" Then
        If MsgBox("��ǰֻ�ɹ���ȡ��" & UBound(Split(mstrCurNos, ",")) + 1 & "�ŵ��ݵķ��ã�" & _
                  "�Ƿ��δ��ȡ�ɹ��ĵ��ݽ��������շѣ�", _
            vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            
            mstrCurNos = GetRemainNos(mInsure.strAllNos, mstrCurNos)
            Set mrsFeeData = GetFeeData(mlngPatiID)
            mrsFeeData.Filter = "��¼����=1"
            If mrsFeeData.RecordCount > 0 Then
                bln�����շ� = True
            End If
        End If
    End If
    
    If bln�����շ� = False Then
        '0-�������շѻ���ʵ�ʱ���Լ��˵��������
        If mbytBillType = 0 And mbytCurType = 1 Then
            mbytCurType = 2
            bln�����շ� = True
        End If
    End If
        
    If bln�����շ� = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    '��������ʣ�����
    Call InitVariableData
    Call InitFace
    If LoadFeeData(mbytCurType) = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    Call LoadPatient
    If LoadԤ�����(mrsInfo!����ID) = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    If Load֧����ʽ() = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    mstrCurNos = GetCurFeeNos(mbytCurType)
    Set mobjBalanceBills = GetBalanceBills(mbytCurType, mstrCurNos, curMoney)
    mCurCharge.dbl���úϼ� = curMoney
    
    Call SetCtlEnable
    Call SetControlMove
    Call SetControlProperty
    Call SetDefaultPrepayMoney
    Call SetBeginFocus '��궨λ
    
    Call ShowLedInfor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function GetRemainNos(ByVal strAllNos As String, ByVal strSavedNos As String) As String
    '��ȡʣ�൥�ݺ�
    '��Σ�
    '   strAllNos ���е��ݺţ�A001,A002,...
    '   strSavedNos �Ա��浥�ݺţ�A001,A002,...
    '���أ�ʣ�൥�ݺţ�A001,A002,...
    Dim varAllNos As Variant, strNOs As String
    Dim i As Integer
    
    varAllNos = Split(strAllNos, ",")
    For i = 0 To UBound(varAllNos)
        If InStr("," & strSavedNos & ",", "," & varAllNos(i) & ",") = 0 Then
            strNOs = strNOs & "," & varAllNos(i)
        End If
    Next
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    GetRemainNos = strNOs
End Function

Private Function VerifyFee(ByRef strPrintNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��˷���
    '���:
    '   strPrintNO ��ӡ���ݺţ���ʽ��'A001','A002',...
    '����:��˳ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2011-06-23 09:59:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNOs As String
    Dim strNosData As String '��ʽ:A0001:1,2|A0002:1,2,3|....
    Dim str���ʱ�� As String
    
    strPrintNo = ""
    strNosData = GetSelectNOsAndSerialNum(strNOs)
     '���ʵĻ�,Ҫ���ñ���
    If Not zlAuditingWarn(mstrPrivs, strNOs, Val(nvl(mrsInfo!����ID))) Then Exit Function
    
    '�������
    str���ʱ�� = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    'strNos-������Ϣ, ��ʽ��NO1:���1,���2,...|NO1:���1,���2,...|...
    If mclsExpenceSvr.zlVerfyBillingPriceBill(Val("1-����"), strNosData, str���ʱ��) = False Then Exit Function
    
    'ҩƷ���շ�״̬ȷ��
    Call mclsExpenceSvr.zlDrugRecipeAffirm(strNOs, 1, 2)
    '�������շ�״̬ȷ��
    Call mclsExpenceSvr.zlStuffBillAffirm(strNOs, 1, 2, mbln�����Զ�����)
    
    strPrintNo = "'" & Replace(strNOs, ",", "','") & "'"
    
    VerifyFee = True
    
    '���ð�ҩ��
    Call mobjDrugStuff.DrugMachine_Charge(2, strNOs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCharge(ByRef strPrintNo As String, Optional ByRef blnPartialSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����շ�
    '���Σ�
    '   strPrintNO ��ӡ���ݺţ���ʽ��'A001','A002',...
    '   blnPartialSaved - �Ƿ񲿷ֱ���ɹ�
    '����:�շѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-23 11:38:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblThreeMoney As Double, dbl��Ԥ�� As Double, dbl���� As Double
    Dim strSQL As String, cllUpdate As Collection, cllOthers As Collection
    Dim str������ˮ�� As String, str����˵�� As String
    Dim str���㷽ʽ As String, j As Integer
    Dim blnTrans As Boolean, strExpend As String, dblOutMoney As Double
    Dim lng��������ID As Long, blnHaveMoney As Boolean
    Dim cll���㷽ʽ As Collection, i As Integer
    Dim bln���ڽ��� As Boolean, strErrMsg As String, blnCommit As Boolean
    Dim lng����ID  As Long, lng����ID As Long
 
    Err = 0: On Error GoTo ErrHandler
    strPrintNo = "": blnPartialSaved = False
    
    If mblnCliniqueRoomPay Then
        dblThreeMoney = mCurCharge.dbl��ǰδ��
    Else
        If txt��Ԥ��.Visible And txt��Ԥ��.Enabled Then dbl��Ԥ�� = Val(txt��Ԥ��.Text)
        If txt���.Visible And txt���.Enabled Then dblThreeMoney = Val(txt���.Text)
    End If
    dbl���� = mCurCharge.dbl��������
    
    blnTrans = True
    If SaveFeeBill() = False Then Exit Function
    
    lng����ID = mlngPatiID
    lng����ID = mCurCharge.lng����ID
    
    If RoundEx(dblThreeMoney, 6) = 0 Then
        'ȫ��ʹ��Ԥ����֧��
        strSQL = SetCurBalanceSQL(0, lng����ID, lng����ID, "", dbl��Ԥ��, mstr����IDs, dbl����, True)
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    Else
        'bytType-1-�����ӿ�֧��;2-���ѿ�֧��,0-����
        If mCurCardPay.bln���ѿ� Then
            If mcllSquareBalance Is Nothing Then Exit Function
            If mcllSquareBalance.Count = 0 Then Exit Function
            '�����ID|����|���ѿ�ID|���ѽ��||...
            '���ѿ�ID���Բ���,��Ϊ0ʱ,�Կ����Զ�����
            str���㷽ʽ = ""
            For j = 1 To mcllSquareBalance.Count
                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
                str���㷽ʽ = str���㷽ʽ & "||" & Val(mcllSquareBalance(j)(0))
                str���㷽ʽ = str���㷽ʽ & "|" & mcllSquareBalance(j)(3)
                str���㷽ʽ = str���㷽ʽ & "|" & Val(mcllSquareBalance(j)(1))
                str���㷽ʽ = str���㷽ʽ & "|" & Val(mcllSquareBalance(j)(2))
            Next
            If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 3)
            strSQL = SetCurBalanceSQL(2, lng����ID, lng����ID, str���㷽ʽ, dbl��Ԥ��, _
                mstr����IDs, dbl����, True, mCurCardPay.lng�����ID, mCurCardPay.strˢ������)
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Else '������֧��,Ԥ��������ѷ������֧��ʱ
            '���㷽ʽ|������|�������|����ժҪ
            If mCurCardPay.str֧������ = "" Then
                str���㷽ʽ = mCurCardPay.str���㷽ʽ
                str���㷽ʽ = str���㷽ʽ & "|" & dblThreeMoney
            Else
                str���㷽ʽ = mCurCardPay.str֧������
            End If
            str���㷽ʽ = str���㷽ʽ & "| |" & IIf(Trim(txtժҪ.Text) = "", " ", Trim(txtժҪ.Text))
            lng��������ID = zlDatabase.GetNextId("����Ԥ����¼")
            strSQL = SetCurBalanceSQL(3, lng����ID, lng����ID, str���㷽ʽ, 0, _
                "", 0, False, mCurCardPay.lng�����ID, mCurCardPay.strˢ������, lng��������ID)
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    End If
    
    If mblnCliniqueRoomPay = False Then
        If RoundEx(dblThreeMoney, 6) = 0 Or mCurCardPay.bln���ѿ� Then
            '����϶��ǳ�Ԥ������Ϊ���ѿ���ҽԺ�Ŀ��ʻ�
            gcnOracle.CommitTrans
            
            GoTo SuccessHandler:
            Exit Function
        End If
    End If
    
    If gblnAsyncCharge Then '���ý����첽���ƣ����ύ����
        gcnOracle.CommitTrans: blnTrans = False
        blnCommit = True
    End If
    
    If mobjThreeSwap.ExecutePay(mCurCardPay.lng�����ID, mCurCardPay.bln���ѿ�, _
        mCurCardPay.strˢ������, lng����ID, dblThreeMoney, str������ˮ��, str����˵��, _
        strExpend, dblOutMoney, cll���㷽ʽ, bln���ڽ���, strErrMsg, mCurCardPay.strQRCode) = False Then
        If blnTrans Then
            gcnOracle.RollbackTrans
            If strErrMsg <> "" Then ShowMsgbox strErrMsg
            Exit Function
        End If
        
        If bln���ڽ��� Then
            MsgBox IIf(strErrMsg = "", "", strErrMsg & vbCrLf) & _
                "    " & mCurCardPay.str���㷽ʽ & " ֧�����׳����쳣����ȷ�������Ƿ�ɹ����޷���ɽ��㡣" & vbCrLf & _
                "�뵽�շѴ��ڽ��д���", vbExclamation + vbOKOnly, gstrSysName
            blnPartialSaved = True: SaveCharge = True
            Exit Function
        Else
            If strErrMsg <> "" Then ShowMsgbox strErrMsg
            
            gcnOracle.BeginTrans: blnTrans = True
            '1.ɾ������Ԥ����¼
            'Zl_���˽����¼_Delete(
            strSQL = "Zl_���˽����¼_Delete("
            '  ����id_In     ����Ԥ����¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ","
            '  ��������id_In ����Ԥ����¼.��������id%Type
            strSQL = strSQL & "" & lng��������ID & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            
            '2.ɾ�����ý������,�ָ�Ϊ���۵�
            'Zl_�����շѽ���_Cancel(
            strSQL = "Zl_�����շѽ���_Cancel("
            '  ����id_In   ������ü�¼.����id%Type,
            strSQL = strSQL & "" & lng����ID & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            gcnOracle.CommitTrans: blnTrans = False
        End If
        
        Exit Function
    End If
    blnHaveMoney = RoundEx(dblThreeMoney, 6) <> RoundEx(dblOutMoney, 6)
    
    Set cllUpdate = New Collection
    If cll���㷽ʽ Is Nothing Then
        Call zlAddUpdateSwapSQL(False, lng����ID, mCurCardPay.lng�����ID, mCurCardPay.bln���ѿ�, _
            mCurCardPay.strˢ������, str������ˮ��, str����˵��, cllUpdate, 2)
    Else
        'Array("���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����|����",������ˮ��,����˵��)
        For i = 1 To cll���㷽ʽ.Count
            If Trim(Split(cll���㷽ʽ(i)(0), "|")(6)) <> "" Then mCurCardPay.strˢ������ = Split(cll���㷽ʽ(i)(0), "|")(6)
            strSQL = SetCurBalanceSQL(3, lng����ID, lng����ID, cll���㷽ʽ(i)(0), 0, _
                "", 0, False, mCurCardPay.lng�����ID, mCurCardPay.strˢ������, _
                lng��������ID, (i = 1), cll���㷽ʽ(i)(1), cll���㷽ʽ(i)(2), 2)
            zlAddArray cllUpdate, strSQL
        Next
    End If
    
    If blnTrans = False Then gcnOracle.BeginTrans: blnTrans = True
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, blnNoBeginTrans:=True
    blnTrans = False
    blnCommit = True
    
    Err = 0: On Error GoTo ErrOthers
    Set cllOthers = New Collection
    Call zlAddThreeSwapSQLToCollection(False, lng����ID, mCurCardPay.lng�����ID, mCurCardPay.bln���ѿ�, _
        mCurCardPay.strˢ������, strExpend, cllOthers)
    zlExecuteProcedureArrAy cllOthers, Me.Caption
    
ChargeOver:
    Err = 0: On Error GoTo ErrHandler
    If blnHaveMoney Then
        MsgBox "    " & mCurCardPay.str���㷽ʽ & " ʵ��֧�����(" & Format(dblOutMoney, "0.00") & ")������Ӧ�����(" & Format(dblThreeMoney, "0.00") & ")���޷���ɽ��㡣" & vbCrLf & _
            "�뵽�շѴ��ڽ��д���", vbExclamation + vbOKOnly, gstrSysName
        blnPartialSaved = True: SaveCharge = True
        Exit Function
    End If
    
    '��ɽ���
    strSQL = SetCurBalanceSQL(0, lng����ID, lng����ID, "", dbl��Ԥ��, mstr����IDs, dbl����, True)
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '�շѳɹ���Ĵ���
SuccessHandler:
    'ִ��ҩƷ���Ĵ�������״̬���£��Զ���ҩ������
    Call mclsExpenceSvr.zlDrugRecipeAffirm(Replace(mstrCurNos, "'", ""), 1, 1)
    Call mclsExpenceSvr.zlStuffBillAffirm(Replace(mstrCurNos, "'", ""), 1, 1)
    
    mlng����ID = lng����ID
    strPrintNo = "'" & Replace(mstrCurNos, ",", "','") & "'"
    SaveCharge = True
    
    '���ð�ҩ��
    Call mobjDrugStuff.DrugMachine_Charge(1, Replace(mstrCurNos, "'", ""))
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnCommit Then
        MsgBox IIf(Err.Description = "", "", Err.Description & vbCrLf) & _
            "    " & mCurCardPay.str���㷽ʽ & " ֧�����׳����쳣���޷���ɽ��㡣" & vbCrLf & _
            "�뵽�շѴ��ڽ��д���", vbExclamation + vbOKOnly, gstrSysName
        blnPartialSaved = True: SaveCharge = True
        Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrOthers:
    gcnOracle.CommitTrans   '�ܱ�����ٱ������
    Call ErrCenter
    GoTo ChargeOver:
End Function

Public Function GetBrushCardXMLExpend(ByVal strXMLExpend As String, ByRef dblOutMoney As Double, _
    ByRef strBalance As String) As Boolean
    '���ܣ���������֧��ˢ����֤����
    '��Σ�
    '   strXMLExpend:XML��
    '    <OUTPUT>
    '        <JS> //������Ϣ(Ŀǰֻ֧�ַ���һ�ַ�ʽ)
    '            <JYFS>���׷�ʽ</JYFS> //���׷�ʽ:�����㷽ʽ.����
    '            <JYJE>���׽��</JYJE>
    '        </JS>
    '        ...
    '    </OUTPUT>
    '���Σ�
    '   dblOutMoney - ʵ��֧�����
    '   strBalance - �������ݣ���ʽ�����㷽ʽ|������||...
    Dim lngCount As Long, strValue As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    dblOutMoney = 0
    strBalance = ""
    If zlXML_Init() = False Then Exit Function
    If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
    '������Ϣ
    Call zlXML_GetRows("JS", lngCount)
    For i = 0 To lngCount - 1
        Call zlXML_GetNodeValue("JYFS", i, strValue)
        strBalance = strBalance & "||" & strValue '���㷽ʽ
        Call zlXML_GetNodeValue("JYJE", i, strValue)
        strBalance = strBalance & "|" & Val(strValue) '������
        dblOutMoney = dblOutMoney + Val(strValue)
    Next
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    GetBrushCardXMLExpend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData(ByRef strPrintNo As String, Optional ByRef blnPartialSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:
    '   strPrintNO ��ӡ���ݺţ���ʽ��'A001','A002',...
    '����:���˺�
    '����:2011-06-22 16:01:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '1-�շѼ�¼;2-���ʼ�¼
    
    Select Case mbytCurType
    Case 1  '�շѻ��۴���
        If SaveCharge(strPrintNo, blnPartialSaved) = False Then Exit Function
        '��ӡ��ص�Ʊ��
    Case 2 '���ۼ������
        If VerifyFee(strPrintNo) = False Then Exit Function
        SaveData = True
    Case Else
        Exit Function
    End Select
    
    SaveData = True
End Function

Private Sub cmdPrintSet_Click()
    If frmSquareAffirmParaSet.SetPara(Me) = False Then Exit Sub
    Call InitFactPara
End Sub

Private Function SetCurBalanceSQL(ByVal bytType As Byte, ByVal lng����ID As Long, ByVal lng����ID As Long, _
    ByVal str���㷽ʽ As String, ByVal dbl��Ԥ�� As Double, ByVal str����IDs As String, _
    ByVal dbl�������� As Double, ByVal bln��ɽ��� As Boolean, _
    Optional ByVal lngCardTypeID As Long, Optional ByVal str���� As String, _
    Optional ByVal lng��������ID As Long, Optional ByVal blnɾ��ԭ���� As Boolean, _
    Optional ByVal str������ˮ�� As String, Optional ByVal str����˵�� As String, _
    Optional ByVal bytУ�Ա�־ As Byte = 1) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ�ǰ�����SQL��cllpro����
    '���:  bytType-1-�����ӿ�֧��;2-���ѿ�֧��;3-�����ӿڶ��ֽ��㷽ʽ֧��;0-����
    '       dbl��Ԥ��-Ԥ����֧��
    '       dbl��������-���β���������
    '����:
    '����:���óɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL  As String
    
    On Error GoTo errHandle
    ' Zl_�����շѽ���_Modify
    strSQL = "Zl_�����շѽ���_Modify("
    '  ------------------------------------------------------------------------------------------------------------------------------
    '  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
    '  --��������_In:
    '  --   0-��ͨ�շѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
    '  --   1.����������:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����֧Ʊ��_In:������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����֧Ʊ��_In:������
    '  --   3-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ
    '  --     ����֧Ʊ��_In:������
    '  --   4-���������㣬���ֽ��㷽ʽ:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ|���ݺ�|�Ƿ���ͨ����"
    '  --     ����֧Ʊ��_In:������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  -- ��Ԥ��_In: ���ڳ�Ԥ��ʱ,����
    '  -- �����_In:��������ʱ,����
    '  -- ��ɽ���_In:1-����շ�;0-δ����շ�
    '  ------------------------------------------------------------------------------------------------------------------------------
    Select Case bytType
    Case 1  '1-�����ӿ�֧��
        strSQL = strSQL & "1" & ","
    Case 2 ' 2-���ѿ�֧��
        strSQL = strSQL & "3" & ","
    Case 3 ' 3-�����ӿڶ��ֽ��㷽ʽ֧��
        strSQL = strSQL & "4" & ","
    Case Else
        strSQL = strSQL & "0" & ","
    End Select
    '    ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & lng����ID & ","
    '    ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & lng����ID & ","
    '    ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '    ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & ZVal(dbl��Ԥ��) & ","
    '    ��֧Ʊ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '    �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & ZVal(lngCardTypeID) & ","
    '    ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "" & IIf(lngCardTypeID = 0, "NULL", "'" & str���� & "'") & ","
    '    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & str������ˮ�� & "',"
    '    ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & str����˵�� & "',"
    '    �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '    �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '    �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '    -- �����_In:��������ʱ,����
    strSQL = strSQL & "" & dbl�������� & ","
    '    ��ɽ���_In Number:=0
    '    -- ��ɽ���_In:1-����շ�;0-δ����շ�
    strSQL = strSQL & IIf(bln��ɽ���, "1", "0") & ","
    '  ȱʡ���㷽ʽ_In  ���㷽ʽ.����%Type := Null,
    strSQL = strSQL & "NULL" & ","
    '79868,Ƚ����,2015-06-10,ʹ�ò��˼���Ԥ��
    '  ��Ԥ������ids_In Varchar2:=Null,
    strSQL = strSQL & "'" & lng����ID & "," & str����IDs & "',"
    '  ���½������_In  Number := 1,
    strSQL = strSQL & "" & 1 & ","
    '  ��������id_In    ����Ԥ����¼.��������id%Type := Null
    strSQL = strSQL & "" & ZVal(lng��������ID) & ","
    '  ɾ��ԭ����_In    Number := 0
    strSQL = strSQL & "" & IIf(blnɾ��ԭ����, "1", "0") & ","
    '  У�Ա�־_In      ����Ԥ����¼.У�Ա�־%Type := 0
    strSQL = strSQL & "" & bytУ�Ա�־ & ")"
    SetCurBalanceSQL = strSQL
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ������Ϣ�Լ��������
    '����:���ϴ�
    '����:2014-10-29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String, lngPatient As Long
    
    On Error GoTo Errhand
    If Not gblnLED Then Exit Sub
    
    zl9LedVoice.Reset mscCom
    strInfo = nvl(mrsInfo!����) & " " & nvl(mrsInfo!�Ա�) & " " & nvl(mrsInfo!����)
    lngPatient = Val("" & mrsInfo!����ID)
    zl9LedVoice.DisplayPatient strInfo, lngPatient
    
    '�����ܶ�:������Ҫ֧���Ľ�Ԥ�����:���˵�ǰ��Ԥ�����
    Call zl9LedVoice.DisplayBank("�����ܶ�:" & mCurCharge.dbl���úϼ� & "Ԫ" & _
        IIf(mCurCharge.dblԤ����� = 0, "", ",Ԥ�����:" & mCurCharge.dblԤ����� & "Ԫ"))
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function InitThreeSwap(frmMain As Object) As Boolean
    '��ʼ����֧������
    On Error GoTo ErrHandler
    If Not mobjThreeSwap Is Nothing Then InitThreeSwap = True: Exit Function
    
    Set mobjThreeSwap = New clsThreeSwap
    mobjThreeSwap.Init frmMain, mlngModule, nvl(mrsInfo!����ID), nvl(mrsInfo!����), nvl(mrsInfo!�Ա�), nvl(mrsInfo!����)
    InitThreeSwap = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPayDrugWindow(ByVal lng����ID As Long, ByVal dt�շ�ʱ�� As Date, _
    ByVal cllDept As Collection) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ����䷢ҩ����
    '���:lng����ID-����ID
    '     dt�շ�ʱ��-�շ�ʱ��
    '     cllDept-����ִ�в���:array(�շ����,ִ�в���ID,��ҩ����)
    '���أ���ҩ��������
    '���ƣ����ϴ�
    '���:strNO
    'ʱ�䣺2014-6-12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��ҩ���� As String, strPayDrugWins As String
    Dim i As Long, varData As Variant
    Dim blnFirst As Boolean
    
    On Error GoTo ErrHandler
    blnFirst = True
    strPayDrugWins = ""
    For i = 1 To cllDept.Count
        varData = cllDept(i)
        str��ҩ���� = varData(2)
        If str��ҩ���� = "" Then
            str��ҩ���� = mobjDrugStuff.Get��ҩ����(lng����ID, Trim(varData(0)), Val(varData(1)), blnFirst)
            If blnFirst Then blnFirst = False
        End If
        If InStr(1, strPayDrugWins & ";", ";" & Val(varData(1)) & "|") = 0 Then
            strPayDrugWins = strPayDrugWins & ";" & Val(varData(1)) & "|" & str��ҩ����
        End If
    Next
    GetPayDrugWindow = strPayDrugWins
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function WriteInforToCard(frmMain As Object, ByVal lngModul As Long, ByVal strPrivs As String, _
      ByVal lngCardTypeID As Long, ByVal strNOs As String) As Boolean
    '����:��������Ϣд�뿨��
    '��Σ�
    '    frmMain - ���ô���
    '    lngModul - ģ���
    '    strPrivs - Ȩ�޴�
    '    objSquareCard - ҽ�ƿ�����
    '    strNOs - ���ݺţ���ʽ��'A0001','A0002','A0003',...��A0001,A0002,A0003,...
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim lng����ID As Long, lng������� As Long
    
    Err = 0: On Error GoTo errH:
    '����:56615
    'If InStr(strPrivs, ";������Ϣд��;") = 0 Then Exit Function
    
    strSQL = "Select /*+Cardinality(j,10)*/ Distinct A.����ID,B.�������" & _
            " From ������ü�¼ A,����Ԥ����¼ B,Table(f_Str2list([1])) J" & _
            " Where A.����ID=B.����ID And A.NO=J.Column_Value And A.��¼���� = 1 And A.��¼״̬ in(1,3)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ݽ������", Replace(strNOs, "'", ""))
    If rsTemp.EOF Then Exit Function
    
    Do While Not rsTemp.EOF
        lng����ID = Val(nvl(rsTemp!����ID))
        lng������� = Val(nvl(rsTemp!�������))
        '���ý�����д���ӿ�
        If lng����ID <> 0 And lng������� <> 0 Then
            Call gobjOneCardComLib.zlMzInforWriteToCard(frmMain, lngModul, lngCardTypeID, lng����ID, lng�������)
        End If
        rsTemp.MoveNext
    Loop
    
    WriteInforToCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtժҪ_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtժҪ
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlControl.ControlSetFocus cmdOK
End Sub

Private Sub txtժҪ_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    On Error GoTo ErrHandler
    With vsBalance
        If mInsure.intInsure = 0 Or mInsurePara.�൥�ݷֵ��ݽ��� Then Cancel = True: Exit Sub
        If mblnYbBalanced Then Cancel = True: Exit Sub
        
        If Row < .FixedRows Or Col < .FixedCols Then Cancel = True: Exit Sub
        If .TextMatrix(Row, .ColIndex("֧����ʽ")) = "" Then Cancel = True: Exit Sub
        If Col <> .ColIndex("֧�����") Then Cancel = True: Exit Sub
        
        '�������޸ĵ�ҽ����Ŀ
        If Val(.RowData(Row)) = 0 Then Cancel = True: Exit Sub
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsBalance_EnterCell()
    If vsBalance.Editable = flexEDNone Then Exit Sub
    vsBalance.EditCell
End Sub

Private Sub vsBalance_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Call Grid.CheckKeyPress(vsBalance, Row, Col, KeyAscii, m�����ʽ)
End Sub

Private Sub vsBalance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strMoney As String, str֧����ʽ As String
    Dim curOrig As Currency, curTotal As Currency
    Dim p As Integer, objItem As BalanceMoney
    
    On Error GoTo ErrHandler
    With vsBalance
        If Row < 0 Then Exit Sub
        If Col <> 1 Or Col < 0 Then Exit Sub
        
        If zlCommFun.DblIsValid(.EditText, 10, False, False) = False Then Cancel = True: Exit Sub
        .EditText = Format(Val(.EditText), "0.00")
            
        strMoney = Trim(.EditText)
        If Not IsNumeric(strMoney) Then
            ShowMsgbox "�����˷Ƿ���֧����""" & strMoney & """��"
            .EditCell
            .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True: Exit Sub
        End If
        
        str֧����ʽ = Trim(.TextMatrix(.Row, .ColIndex("֧����ʽ")))
        If str֧����ʽ = "" Then Exit Sub
        
        If str֧����ʽ = mstr�����ʻ� Then '�����ʻ����
            '������������͸֧���
            If Val(strMoney) > mInsure.dbl������� + mInsure.dbl����͸֧ Then
                ShowMsgbox "�ʻ����:" & Format(mInsure.dbl�������, "0.00") & _
                    IIf(mInsure.dbl����͸֧ = 0, "", "(" & "����͸֧:" & Format(mInsure.dbl����͸֧, "0.00") & ")") & _
                    "����Ҫ֧���Ľ�"
                .EditCell
                .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True: Exit Sub
            End If
        Else
            '��������������ص�ԭʼ���(�����ʻ�����͸֧ʱ���ж�)
            curOrig = GetMedicareSum(mInsure.colBalance, str֧����ʽ, , True)   '�ý��㷽ʽ����ԭʼ���ؽ���
            If Val(strMoney) > curOrig Then
                ShowMsgbox "�����""" & .TextMatrix(Row, 0) & """֧�����ܳ��� " & Format(curOrig, "0.00") & " ��"
                .EditCell
                .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
                Cancel = True: Exit Sub
            End If
        End If
        
        '������������ʣ��ɽ�����
        curTotal = mCurCharge.dbl��ǰδ��
        For p = 1 To mInsure.colBalance.Count
            For Each objItem In mInsure.colBalance(p)
                If objItem.���㷽ʽ <> str֧����ʽ Then
                    curTotal = curTotal - objItem.��Ч���
                End If
            Next
        Next
        If Val(strMoney) > curTotal Then
            ShowMsgbox "֧�������󣬳�����������֧�����:" & Format(curTotal, "0.00") & "��"
            .EditCell
            .EditSelStart = 0: .EditSelLength = zlCommFun.ActualLen(.EditText)
            Cancel = True: Exit Sub
        End If
        
        Call SetBalanceVal(mInsure.colBalance, 1, str֧����ʽ & "|" & CCur(Val(strMoney)))
    End With
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetCurCard(ByRef objCard As Card) As Boolean
    '��ȡ��ǰ��
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    GetCurCard = True
    Exit Function
errHandle:
    Set objCard = New Card
End Function

Private Sub RestorePrePayTypeFromTag()
    '�ָ����ϴ�ѡ���֧����
    '˵��:cbo֧����ʽ.Tag�洢�����ϴ�ѡ���֧��������,��ʽ:Index:�ɿ���
    Dim varTemp As Variant, intIndex As Integer
    
    On Error GoTo ErrHandler
    mCurCardPay.strQRCode = ""
    If cbo֧����ʽ.Tag = "" Then Exit Sub
    
    '���ϴ�ѡ��Ŀ����ID,�ָ�
    varTemp = Split(cbo֧����ʽ.Tag & ":", ":")
    cbo֧����ʽ.Tag = ""
    
    intIndex = Val(varTemp(0))
    cbo֧����ʽ.ListIndex = intIndex
    txt���.Text = varTemp(1)
    zlControl.ControlSetFocus txt���
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetPayTypeFromCardTypeID(ByVal lngCardTypeID As Long, Optional ByVal bln���ѿ� As Boolean) As Boolean
    '���ݿ����ID,��λ��ָ����֧�������
    '���:
    '   lngCardTypeID-�����ID
    '   bln���ѿ�-�Ƿ����ѿ�
    '   blnOnlyChangePayType '�Ƿ���ı�֧�����
    '����:��λ�ɹ�����true,���򷵻�False
    Dim objCard As Card, blnFind As Boolean, i As Integer
    Dim intIndex As Integer
    
    On Error GoTo ErrHandler
    If lngCardTypeID <= 0 Then Exit Function
    For i = 1 To mobjPayCards.Count
        Set objCard = mobjPayCards(i)
        If objCard.�ӿ���� = lngCardTypeID And objCard.���ѿ� = bln���ѿ� Then intIndex = i: Exit For
    Next
    If intIndex = 0 Then Exit Function
    
    '�����ID��������Ч��֧��ɨ�븶�Ŀ������
    If InStrEx(btQRCodePay.Tag, lngCardTypeID) = False Then Exit Function
    
    With cbo֧����ʽ
        For i = 0 To .ListCount - 1
            If .ItemData(i) = intIndex Then
                .ListIndex = i
                blnFind = True: Exit For
            End If
        Next
    End With
    SetPayTypeFromCardTypeID = blnFind
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub btQRCodePay_zlErrShow(ByVal strErrMsg As String, ByVal lngErrNum As Long)
    On Error GoTo ErrHandler
    Call RestorePrePayTypeFromTag '�ָ��ϴ�ѡ����
    If strErrMsg = "" Then Exit Sub
    ShowMsgbox strErrMsg
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub btQRCodePay_zlGetPayMoney(dblMoney As Double, strExpend As String, blnCancel As Boolean)
    Dim varTemp As Variant
    
    On Error GoTo ErrHandler
    cbo֧����ʽ.Tag = cbo֧����ʽ.ListIndex & ":" & txt���.Text '�ȼ�¼ԭ֧����Ϣ
    zlControl.ControlSetFocus txt���
    Call txt���_GotFocus
    
    '��λ��ָ�������
    dblMoney = Val(txt���.Text)
    varTemp = Split(btQRCodePay.Tag & ",", ",") '�洢����Ч�Ŀ����IDs
    If SetPayTypeFromCardTypeID(Val(varTemp(0))) = False Then
        ShowMsgbox "������ָ����ɨ�븶������飡"
        blnCancel = True
        Call RestorePrePayTypeFromTag '�ָ��ϴ�ѡ����
        Exit Sub
    End If
    
    '��ȡ����֧�����
    txt���.Text = Format(dblMoney, "0.00")
    
    If RoundEx(dblMoney, 6) <= 0 Then
        If RoundEx(dblMoney, 6) = 0 Then
            ShowMsgbox "����δ�����Ϊ�㣬����Ҫ����ɨ�븶�"
        Else
            ShowMsgbox "��ǰΪ�˿ɨ�븶��֧���˿������"
        End If
        blnCancel = True
        Call RestorePrePayTypeFromTag '�ָ��ϴ�ѡ����
        Exit Sub
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    blnCancel = True
    Call RestorePrePayTypeFromTag '�ָ��ϴ�ѡ����
End Sub

Private Sub btQRCodePay_zlQRCodePayment(ByVal lngCardTypeID As Long, ByVal strPayMentQRCode As String, ByVal strExpendXML As String, blnCancel As Boolean)
    '����ɨ�븶��
    '���:
    '   lngCardTypeID-�����ID
    '   strPayMentQRCode-��ά�븶������
    '   strExpendXML-����
    '����:strExpendXML-����
    '     blnCancel-true��ʾȡ������ɨ�븶,False-��ʾ����ɨ�븶�ɹ�
    
    On Error GoTo ErrHandler
    If lngCardTypeID = 0 Or blnCancel Then
        blnCancel = True
        Call RestorePrePayTypeFromTag '�ָ��ϴ�ѡ����
        Exit Sub
    End If
    
    blnCancel = False
    If SetPayTypeFromCardTypeID(lngCardTypeID, False) = False Then    '��λ��ɨ�븶��ָ�������
        ShowMsgbox "������Чʶ��ǰɨ�븶����𣬿��ܱ�����֧�ָ�����ɨ�븶���������Ա��ϵ��"
        blnCancel = True
        Call RestorePrePayTypeFromTag '�ָ��ϴ�ѡ����
        Exit Sub
    End If
    
    mCurCardPay.strQRCode = strPayMentQRCode
    Call cmdOK_Click
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
    blnCancel = True
    Call RestorePrePayTypeFromTag '�ָ��ϴ�ѡ����
End Sub

