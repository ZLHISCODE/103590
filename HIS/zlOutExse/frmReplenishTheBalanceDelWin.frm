VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReplenishTheBalanceDelWin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������˷���Ϣ"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9405
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReplenishTheBalanceDelWin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame3 
      Height          =   60
      Left            =   -30
      TabIndex        =   29
      Top             =   1230
      Width           =   10260
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -690
      TabIndex        =   25
      Top             =   5085
      Width           =   10635
   End
   Begin VB.TextBox txtAge 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   7905
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   150
      Width           =   1185
   End
   Begin VB.TextBox txtSex 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   150
      Width           =   1185
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7755
      TabIndex        =   20
      Top             =   5265
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -30
      TabIndex        =   24
      Top             =   660
      Width           =   10260
   End
   Begin VB.TextBox txtPatiName 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1365
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   2895
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   3675
      Left            =   75
      ScaleHeight     =   3675
      ScaleWidth      =   9300
      TabIndex        =   22
      Top             =   1290
      Width           =   9300
      Begin VB.PictureBox PicBalanceBack 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   3150
         Left            =   45
         ScaleHeight     =   3120
         ScaleWidth      =   4020
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   480
         Width           =   4050
         Begin VSFlex8Ctl.VSFlexGrid vsBalance 
            Height          =   2430
            Left            =   90
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   90
            Width           =   3810
            _cx             =   6720
            _cy             =   4286
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   14.25
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
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   460
            RowHeightMax    =   500
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmReplenishTheBalanceDelWin.frx":6852
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
         Begin VB.Label lblYBMoney 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   1350
            TabIndex        =   30
            Top             =   2640
            Width           =   2565
         End
         Begin VB.Label lblҽ���ϼ� 
            AutoSize        =   -1  'True
            Caption         =   "ҽ���ϼ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   8
            Top             =   2685
            Width           =   1140
         End
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   3150
         Left            =   4290
         ScaleHeight     =   3120
         ScaleWidth      =   4875
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   480
         Width           =   4905
         Begin VB.TextBox txtԤ��� 
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
            Left            =   1350
            MaxLength       =   12
            TabIndex        =   11
            Top             =   2430
            Visible         =   0   'False
            Width           =   3210
         End
         Begin VB.TextBox txtժҪ 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   1350
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   1845
            Width           =   3210
         End
         Begin VB.TextBox txt�ɿ� 
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
            Left            =   1350
            MaxLength       =   12
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   675
            Width           =   3210
         End
         Begin VB.TextBox txt������� 
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   1350
            MaxLength       =   30
            TabIndex        =   17
            Top             =   1260
            Width           =   3210
         End
         Begin VB.ComboBox cbo֧����ʽ 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   405
            Left            =   1350
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   120
            Width           =   3210
         End
         Begin VB.Label lblԤ��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ԥ���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   2513
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label lblժҪ 
            AutoSize        =   -1  'True
            Caption         =   "ժ  Ҫ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   420
            TabIndex        =   18
            Top             =   1815
            Width           =   870
         End
         Begin VB.Label lbl�˿��� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�˿���"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   14
            Top             =   750
            Width           =   1200
         End
         Begin VB.Label lbl������� 
            AutoSize        =   -1  'True
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   16
            Top             =   1335
            Width           =   1140
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�˿ʽ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   12
            Top             =   195
            Width           =   1200
         End
      End
      Begin XtremeSuiteControls.ShortcutCaption stcBalanceTittle 
         Height          =   405
         Left            =   4275
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   30
         Width           =   4920
         _Version        =   589884
         _ExtentX        =   8678
         _ExtentY        =   714
         _StockProps     =   6
         Caption         =   "��ǰ������Ϣ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         GradientColorLight=   16711680
         GradientColorDark=   16711680
      End
      Begin XtremeSuiteControls.ShortcutCaption stcYbTittle 
         Height          =   405
         Left            =   45
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   45
         Width           =   4065
         _Version        =   589884
         _ExtentX        =   7170
         _ExtentY        =   714
         _StockProps     =   6
         Caption         =   "��ǰ�˷���Ϣ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   510
      Left            =   0
      TabIndex        =   28
      Top             =   5850
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   900
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2619
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceDelWin.frx":68A0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8361
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1138
            MinWidth        =   1146
            Object.Tag             =   "�����շ�Ԥ�������ʾ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1164
            MinWidth        =   1162
            Object.Tag             =   "�����շ�������������ʾ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmReplenishTheBalanceDelWin.frx":7134
            Key             =   "Calc"
            Object.ToolTipText     =   "������:ALT+?"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1111
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�(&E)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6195
      TabIndex        =   21
      Top             =   5265
      Width           =   1470
   End
   Begin VB.Label lbl�˷Ѻϼ� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1365
      TabIndex        =   32
      Top             =   810
      Width           =   2655
   End
   Begin VB.Label lbl�ϼ� 
      Caption         =   "�˷Ѻϼ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   31
      Top             =   855
      Width           =   1230
   End
   Begin VB.Label lbl��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������:0.00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5265
      TabIndex        =   27
      Top             =   5415
      Width           =   2025
   End
   Begin VB.Label lblAge 
      AutoSize        =   -1  'True
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7260
      TabIndex        =   4
      Top             =   210
      Width           =   570
   End
   Begin VB.Label lblSex 
      AutoSize        =   -1  'True
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5010
      TabIndex        =   2
      Top             =   210
      Width           =   570
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   1140
   End
End
Attribute VB_Name = "frmReplenishTheBalanceDelWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------
'���������ر���
Public Enum gEM_BalanceDel
    EM_BalanceDel = 0   '�����˷�
    EM_BalanceReDel = 1  '�����˷�
End Enum
Private mbytFunc As gEM_BalanceDel
Private mobjDelBalance As clsCliniDelBalance
Private mfrmMain As Object
Private mlngModule As Long, mstrPrivs As String
Private mblnҽ���ֱ� As Boolean
Private mcllForceDelToCash As Collection 'ǿ��������Ϣ��Array(����Ա,���������,���㷽ʽ)
Private mstrDefaultBalance As String
Private mstr�ų����㷽ʽ As String '����ʹ�õĽ��㷽ʽ,����ö��ŷָ�
Private mblnRegister As Boolean
'------------------------------------------------------------------------------------------
'�ֲ�����
Private mobjPayCards As Cards
Private mstrTittle As String
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean '�Ƿ�Unload����
Private mdblDelMoney As Double '�����˿���
Private mdbl��ǰδ�� As Double
Private mdbl����Ԥ�� As Double
Private mblnOK As Boolean, mblnNotClick As Boolean
Private mbln�ѱ��� As Boolean
Private mlngR As Long
Private Type TY_BrushCard    'ˢ������
    str���� As String
    str���� As String
    str������ˮ�� As String    '������ˮ��
    str����˵��  As String     '������Ϣ
    str��չ��Ϣ As String    '���׵���չ��Ϣ
End Type
Private mCurBrushCard As TY_BrushCard   '��ǰ��ˢ����Ϣ
Private Enum Pan
    C2��ʾ��Ϣ = 2
    C3�����ʻ� = 3
    C4�����ʻ���Ϣ = 4
End Enum

'------------------------------------------------------------------------------------------
'API����
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:�Ƿ񻺴��˻س���,���ܴ������շѽ���ˢ���б�������˻س�,�����Ҫ�ж�
Private mlngPre֧����ʽ As Long
Private mrsOldBalance As ADODB.Recordset
Private mblnThreeSwapSingle As Boolean '�Ƿ񵥶������˷ѽӿ�

Public Function zlChargeWin(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal strPrivs As String, ByVal bytFunc As gEM_BalanceDel, _
    ByVal objPayCards As Cards, ByVal objDelBalance As clsCliniDelBalance, _
    ByVal blnҽ���ֱ� As Boolean, _
    Optional ByVal cllForceDelToCash As Collection, Optional ByVal str�ų����㷽ʽ As String, _
    Optional ByVal blnRegister As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������:��ʾ����֧�����㴰��
    '���:frmMain-���õ�������
    '       lngModule -ģ���
    '       strPrivs-Ȩ�޴�
    '       str�������:���ν������
    '       blnҽ���ֱ�-ҽ���Ƿ�ֱҴ���
    '       dtDate-��ǰ�շ�ʱ��
    '      objPayCards-��ǰ��Ч��֧�����
    '       cllForceDelToCash - ǿ��������Ϣ��Array(����Ա,���������,���㷽ʽ)
    '       str�ų����㷽ʽ - ����ʹ�õĽ��㷽ʽ,����ö��ŷָ�
    '       blnRegister - �Ƿ��ǹҺŽ��㵥��
    '����:��ɽ���,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 14:33:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mfrmMain = frmMain: mbytFunc = bytFunc
    mstrPrivs = strPrivs: mlngModule = lngModule
    mblnҽ���ֱ� = blnҽ���ֱ�
    Set mobjDelBalance = objDelBalance
    Set mobjPayCards = objPayCards
    Call InitVar '��ʼ����ر���ģ�����
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    Set mcllForceDelToCash = cllForceDelToCash
    mstr�ų����㷽ʽ = str�ų����㷽ʽ
    mblnRegister = blnRegister
    
    Me.Show 1, frmMain
    zlChargeWin = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���ģ�����
    '����:���˺�
    '����:2014-09-18 17:16:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNotClick = False:  mblnUnLoad = False
    mblnOK = False
    mblnFirst = True
    mstrDefaultBalance = ""
End Sub

Private Sub InitBalanceGrid(Optional ByVal blnInitColHead As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ս�����
    '���:blnInitColHead-����ʼ����ͷ
    '����:���˺�
    '����:2014-09-18 14:06:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    With vsBalance
        .Clear 1
        .Rows = 4
        If blnInitColHead Then
            .COLS = 2
            .TextMatrix(0, 0) = "���㷽ʽ"
            .TextMatrix(0, 1) = "֧�����"
            For i = 0 To .COLS - 1
                .ColKey(i) = .TextMatrix(0, i)
                .FixedAlignment(i) = flexAlignCenterCenter
                If .ColKey(i) Like "*���" Then
                    .ColAlignment(i) = flexAlignRightCenter
                Else
                    .ColAlignment(i) = flexAlignLeftCenter
                End If
            Next
            .ColWidth(.ColIndex("���㷽ʽ")) = (vsBalance.Width - 300) * 0.6
            .ColWidth(.ColIndex("֧�����")) = (vsBalance.Width - 300) * 0.4
            .Row = 0: .Col = 1
        End If
        .TabStop = False
        .Cell(flexcpFontBold, 0, 0, .Rows - 1, .COLS - 1) = False
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .COLS - 1) = Me.ForeColor
    End With
End Sub

Private Function LoadData(ByVal str������� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؽ�������
    '���:str�������-�������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 14:22:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Integer, dblYbMoney As Double
    Dim dblDelMoney As Double '�����˿���
    Dim dblDelAllMoney As Double '�˷Ѻϼ�
    
    mdblDelMoney = 0
    mdbl��ǰδ�� = 0
    
    On Error GoTo errHandle
    
    Call InitBalanceGrid
    strSQL = "" & _
    "   Select decode(B.����,null ,0,1) as ҽ��,A.���㷽ʽ,sum(A.��Ԥ��) as ��Ԥ��  " & _
    "   From ����Ԥ����¼ A,(select ���� From ���㷽ʽ where ���� in (3,4)) B" & _
    "   Where  A.������� = [1] and a.���㷽ʽ=b.����(+)" & _
    "   Group by decode(B.����,null ,0,1),A.���㷽ʽ" & _
    "   Order by ҽ�� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�������)
    vsBalance.Appearance = flexFlat
    With rsTemp
        i = 1
        Do While Not .EOF
            If Nvl(rsTemp!���㷽ʽ) <> "" Then
                With vsBalance
                    If .TextMatrix(i, .ColIndex("���㷽ʽ")) <> "" Then
                        .Rows = .Rows + 1
                        i = i + 1
                    End If
                    .RowData(i) = 0
                    .TextMatrix(i, .ColIndex("���㷽ʽ")) = Nvl(rsTemp!���㷽ʽ)
                    .TextMatrix(i, .ColIndex("֧�����")) = Format(-1 * Val(Nvl(rsTemp!��Ԥ��)), "0.00")
                End With
                
                If Val(Nvl(rsTemp!ҽ��)) = 1 Then
                    dblYbMoney = dblYbMoney + Val(Nvl(rsTemp!��Ԥ��))
                End If
            Else
                dblDelMoney = dblDelMoney + Val(Nvl(rsTemp!��Ԥ��))
            End If
            dblDelAllMoney = dblDelAllMoney + Val(Nvl(rsTemp!��Ԥ��))
            .MoveNext
        Loop
    End With
    lblYBMoney.Caption = Format(-1 * dblYbMoney, "0.00")
    lbl�˷Ѻϼ�.Caption = Format(-1 * dblDelAllMoney, "0.00")
    '���㱾��ʵ���˿�
    mdblDelMoney = RoundEx(dblDelMoney, 6)
    txt�ɿ�.Text = Format(-1 * mdblDelMoney, "0.00")
    LoadData = True
   Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ؼ�
    '����:���˺�
    '����:2011-06-13 14:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurBrushCard As TY_BrushCard
    
    zlControl.PicShowFlat PicBalanceBack, -1, , taCenterAlign
    zlControl.PicShowFlat picPay, -1, , taCenterAlign
    
    Call InitBalanceGrid(True)
    If mblnUnLoad = False Then
        mblnUnLoad = Not LoadData(mobjDelBalance.�������)
        Set mrsOldBalance = zlFromIDGetChargeBalance(2, mobjDelBalance.AllNos, , , , IIf(mblnRegister, 4, 1))
        '��ȡ�շ�ʱʹ�õĽ��㷽ʽ��Ϊȱʡ���㷽ʽ
        mrsOldBalance.Filter = "�˷�=0"
        If mrsOldBalance.EOF = False Then
            If Val(mrsOldBalance!����) = 1 Then
                mstrDefaultBalance = "Ԥ���"
            Else
                mstrDefaultBalance = Nvl(mrsOldBalance!���㷽ʽ)
            End If
        End If
        mrsOldBalance.Filter = ""
    End If
    mdbl��ǰδ�� = mdblDelMoney
    Call Load֧����ʽ
    txt�ɿ�.Text = Format(-1 * mdbl��ǰδ��, "0.00")
    If mdbl��ǰδ�� <= 0 Then
        lblPayType.Caption = "�˿ʽ"
        lbl�˿���.Caption = "�˿���"
        txt�ɿ�.ForeColor = lblPati.ForeColor
    Else
        lblPayType.Caption = "�տʽ"
        lbl�˿���.Caption = "�տ���"
        txt�ɿ�.ForeColor = vbRed
    End If
    
    txt�ɿ�.Locked = True
    txt�ɿ�.BackColor = &HE0E0E0

    mCurBrushCard = CurBrushCard
    stbThis.Panels(C4�����ʻ���Ϣ).Text = "": stbThis.Panels(C4�����ʻ���Ϣ).ToolTipText = ""
    stbThis.Panels(C4�����ʻ���Ϣ).Visible = False
    vsBalance.BackColor = Me.BackColor
    vsBalance.BackColorBkg = Me.BackColor
    txtPatiName.Text = mobjDelBalance.����
    txtPatiName.ForeColor = vbRed
    If mobjDelBalance.�������� <> "" Then
         Call SetPatiColor(txtPatiName, mobjDelBalance.��������, vbRed)
    End If
    txtAge.Text = mobjDelBalance.����
    txtAge.ForeColor = txtPatiName.ForeColor
    txtSex.Text = mobjDelBalance.�Ա�
    txtSex.ForeColor = txtPatiName.ForeColor
End Sub
Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim objCard As Card, objCards As Cards, objPayCards As Cards
    Dim lngKey As Long, i As Long, j As Long
    Dim varData As Variant
    
    On Error GoTo errHandle
    
'    If mobjPayCards Is Nothing Then
        Set objCards = New Cards: Set mobjPayCards = New Cards
        Set rsTemp = Get���㷽ʽ("������")
        '83533:���ϴ�,2015/3/25,û����Ч�Ĳ�����
        If rsTemp.RecordCount = 0 Then
            MsgBox "������û�п��õĽ��㷽ʽ�����ȵ������㷽ʽ���������ò������Ӧ�ó��ϡ�", vbInformation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
        If Not gobjSquare Is Nothing Then
            ' zlGetCards(ByVal BytType As Byte)
            '���:bytType-0-����ҽ�ƿ�;
            '             1-���õ�ҽ�ƿ�,
            '             2-���д��������˻���������
            '             3-���õ������˻���ҽ�ƿ�
           Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
        End If
        
        With rsTemp
            .Filter = 0
            If .RecordCount <> 0 Then .MoveFirst
            lngKey = 1
            Do While Not .EOF
                For i = 1 To objCards.Count
                    If objCards(i).���㷽ʽ = Nvl(rsTemp!����) Then
                        blnFind = True
                        Exit For
                    End If
                Next
                If Not blnFind Then
                    '83266:���ϴ�,2015/3/18,ҽ�ƿ������ж��Ƿ�����
                    If InStr(",1,2,", "," & Val(Nvl(rsTemp!����)) & ",") > 0 _
                        And Val(Nvl(rsTemp!Ӧ����)) <> 1 Then
                        '������ҽ���Ľ��㷽ʽ����֧Ʊ��
                         Set objCard = New Card
                         objCard.���� = Mid(Nvl(!����), 1, 1)
                         objCard.�ӿڱ��� = Nvl(!����)
                         objCard.�ӿڳ����� = ""
                         objCard.�ӿ���� = -1 * lngKey
                         objCard.���㷽ʽ = Nvl(!����)
                         objCard.���� = Nvl(!����)
                         objCard.���� = True
                         objCard.ȱʡ��־ = Val(Nvl(rsTemp!ȱʡ)) = 1
                         objCard.֧������ = True
                         objCard.�������� = Val(!����)
                        mobjPayCards.Add objCard, "K" & lngKey
                        lngKey = lngKey + 1
                    End If
                End If
                .MoveNext
            Loop
        End With
        '��������
        For Each objCard In objCards
            If objCard.���ѿ� = False Then
                rsTemp.Filter = "����='" & objCard.���㷽ʽ & "'"
                If Not rsTemp.EOF Then
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
                End If
            End If
        Next
        
        If ExistԤ����() Then
            'ǿ�Ƽ���Ԥ�����
            Set objCard = New Card
            objCard.���� = "Ԥ"
            objCard.�ӿڱ��� = ""
            objCard.�ӿڳ����� = ""
            objCard.�ӿ���� = -1 * lngKey
            objCard.���㷽ʽ = "Ԥ���"
            objCard.���� = "Ԥ���"
            objCard.���� = True
            objCard.ȱʡ��־ = False
            objCard.֧������ = True
            objCard.�������� = "-99"
            mobjPayCards.Add objCard, "K" & lngKey
        End If
        
        If mobjPayCards.Count = 0 Then
            MsgBox "���㿨��������,ԭ���������:" & vbCrLf & _
                "1)δ�������ý��㿨,�뵽��ҽ�ƿ���𡻺͡��豸���á�������" & vbCrLf & _
                "2)δ���ý��㿨��[���ʼ�����]����,���ڡ�ҽ�ƿ����������", vbInformation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
'    End If
    
    mblnNotClick = True
    mlngPre֧����ʽ = -1
    With cbo֧����ʽ
        .Clear
        For i = 1 To mobjPayCards.Count
            Set objCard = mobjPayCards(i)
            blnFind = False
            If mstr�ų����㷽ʽ <> "" Then
                varData = Split(mstr�ų����㷽ʽ, ",")
                For j = 0 To UBound(varData)
                    If objCard.���㷽ʽ = varData(j) Then
                        blnFind = True: Exit For
                    End If
                Next
            End If
            If blnFind = False Then '�ų��Ĳ�����
                If objCard.�ӿ���� <= 0 _
                    Or objCard.�ӿ���� > 0 And (mstrDefaultBalance = objCard.���㷽ʽ _
                                            Or mcllForceDelToCash.Count = 0 And objCard.�Ƿ�ת�ʼ�����) Then
                    .AddItem objCard.����
                    .ItemData(.NewIndex) = i
                    
                    If objCard.ȱʡ��־ And .ListIndex < 0 Then .ListIndex = .NewIndex: mlngPre֧����ʽ = i
                    If objCard.�������� = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex: mlngPre֧����ʽ = i
                    If mstrDefaultBalance = objCard.���㷽ʽ Then .ListIndex = .NewIndex: mlngPre֧����ʽ = i
                End If
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0: mlngPre֧����ʽ = i
        If .ListCount = 0 Then
            MsgBox "û�п��õĽ��㷽ʽ�����ܼ��������ȵ����㷽ʽ����������һ������Ϊ1��2�Ľ��㷽ʽ��" & vbNewLine & _
                vbNewLine & _
                "ԭ���տ�ʱʹ�õĽ��㷽ʽ�������޿��˻ؽ��", vbExclamation, gstrSysName
            mblnUnLoad = True: Exit Sub
        End If
    End With
    mblnNotClick = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ExistԤ����() As Boolean
    '���ܣ��Ƿ񻹴��ڿ���Ԥ����
    Dim lngTop As Long, lngTop1 As Long
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandler
    mdbl����Ԥ�� = 0
    txtԤ���.Text = ""
    If RoundEx(mdblDelMoney, 6) >= 0 Then Exit Function '�տ�˿���Ϊ��ʱ������ʹ��Ԥ����
    If mrsOldBalance Is Nothing Then Exit Function
    
    '2.�շ�Ԥ�����
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    mrsOldBalance.Filter = "����=1"
    Do While Not mrsOldBalance.EOF
        mdbl����Ԥ�� = mdbl����Ԥ�� + Val(mrsOldBalance!��Ԥ��)
        mrsOldBalance.MoveNext
    Loop
    If RoundEx(mdbl����Ԥ��, 6) = 0 Then Exit Function
    
    '2.����������Ԥ�����
    '����������ID
    strSQL = _
        "Select ����id" & vbNewLine & _
        "From ����Ԥ����¼" & vbNewLine & _
        "Where ������� In (Select a.�������" & vbNewLine & _
        "               From ���ò����¼ A, ���ò����¼ B" & vbNewLine & _
        "               Where a.No = b.No And a.��¼���� = b.��¼���� And a.���ӱ�־ = b.���ӱ�־" & vbNewLine & _
        "                     And b.��¼���� = 1 And b.������� = [1])"
    '���ý���ID
    strSQL = strSQL & vbNewLine & _
        "Minus" & vbNewLine & _
        "Select Distinct a.����id As ԭ����id" & vbNewLine & _
        "From ������ü�¼ A, ������ü�¼ B" & vbNewLine & _
        "Where a.��¼���� = b.��¼���� And a.No = b.No And a.��� = b.��� And b.��¼״̬ <> 2" & vbNewLine & _
        "      And b.����id In (Select �շѽ���id From ���ò����¼ Where ��¼���� = 1 And ������� = [1])"
    strSQL = _
        "Select Nvl(��Ԥ��, 0) As ��Ԥ��" & vbNewLine & _
        "From ����Ԥ����¼" & vbNewLine & _
        "Where Mod(��¼����, 10) = 1 And Nvl(Ԥ�����, 0) = 1 And" & vbNewLine & _
        "      ����id In (" & strSQL & ")"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjDelBalance.�������)
    Do While Not rsTemp.EOF
        mdbl����Ԥ�� = mdbl����Ԥ�� + Val(rsTemp!��Ԥ��)
        rsTemp.MoveNext
    Loop
    If RoundEx(mdbl����Ԥ��, 6) = 0 Then Exit Function
    
    mdbl����Ԥ�� = mdbl����Ԥ��
    'ҽ�����������˷����ܽ��
    If RoundEx(mdbl����Ԥ��, 6) < -1 * mdblDelMoney Then
        lblԤ���.Visible = True
        txtԤ���.Visible = True
        txtԤ���.Text = Format(mdbl����Ԥ��, "0.00")
        
        lngTop = lblԤ���.Top: lngTop1 = txtԤ���.Top
        lblԤ���.Top = lblPayType.Top: txtԤ���.Top = cbo֧����ʽ.Top
        lblPayType.Top = lbl�˿���.Top: cbo֧����ʽ.Top = txt�ɿ�.Top
        lbl�˿���.Top = lbl�������.Top: txt�ɿ�.Top = txt�������.Top
        lbl�������.Top = lblժҪ.Top: txt�������.Top = txtժҪ.Top
        lblժҪ.Top = lngTop: txtժҪ.Top = lngTop1: txtժҪ.Height = 600
        
        mdbl��ǰδ�� = mdblDelMoney - (-1 * mdbl����Ԥ��)
        Exit Function
    End If
    
    ExistԤ���� = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ʱ����Ч��
    '����:������Ч,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 15:01:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objCard As Card
    
    On Error GoTo errHandle
    '83222,Ƚ����,2015-3-17,���÷�ʽ����ֻ��һ��ͨ
'    If Val(txt�ɿ�.Text) = 0 And cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) <> 1 Then
'        MsgBox "��ǰ" & lbl�˿���.Caption & "Ϊ�㣬����ʹ�÷��ֽ���㷽ʽ��", vbOKOnly + vbInformation, gstrSysName
'        If cbo֧����ʽ.Enabled And cbo֧����ʽ.Visible Then cbo֧����ʽ.ListIndex = 0
'        Exit Function
'    End If
    '�������
    If mbytFunc = EM_BalanceReDel Then
        If zlIsCheckExistErrBill(mobjDelBalance.�������, True) = False Then
            MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If zlCheckOtherSessionDoing(mobjDelBalance.�������) Then
            MsgBox "��ǰ�����������������㴰���н��д����㲻�ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If Not CheckTextLength("�������", txt�������) Then Exit Function
    If Not CheckTextLength("ժҪ", txtժҪ) Then Exit Function
    If IsValidԤ����() = False Then Exit Function
    
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
    
    If GetCurCard(objCard) = False Then
        MsgBox "��ǰ" & lblPayType.Caption & "δѡ��,��ѡ��!", vbOKOnly + vbInformation, gstrSysName
        If cbo֧����ʽ.Enabled And cbo֧����ʽ.Visible Then cbo֧����ʽ.SetFocus
        Exit Function
    End If
    If CheckThreeSwapIsValied(objCard, mdbl��ǰδ��) = False Then
        If cbo֧����ʽ.Enabled And cbo֧����ʽ.Visible Then cbo֧����ʽ.SetFocus
        Exit Function
    End If
    '��鵱ǰ�����Ƿ�������ִ�����,��Ҫ�ǲ���ԭ����м��
    '��ֹ��������Ա����:
    gstrSQL = "" & _
    "   Select  1  From ����Ԥ����¼ A " & _
    "   Where   A.�������=[1] and nvl(A.У�Ա�־,0)<>0 and Rownum =1 and A.��¼״̬=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.�������)
    
    If rsTemp.EOF Then
        '�����Ǳ�����ִ��,������Ҫ����Ƿ�����ִ��
        gstrSQL = "Select ��¼״̬, ����Ա����,����״̬ From ���ò����¼ Where ����ID=[1] And rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.����ID)
        If Not rsTemp.EOF Then
            If Val(Nvl(rsTemp!����״̬)) <> 1 Then
                MsgBox "�Ѿ��������˷ѽ���,�����ٽ����˷ѽ���!", vbOKOnly + vbInformation, gstrSysName
                'ִ���շ�
                Unload Me
                Exit Function
            End If
            If Nvl(rsTemp!����Ա����) <> UserInfo.���� Then
                MsgBox "�õ��ݲ��Ǳ����˷ѽ��㵥,���ܴ�����������Ա�ĵ���!", vbOKOnly + vbInformation, gstrSysName
                'ִ���շ�
                Unload Me
                Exit Function
            End If
        End If
    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
   
Private Sub cbo֧����ʽ_Click()
    Dim objCard As Card, intSelectIndex As Integer
    Dim i As Integer
    
    If mblnNotClick Then Exit Sub
    If mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) Then Exit Sub
    
    '105432
    If mlngPre֧����ʽ > 0 And mdbl��ǰδ�� < 0 Then 'ֻ���˿�ż��
        If Not mrsOldBalance Is Nothing And Val(txt�ɿ�.Text) <> 0 Then
            '��������շѽ��㷽ʽ�оͲ��ü�飬��Ҫ���֧�֡�ת�ʼ����ۡ���
            Set objCard = mobjPayCards(mlngPre֧����ʽ)
            mrsOldBalance.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "' And �˷�=0"
            
            mblnNotClick = True
            intSelectIndex = cbo֧����ʽ.ListIndex
            cbo֧����ʽ.ListIndex = cbo.FindIndex(cbo֧����ʽ, mlngPre֧����ʽ)
            
            If Not mrsOldBalance.EOF Then
                If ThreeBalanceCheck(Me, mlngModule, mobjPayCards(mlngPre֧����ʽ), _
                      mcllForceDelToCash, cbo֧����ʽ.Text) = False Then mblnNotClick = False: Exit Sub
            End If
            
            Set objCard = mobjPayCards(cbo֧����ʽ.ItemData(intSelectIndex))
            If objCard.�ӿ���� > 0 And objCard.�Ƿ�ת�ʼ����� _
                And mcllForceDelToCash.Count > 0 And mstrDefaultBalance <> objCard.���㷽ʽ Then
                MsgBox "ǿ������ʱ������ѡ������ת�ʼ����۵Ľ��㷽ʽ��", vbInformation, gstrSysName
                mblnNotClick = False: Exit Sub
            End If
            
            cbo֧����ʽ.ListIndex = intSelectIndex
            mblnNotClick = False
        End If
    End If
    
    mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    
    '�л�������Ҫ���
    Set objCard = mobjPayCards(mlngPre֧����ʽ)
    If objCard.�ӿ���� > 0 And objCard.���ѿ� = False Then
        For i = 1 To mcllForceDelToCash.Count
            If mcllForceDelToCash(i)(1) = objCard.���� Then Exit For
        Next
        If i <= mcllForceDelToCash.Count Then mcllForceDelToCash.Remove i
    End If
    Call SetControlEnabled
    Call Show�����(-1 * mdbl��ǰδ��)
End Sub

Private Sub cbo֧����ʽ_GotFocus()
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
End Sub

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub cmdExit_Click()
    If gfrmMain Is Nothing Then
       Call ExcuteMainReshData
    End If
    Unload Me: Exit Sub
End Sub

Private Sub ExcuteMainReshData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ���������ˢ������
    '����:���˺�
    '����:2014-06-17 15:09:44
    '˵��:��Ҫ��Ӧ��ҽ��ˢ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not gfrmMain Is Nothing Then Exit Sub
    Call mfrmMain.zlExeBalanceWinRefrshData(mblnOK, mobjDelBalance)
End Sub

Private Sub cmdOK_Click()
    '���ݽ��水�˻س���
    If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    
    mblnThreeSwapSingle = False
    If isValied = False Then Exit Sub
    If txt�ɿ�.Text <> "0.00" Then Call ShowLedInfor
    If SaveCharge = False Then Exit Sub
    Unload Me
    Call ExcuteMainReshData
End Sub

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ʾ״̬
    '����:���˺�
    '����:2014-09-18 17:10:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngLeft As Long
    
    cmdOk.Visible = True
    If mbytFunc = EM_BalanceReDel Then
        cmdExit.Visible = True
        lngLeft = cmdOk.Left
        cmdOk.Left = cmdExit.Left
        cmdExit.Left = lngLeft
    Else
        cmdExit.Visible = False
    End If
End Sub
 

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnLoad Then Unload Me: Exit Sub
    
    Call SetControlEnabled
    Call SetCtrlVisible
    If cbo֧����ʽ.Enabled And cbo֧����ʽ.Visible Then cbo֧����ʽ.SetFocus
    Call Show�����(-1 * mdbl��ǰδ��)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
    Case vbKeyAdd, vbKeyF4
        If gTy_Module_Para.blnʹ�üӼ��л� = False And KeyCode = vbKeyAdd Then Exit Sub
        If Me.ActiveControl Is cbo֧����ʽ Then
            i = cbo֧����ʽ.ListIndex
            If i >= cbo֧����ʽ.ListCount - 1 Then
                i = 0
            Else
                i = i + 1
            End If
            cbo֧����ʽ.ListIndex = i
        End If
    Case vbKeySubtract
        If gTy_Module_Para.blnʹ�üӼ��л� = False And KeyCode = vbKeySubtract Then Exit Sub
        If Me.ActiveControl Is cbo֧����ʽ Then
            i = cbo֧����ʽ.ListIndex
            If i <= 0 Then
                i = cbo֧����ʽ.ListCount - 1
            Else
                i = i - 1
            End If
            cbo֧����ʽ.ListIndex = i
        End If
     Case vbKeyF12
            If Shift = vbCtrlMask Then
                'ǿ����LED����,(�ϼ�)
                 Call LedVoiceSpeak
            End If
    Case vbKeyF2
        cmdOK_Click '43169
    Case vbKeyReturn
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    'ѡ������������Ƿ����˻س�����
    mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    mstrTittle = "ҽ���������˷���Ϣ"
    Me.Caption = mstrTittle
    Call InitFace
    zlControl.CboSetWidth cbo֧����ʽ.hWnd, cbo֧����ʽ.Width * 1.3
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    SaveWinState Me, App.ProductName, mstrTittle
    If Not mrsOldBalance Is Nothing Then Set mrsOldBalance = Nothing
End Sub
Private Sub LedVoiceSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2014-09-18 17:20:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gblnLED = False Then Exit Sub
    zl9LedVoice.Speak "#21 " & Format(mdbl��ǰδ��, "0.00")
    mbln�ѱ��� = True
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
   If Panel.Key = "Calc" Then
        mlngR = FindWindow("SciCalc", "������")
        If mlngR <> 0 Then
            BringWindowToTop mlngR
        Else
            On Error Resume Next
            Shell "calc.exe", vbNormalFocus
        End If
  End If
End Sub
 
Private Sub txt�ɿ�_GotFocus()
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
End Sub
Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾLed��Ϣ
    '����:���˺�
    '����:2014-09-18 17:24:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    If gblnLED = False Then Exit Sub
    If Not GetCurCard(objCard) Then Exit Sub
    
    'ֻ�н��ֲ���ʾ
    If objCard.�������� = 1 Then
        zl9LedVoice.DispCharge mdbl��ǰδ��, 0, 0
    Else '����֧���ֽ�ʱ�Ĵ���
        Call zl9LedVoice.DisplayBank( _
            "�ϼ�:" & mdbl��ǰδ�� & "Ԫ,Ӧ��:" & -1 * mdbl��ǰδ�� & "Ԫ")
    End If
    zl9LedVoice.Speak "#22 " & -1 * Val(txt�ɿ�.Text)
    zl9LedVoice.Speak "#3"
End Sub

Private Sub LedDisplayBank()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҽ����Ϣ
    '����:���˺�
    '����:2014-09-18 17:28:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double, i As Long
    Dim strҽ�� As String, str�������� As String, str��һ��ͨ As String, str��ͨ���� As String
    Dim varPara  As Variant, str���㷽ʽ As String
    If Not gblnLED Then Exit Sub
    
    With vsBalance
        For i = 1 To .Rows - 1
            'ҽ������
            If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then
                strҽ�� = strҽ�� & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.TextMatrix(i, .ColIndex("֧�����"))), "0.00")
            End If
        Next
    End With
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
        If str���㷽ʽ > "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str���㷽ʽ
    End Select
    zl9LedVoice.Speak "#21 " & Format(-1 * mdbl��ǰδ��, "0.00")
End Sub
 
Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    Dim objCard As Card
    zlControl.TxtCheckKeyPress txt�ɿ�, KeyAscii, m���ʽ
    If KeyAscii <> 13 Then Exit Sub
    If mblnCacheKeyReturn = True Then mblnCacheKeyReturn = False
    If GetCurCard(objCard) = False Then Exit Sub
    KeyAscii = 0
    If objCard.�������� = 1 Then
        If cmdOk.Enabled And cmdOk.Visible Then cmdOk.SetFocus
        Exit Sub
    End If
    zlCommFun.PressKey vbKeyTab
End Sub
 
Private Sub txt�ɿ�_LostFocus()
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�������_GotFocus()
   zlControl.TxtSelAll txt�������
End Sub
Private Sub txt�������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txt�������, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtժҪ_GotFocus()
    zlControl.TxtSelAll txtժҪ
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdOk.Visible And cmdOk.Enabled Then cmdOk.SetFocus
    End If
End Sub
 
    
Private Function SaveCharge() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 15:58:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTrans  As Boolean, strSQL As String, dblErrMoney As Double '����
    Dim objCard As Card, dblMoney As Double, dblTemp As Double
    Dim str���㷽ʽ  As String, str����ID As String
    Dim cllPro As Collection, rsTemp As ADODB.Recordset
    Dim str����IDs As String, dbl��Ԥ�� As Double
    
    Err = 0: On Error GoTo errHandle
    If GetCurCard(objCard) = False Then
        MsgBox lblPayType.Caption & "��ʽδѡ��!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If txtԤ���.Visible Then
        dbl��Ԥ�� = -1 * Val(txtԤ���.Text)
    End If
    dblMoney = -1 * mdbl��ǰδ��
    
    Call Show�����(dblMoney, dblErrMoney)
    If objCard.�������� = 1 Then
        '���ܴ���10��Ǯ
        If Abs(dblErrMoney) > 1.5 Then
            Call MsgBox("������,�����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    ElseIf objCard.�������� = -99 Then
        dbl��Ԥ�� = -1 * dblMoney
    End If
    strSQL = "Select distinct ����ID From ����Ԥ����¼ where �������=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mobjDelBalance.�������)
    With rsTemp
        Do While Not .EOF
            str����IDs = str����IDs & "," & Val(Nvl(!����ID))
            .MoveNext
        Loop
    End With
    If str����IDs = "" Then str����IDs = "," & mobjDelBalance.����ID
    str����IDs = Mid(str����IDs, 2)
    
    Set cllPro = New Collection
    If mbytFunc = EM_BalanceReDel Then
        strSQL = "Zl_�����շ��쳣_Update("
        strSQL = strSQL & "Null,"
        strSQL = strSQL & "To_Date('" & Format(mobjDelBalance.�˷�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
        strSQL = strSQL & mobjDelBalance.����ID & ")"
        zlAddArray cllPro, strSQL
        If mobjDelBalance.����ID <> 0 Then
            strSQL = "Zl_�����շ��쳣_Update("
            strSQL = strSQL & "Null,"
            strSQL = strSQL & "To_Date('" & Format(mobjDelBalance.�˷�ʱ��, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            strSQL = strSQL & mobjDelBalance.����ID & ")"
            zlAddArray cllPro, strSQL
        End If
    End If
    
    If mblnThreeSwapSingle = False Then
        If objCard.�������� = -99 Then 'Ԥ����
            str���㷽ʽ = ""
        Else
            str���㷽ʽ = objCard.���㷽ʽ
            str���㷽ʽ = str���㷽ʽ & "|" & dblMoney
            str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(txt�������.Text) <> "", txt�������.Text, " ")
            str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(txtժҪ.Text) <> "", txtժҪ.Text, " ")
        End If
        
        'Zl_���ò������_����˷�
        strSQL = "Zl_���ò������_����˷�("
        '  ����id_In     In ���ò����¼.����id%Type,
        strSQL = strSQL & "" & IIf(mobjDelBalance.����ID = 0, mobjDelBalance.����ID, mobjDelBalance.����ID) & ","
        '  ���㷽ʽ_In   Varchar2,��ʽ:���㷽ʽ|������|�������|����ժҪ
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "" & IIf(objCard.�ӿ���� > 0, objCard.�ӿ����, "NULL") & ","
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "'" & IIf(objCard.�ӿ���� > 0, mCurBrushCard.str����, "") & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "'" & IIf(objCard.�ӿ���� > 0, mCurBrushCard.str������ˮ��, "") & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & IIf(objCard.�ӿ���� > 0, mCurBrushCard.str����˵��, GetForceDelToCashNote(mcllForceDelToCash)) & "',"
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null
        strSQL = strSQL & "" & dblErrMoney & ","
        '  ��ɽ���_In       Number := 1,
        strSQL = strSQL & "" & 1 & ","
        '  ���������ν���_In Number := 0,
        strSQL = strSQL & "" & 0 & ","
        '  ��Ԥ��_In         ����Ԥ����¼.��Ԥ��%Type := Null���˿�ʱΪ�����տ�ʱΪ��
        strSQL = strSQL & "" & dbl��Ԥ�� & ")"
        zlAddArray cllPro, strSQL
        
        Err = 0: On Error GoTo ErrRoll:
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        '83222,Ƚ����,2015-3-17,������Ϊ��ʱ�����ýӿ�ֱ��ͨ��
        If objCard.�ӿ���� > 0 And RoundEx(dblMoney, 6) <> 0 Then
            If ExecuteThreeSwapPayInterface(objCard, mobjDelBalance.�������, str����IDs, dblMoney) = False Then Exit Function
        Else
            gcnOracle.CommitTrans
        End If
        blnTrans = False
        mblnOK = True: SaveCharge = True
        Exit Function
    End If

    '��������ÿһ�ʵ��������˷ѽӿ�
    Err = 0: On Error GoTo ErrRoll
    gcnOracle.BeginTrans
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True, True
    blnTrans = False
    
    '83222,Ƚ����,2015-3-17,������Ϊ��ʱ�����ýӿ�ֱ��ͨ��
    If RoundEx(dblMoney, 6) <> 0 Then
       If ExecuteThreeSwapPayInterface(objCard, mobjDelBalance.�������, str����IDs, dblMoney) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    mblnOK = True: SaveCharge = True
    Exit Function
ErrRoll:
    If blnTrans Then gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
End Function

Private Sub SetControlEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�������
    '����:���˺�
    '����:2012-02-03 15:08:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, objCard As Card
    
    blnEdit = GetCurCard(objCard)
    txt�������.Enabled = blnEdit And objCard.�������� <> 1 And objCard.�������� <> -99
    txtժҪ.Enabled = blnEdit And objCard.�������� <> 1 And objCard.�������� <> -99
    txt�������.BackColor = IIf(txt�������.Enabled, &H80000005, Me.BackColor)
    txtժҪ.BackColor = IIf(txtժҪ.Enabled, &H80000005, Me.BackColor)
    cbo֧����ʽ.Enabled = mbytFunc <> EM_Balance_Err_Cancel
    cbo֧����ʽ.BackColor = IIf(cbo֧����ʽ.Enabled, &H80000005, Me.BackColor)
End Sub

Private Function GetCurCard(ByRef objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ǰ��
    '����:objCard-���ص�ǰ�˿��ɿ�Ŀ�����
    '����:�ɹ�,���ؿ�����
    '����:���˺�
    '����:2014-07-09 11:03:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
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

Private Sub Show�����(ByRef dblMoney As Double, Optional ByRef dblErrMoney As Double)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����
    '���:dblMoney-�����˵Ľ��
    '����:dblMoney-����ʵ���˵Ľ��
    '     dblErrMoney-����������
    '����:���˺�
    '����:2014-07-09 18:44:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, dblTemp As Double
    
    On Error GoTo errHandle
    
    If GetCurCard(objCard) = False Then Exit Sub
    dblErrMoney = 0

    If objCard.�������� = 1 Then
        '�ֽ�
        dblTemp = dblMoney
        If mobjDelBalance.intInsure > 0 Then
            If mblnҽ���ֱ� Then
                dblMoney = CentMoney(CCur(dblTemp))
            Else
                dblMoney = Format(dblTemp, "0.00")
            End If
        Else
             dblMoney = CentMoney(CCur(dblTemp))
        End If
       dblErrMoney = RoundEx(dblMoney - dblTemp, 6)
    End If
    
    lbl���.Visible = dblErrMoney <> 0
    lbl���.Caption = "����:" & zlFormatNum(dblErrMoney)
    lbl���.Left = cmdOk.Left - lbl���.Width - 100
    txt�ɿ�.Text = Format(Abs(dblMoney), "0.00")
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function ExecuteThreeSwapPayInterface(objCard As Card, ByVal str������� As String, ByVal str����IDs As String, _
    ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�����ӿ�)
    '���:str�������-��������Ž��д���
    '     str����Ids-���θ��µĽ���IDs
    '     dblMoney-����֧�����
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String, strXMLExpend As String
    Dim i As Long, strSQL As String, strTemp As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim rsBalance As ADODB.Recordset
    Dim strInXML As String, strOutXML As String
    Dim objXml As clsXML
    Dim dbl��Ԥ�� As Double, cllThreeSwapDel As Collection
    Dim rsTemp As ADODB.Recordset, dblTemp As Double
    Dim lngRow As Long, strValue As String
    Dim lngԭ����ID As Long, str���� As String
    Dim str���㷽ʽ As String, lngԭ����ID As Long
    
    On Error GoTo errHandle
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    If objCard.�Ƿ�ת�ʼ����� Then
        'zlTransferAccountsMoney
        '������  ��������    ��/��   ��ע
        'frmMain Object  In  ���õ�������
        'lngModule   Long    In  HIS����ģ���
        'lngCardTypeID   Long    In  �����ID
        'strCardNo   String  In  ����
        'strBalanceID    String  In  ����ID
        'dblMoney    Double  In  ת�ʽ��
        'strSwapGlideNO  String  Out ������ˮ��
        'strSwapMemo String  Out ����˵��
        'strSwapExtendInfor  String  In �˷�ҵ��ʱ�����뱾���˷ѵĳ���ID:
        '                               ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                               �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
        '                           Out ������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
        'strXMLExpend String In   XML��:
        '                            <IN>
        '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�˷�ҵ��
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
        strXMLExpend = "<IN><CZLX>1</CZLX></IN>"
        '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
        strSwapExtendInfor = "3|" & str����IDs: strTemp = strSwapExtendInfor
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.�ӿ����, mCurBrushCard.str����, _
            str�������, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor, strXMLExpend) = False Then
            gcnOracle.RollbackTrans: Call ShowErrMsg(1, strXMLExpend)
            Exit Function
        End If
        mCurBrushCard.str������ˮ�� = strSwapGlideNO
        mCurBrushCard.str����˵�� = strSwapMemo
        
        Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapGlideNO, strSwapMemo, cllUpdate, 0)
        
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
        End If
        zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
        gcnOracle.CommitTrans
    Else
        If mblnThreeSwapSingle Then
            Set rsBalance = zlGetCanDelBalanceRecords(str�������, objCard.�ӿ����)
            dblTemp = dblMoney
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                If Val(Nvl(rsBalance!���)) > RoundEx(dblTemp, 6) Then
                    dbl��Ԥ�� = RoundEx(dblTemp, 6)
                    dblTemp = 0
                Else
                    dbl��Ԥ�� = Val(Nvl(rsBalance!���))
                    dblTemp = dblTemp - Val(Nvl(rsBalance!���))
                End If
                
                lngԭ����ID = Val(Nvl(rsBalance!ԭ����ID))
                lngԭ����ID = Val(Nvl(rsBalance!����ID))
                str���� = Nvl(rsBalance!����)
                strSwapGlideNO = Nvl(rsBalance!������ˮ��)
                strSwapMemo = Nvl(rsBalance!����˵��)
                strSwapExtendInfor = "3|" & str����IDs: strTemp = strSwapExtendInfor
                'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
                    ByVal lngCardTypeID As Long, ByVal strCardNo As String, _
                    ByVal strBalanceIDs As String, ByVal dblMoney As Double, _
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
                '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
                '       dblMoney-�˿���
                '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
                '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
                '       strSwapExtendInfor-���룬�����˷ѵĳ���ID��
                '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
                '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
                '       strSwapExtendInfor-���������׵���չ��Ϣ
                '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
                If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.�ӿ����, _
                    objCard.���ѿ�, str����, "3|" & lngԭ����ID, dbl��Ԥ��, _
                    strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
                    gcnOracle.RollbackTrans
                    
                    Call LoadData(str�������)
                    mdbl��ǰδ�� = mdblDelMoney
                    Exit Function
                End If
                
                'Zl_���ò������_����˷�
                strSQL = "Zl_���ò������_����˷�("
                '  ����id_In     In ���ò����¼.����id%Type,
                strSQL = strSQL & "" & IIf(mobjDelBalance.����ID = 0, mobjDelBalance.����ID, mobjDelBalance.����ID) & ","
                str���㷽ʽ = objCard.���㷽ʽ
                str���㷽ʽ = str���㷽ʽ & "|" & -1 * dbl��Ԥ��
                str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(txt�������.Text) <> "", txt�������.Text, " ")
                str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(txtժҪ.Text) <> "", txtժҪ.Text, " ")
                '  ���㷽ʽ_In   Varchar2,��ʽ:���㷽ʽ|������|�������|����ժҪ
                strSQL = strSQL & "'" & str���㷽ʽ & "',"
                '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
                strSQL = strSQL & "" & IIf(objCard.�ӿ���� > 0, objCard.�ӿ����, "NULL") & ","
                '  ����_In       ����Ԥ����¼.����%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  �����_In   ������ü�¼.ʵ�ս��%Type := Null
                strSQL = strSQL & "" & 0 & ","
                '  ��ɽ���_In Number:=0:1-��ɲ������;0-δ��ɲ������
                strSQL = strSQL & "" & IIf(RoundEx(dblTemp, 6) > 0, 0, 1) & ","
                '  ���������ν���_In Number:=0
                strSQL = strSQL & "" & 1 & ")"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                strSQL = "Zl_�����˿���Ϣ_Insert("
                strSQL = strSQL & "" & Val(str�������) & ","
                strSQL = strSQL & "" & lngԭ����ID & ","
                strSQL = strSQL & "" & dbl��Ԥ�� & ","
                strSQL = strSQL & "'" & str���� & "',"
                strSQL = strSQL & "'" & strSwapGlideNO & "',"
                strSQL = strSQL & "'" & strSwapMemo & "')"
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
                
                gcnOracle.CommitTrans
                
                Set cllThreeSwap = New Collection
                If strTemp <> strSwapExtendInfor Then
                    Call zlAddThreeSwapSQLToCollection(False, Abs(Val(str�������)), _
                        objCard.�ӿ����, objCard.���ѿ�, str����, strSwapExtendInfor, cllThreeSwap, lngԭ����ID)
                End If
                zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
                gcnOracle.BeginTrans
                
                rsBalance.MoveNext
            Loop
            gcnOracle.CommitTrans
            
            If RoundEx(dblTemp, 6) > 0 Then
                MsgBox "�˿���(" & Format(dblMoney, "0.00") & ")���ڿ��˽��(" & Format(dblMoney - dblTemp, "0.00") & ")��", vbOKOnly + vbInformation, gstrSysName
                
                Call LoadData(str�������)
                mdbl��ǰδ�� = mdblDelMoney
                Exit Function
            End If
            
            ExecuteThreeSwapPayInterface = True
            Exit Function
        Else
            'Public Function zlReturnMultiMoney(frmMain As Object, ByVal lngModule As Long, _
                ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByVal strInXML As String, _
                ByVal lng����ID As Long, ByRef strOutXml As String, ByRef strExpend As String) As Boolean
            '---------------------------------------------------------------------------------
            '����:�ʻ��ۿ���˽���(��ʻ���)
            '���:frmMain-���õ�������
            '       lngModule-���õ�ģ���
            '       lngCardTypeID-�����ID:ҽ�ƿ����.ID
            '       strInXML-XML��:
            '       <JSLIST>
            '           <JS>
            '               <KH>����</KH>
            '               <JYLSH>������ˮ��</JYLSH>
            '               <JYSM>����˵��</JYSM>
            '               <ZFJE>���Ͻ��</ZFJE>
            '               <JSLX>����</JSLX>  //1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�,6-�������
            '               <ID></ID>    //����=1ʱ,Ԥ��ID;����=2,6ʱ��Ϊԭ����ID
            '           </JS>
            '       </JSLIST>
            '       lng����ID-����ʱ�ĳ���ID(����ʱ���˷�ʱ��Ч������Ϊ0��;����=6������������
            '       strExpend-�ޣ����������Ժ���չ)
            '����:
            '     strOutXML-����XML��
            '       <JSLIST>
            '           <JS>
            '               <KH>����</KH>
            '               <TKLSH>�˿����ˮ��</TKLSH>
            '               <TKSM>�˿��˵��</TKSM>
            '               <ID></ID>
            '           </JS>
            '       </JSLIST>
            '      strExpend-���׵���չ��Ϣ
            '       <EXPENDS>
            '           <EXPEND>
            '               <XMMC>��Ŀ����1</XMMC>
            '               <XMNR>��Ŀ����2</XMNR>
            '           </EXPEND>
            '       </EXPENDS>
            '����:��������    True:���óɹ�,False:����ʧ��
            '����:2015-11-10
            '˵��:
            '   Ŀǰֻ�н��ʳ���ʱ��Ч�������˿�),����һ���Դ���ͬһ�����Ķ��������������
            '--------------------------------------------------------------------------------
            Set cllThreeSwap = New Collection: Set cllThreeSwapDel = New Collection
            Set objXml = New clsXML
            objXml.ClearXmlText
            
            Set rsBalance = zlGetCanDelBalanceRecords(str�������, objCard.�ӿ����)
            dblTemp = dblMoney
            
            objXml.AppendNode "JSLIST"
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                If Val(Nvl(rsBalance!���)) > RoundEx(dblTemp, 6) Then
                    dbl��Ԥ�� = RoundEx(dblTemp, 6)
                    dblTemp = 0
                Else
                    dbl��Ԥ�� = Val(Nvl(rsBalance!���))
                    dblTemp = dblTemp - Val(Nvl(rsBalance!���))
                End If
                
                objXml.AppendNode "JS"
                    objXml.appendData "KH", Nvl(rsBalance!����), xsString
                    objXml.appendData "JYLSH", Nvl(rsBalance!������ˮ��), xsString
                    objXml.appendData "JYSM", Nvl(rsBalance!����˵��), xsString
                    objXml.appendData "ZFJE", dbl��Ԥ��, xsNumber
                    objXml.appendData "JSLX", 6, xsNumber
                    objXml.appendData "ID", Val(Nvl(rsBalance!����ID)), xsNumber
                objXml.AppendNode "JS", True
                
                strSQL = "Zl_�����˿���Ϣ_Insert("
                strSQL = strSQL & "" & Val(str�������) & ","
                strSQL = strSQL & "" & Val(Nvl(rsBalance!����ID)) & ","
                strSQL = strSQL & "" & dbl��Ԥ�� & ","
                strSQL = strSQL & "'" & Nvl(rsBalance!����) & "',"
                strSQL = strSQL & "'" & Nvl(rsBalance!������ˮ��) & "',"
                strSQL = strSQL & "'" & Nvl(rsBalance!����˵��) & "')"
                zlAddArray cllThreeSwapDel, strSQL
                rsBalance.MoveNext
            Loop
            objXml.AppendNode "JSLIST", True
            
            strInXML = objXml.XmlText
            strOutXML = "": strXMLExpend = ""
            If gobjSquare.objSquareCard.zlReturnMultiMoney(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, strInXML, _
                 Val(str�������), strOutXML, strXMLExpend) = False Then
                gcnOracle.RollbackTrans
                Call ShowErrMsg(1, strXMLExpend)
                Exit Function
            End If
                 
            If strOutXML <> "" Then
                If zlXML_Init = False Then Exit Function
                If zlXML_LoadXMLToDOMDocument(strOutXML, False) = False Then Exit Function
                Call zlXML_GetChildRows("JSLIST", "JS", lngRow)
                For i = 0 To lngRow - 1
                    strSQL = "Zl_�����˿���Ϣ_Insert("
                    strSQL = strSQL & "" & Val(str�������) & ","
                    Call zlXML_GetNodeValue("ID", i, strValue)
                    strSQL = strSQL & "" & Val(strValue) & ","
                    strSQL = strSQL & "" & 0 & ","
                    Call zlXML_GetNodeValue("KH", i, strValue)
                    strSQL = strSQL & "'" & strValue & "',"
                    Call zlXML_GetNodeValue("TKLSH", i, strValue)
                    strSQL = strSQL & "'" & strValue & "',"
                    Call zlXML_GetNodeValue("TKSM", i, strValue)
                    strSQL = strSQL & "'" & strValue & "',"
                    strSQL = strSQL & "" & 1 & ")"
                    zlAddArray cllThreeSwapDel, strSQL
                Next
            End If
            
            If strXMLExpend <> "" Then
                strSwapExtendInfor = ""
                If zlXML_LoadXMLToDOMDocument(strXMLExpend, False) = False Then Exit Function
                Call zlXML_GetChildRows("EXPENDS", "EXPEND", lngRow)
                For i = 0 To lngRow - 1
                    Call zlXML_GetNodeValue("XMMC", i, strValue)
                    strSwapExtendInfor = strSwapExtendInfor & "||" & strValue
                    Call zlXML_GetNodeValue("XMNR", i, strValue)
                    strSwapExtendInfor = strSwapExtendInfor & "|" & strValue
                Next i
            End If
            If strSwapExtendInfor <> "" Then strSwapExtendInfor = Mid(strSwapExtendInfor, 3)
            
            Call zlAddUpdateSwapSQL(False, Abs(Val(str�������)), objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, "", "", cllUpdate, 0)
            Call zlAddThreeSwapSQLToCollection(False, Abs(Val(str�������)), objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
            zlExecuteProcedureArrAy cllThreeSwapDel, Me.Caption, True, True
            zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
            gcnOracle.CommitTrans
        End If
    End If
    
    Err = 0: On Error GoTo ErrOtherHand:
    '��������������Ϣ
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    ExecuteThreeSwapPayInterface = True
    Exit Function
ErrOtherHand:
    ExecuteThreeSwapPayInterface = True
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    If mblnThreeSwapSingle Then
        Call LoadData(str�������)
        mdbl��ǰδ�� = mdblDelMoney
    End If
End Function

Private Sub ShowErrMsg(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ת�˼�������ҵ�������ʾ
    '����:Ƚ����
    'ʱ��:2014-12-2
    '����:
    '   bytType:0-ת�˼��,1-ת�˽���
    '   strXMLErrMsg:��ʽ����
    '            <OUT>
    '               <ERRMSG>������Ϣ</ERRMSG >
    '            </OUT>
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
    MsgBox strValue, vbExclamation + vbOKOnly, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ˢ����֤
    '���:objCard-��ǰ��
    '����:ˢ���ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-09-18 15:03:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExpend As String, strExpand As String
    Dim cllSquareBalance As New Collection
    Dim dblTemp As Double, dbl�ʻ���� As Double
    Dim blnTransfer As Boolean, strBalanceIDs As String
    Dim rsBalance As ADODB.Recordset
    Dim strCardNo As String, strPassWord As String
    Dim dbl��Ԥ�� As Double
    
    On Error GoTo errHandle
    
    If objCard.�ӿ���� <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If dblMoney = 0 Then CheckThreeSwapIsValied = True: Exit Function
    
    'ҽ�������������ܷ��ý���ֻ�ܽ���ת�ʼ�����
    blnTransfer = zlCheckOnlyUseTrans(mobjDelBalance.�������)
    If blnTransfer And objCard.�Ƿ�ת�ʼ����� = False Then
        MsgBox "ҽ���������������ܷ��ý� " & objCard.���� & " ��֧��ת�ʼ����ۣ���ѡ������֧����ʽ��", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If dblMoney > 0 And objCard.�Ƿ�ת�ʼ����� = False Then
        MsgBox "��ǰΪ�տ " & objCard.���� & " ��֧��ת�ʼ����ۣ���ѡ������֧����ʽ��", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
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
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.�ӿ����, False, _
            mobjDelBalance.����, mobjDelBalance.�Ա�, mobjDelBalance.����, dblMoney, strCardNo, strPassWord, _
            False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
        mCurBrushCard.str���� = strCardNo
        mCurBrushCard.str���� = strPassWord
    
        '����ת�ʽӿ�
        'zlTransferAccountsCheck ת�ʼ��ӿ�
        '������  ��������    ��/��   ��ע
        'frmMain Object  In  ���õ�������
        'lngModule   Long    In  HIS����ģ���
        'lngCardTypeID   Long    In  �����ID
        'strCardNo   String  In  ����
        'dblMoney    Double  In  ת�ʽ��(����ʱΪ����)
        'strBalanceIDs   String  In  ����IDs������ö��ŷ��룬��ʾ���ζ��Ĵ��շ���Ŀ��������ҽ��������
        'strXMLExpend String In   XML��:
        '                            <IN>
        '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�˷�ҵ��
        '                            </IN>
        '                    Out  XML��:
        '                            <OUT>
        '                               <ERRMSG>������Ϣ</ERRMSG >
        '                            </OUT>
        '    Boolean ��������    �������ݺϷ�,����True:���򷵻�False
        '˵��:
        '��. ��ҽ���������ʱ���е�����ת��ʱ��һЩ�Ϸ��Լ�飬������ת��ʱ�����Ի���֮��ĵȴ������������������ķ�����
        '��. �����ڼ�����Ҫ����ΪTrue�����������ת�ʹ��ܵĵ��á�
        '����XML��
        strXMLExpend = "<IN><CZLX>1</CZLX></IN>"
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModule, objCard.�ӿ����, _
            mCurBrushCard.str����, -1 * dblMoney, mobjDelBalance.�������, strXMLExpend) = False Then
            Call ShowErrMsg(0, strXMLExpend)
            Exit Function
        End If
    Else
        'ZlGetParaConfig(ByVal frmMain As Object, _
            ByVal lng�����ID As Long, ByVal bln���ѿ� As Boolean, ByVal intPara As Integer, _
            Optional strErrMsg As String, Optional strExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '����:��ȡ�ӿڲ���
            '���: frmMain-���õ�������
            '       intPara: ��������ֵ
            '                1-ˢ����֧����ͬһҳ��:true-��ģʽ��False-��ģʽ
            '                2-����ʱ�Ƿ񵥶������˷ѽӿ�
            '       strExpend-��չ�������������ִ�Ϊ��
            '����:strErrMsg-���صĴ�����Ϣ
            '       strExpend-��չ�������������ִ�Ϊ��
            '����:��������True:���óɹ�,False:����ʧ��
        mblnThreeSwapSingle = gobjSquare.objSquareCard.ZlGetParaConfig(Me, objCard.�ӿ����, objCard.���ѿ�, 2)
        If mblnThreeSwapSingle Then
            Set rsBalance = zlGetCanDelBalanceRecords(mobjDelBalance.�������, objCard.�ӿ����)
            dblTemp = -1 * dblMoney
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                If Val(Nvl(rsBalance!���)) > RoundEx(dblTemp, 6) Then
                    dbl��Ԥ�� = RoundEx(dblTemp, 6)
                    dblTemp = 0
                Else
                    dbl��Ԥ�� = Val(Nvl(rsBalance!���))
                    dblTemp = dblTemp - Val(Nvl(rsBalance!���))
                End If
                
                strBalanceIDs = "6|" & Nvl(rsBalance!����ID)
                mCurBrushCard.str���� = Nvl(rsBalance!����)
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
                    '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
                    '       dblMoney-�˿���
                    '       strSwapNo-������ˮ��(�˿�ʱ���)
                    '       strSwapMemo-����˵��(�˿�ʱ����)
                    '       strXMLExpend    XML IN  ��ѡ����:�쳣���������˷�(1)
                    '����:�˿�Ϸ�,����true,���򷵻�Flase
                    '˵��:
                    '    �ڵ��ÿۿ�ǰ�����ڴ���Oracle�������⣬��ˣ��ٵ��û��˽���ǰ���Ƚ������ݵĺϷ��Լ��,
                    '    �Ա�������������
                If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, objCard.�ӿ����, _
                    objCard.���ѿ�, mCurBrushCard.str����, strBalanceIDs, dbl��Ԥ��, _
                    Nvl(rsBalance!������ˮ��), Nvl(rsBalance!����˵��), _
                    strXMLExpend) = False Then Exit Function
                
                If objCard.�Ƿ��˿��鿨 Then
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
                        '       strXmlIn-XML���,Ŀǰ��ʽ����:
                        '       <IN>
                        '           <CZLX>0</CZLX>    //��������,0-��������ˢ��,1-ת�˵���ˢ��,2-�˿����ˢ��
                        '       </IN>
                        '����:str�������-�������(���ѿ�����)
                        '        lng���ѿ�ID-���ѿ�Ŀ¼.ID(���ѿ�����)
                        '       strCardNO-����ˢ���Ŀ���
                        '       strPassWord-����ˢ������Ӧ������
                        '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                        '����:�ɹ�,����true,���򷵻�False
                    strCardNo = mCurBrushCard.str����
                    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                        objCard.�ӿ����, False, mobjDelBalance.����, mobjDelBalance.�Ա�, mobjDelBalance.����, _
                        dbl��Ԥ��, strCardNo, strPassWord, _
                        False, True, False, False, cllSquareBalance, False, False, _
                        "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
                    mCurBrushCard.str���� = strCardNo
                    mCurBrushCard.str���� = strPassWord
                End If
                
                rsBalance.MoveNext
            Loop
            
            If RoundEx(dblTemp, 6) > 0 Then
                MsgBox "�˿���(" & Format(-1 * dblMoney, "0.00") & ")���ڿ��˽��(" & Format(-1 * dblMoney - dblTemp, "0.00") & ")��", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        Else
            Set rsBalance = zlGetCanDelBalanceRecords(mobjDelBalance.�������, objCard.�ӿ����)
            dblTemp = -1 * dblMoney: strBalanceIDs = ""
            Do While Not rsBalance.EOF And RoundEx(dblTemp, 6) > 0
                strBalanceIDs = strBalanceIDs & "," & Nvl(rsBalance!����ID)
                dblTemp = dblTemp - Val(Nvl(rsBalance!���))
                rsBalance.MoveNext
            Loop
            If strBalanceIDs <> "" Then strBalanceIDs = Mid(strBalanceIDs, 2)
            strBalanceIDs = "6|" & strBalanceIDs
            If RoundEx(dblTemp, 6) > 0 Then
                MsgBox "�˿���(" & Format(-1 * dblMoney, "0.00") & ")���ڿ��˽��(" & Format(-1 * dblMoney - dblTemp, "0.00") & ")��", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
                
            strXMLExpend = mfrmMain.GetDelXMLExpend()
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
            '                                   �շ�����: 1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
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
            If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, _
                strBalanceIDs, -1 * dblMoney, "", "", strXMLExpend) = False Then Exit Function
                
            If objCard.�Ƿ��˿��鿨 Then
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
                If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.�ӿ����, False, _
                    mobjDelBalance.����, mobjDelBalance.�Ա�, mobjDelBalance.����, -1 * dblMoney, strCardNo, strPassWord, _
                    False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
                mCurBrushCard.str���� = strCardNo
                mCurBrushCard.str���� = strPassWord
            End If
        End If
    End If
    
    'zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
    'ByVal strCardTypeID As Long, _
    'ByVal strCardNo As String, strExpand As String, dblMoney As Double
    '���:frmMain-���õ�������
    '        lngModule-ģ���
    '        strCardNo-����
    '        strExpand-Ԥ����Ϊ��,�Ժ���չ
    '����:dblMoney-�����ʻ����
    Call gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, objCard.�ӿ����, _
          mCurBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�)
    If dbl�ʻ���� <> 0 Then
        stbThis.Panels(C4�����ʻ���Ϣ).Text = objCard.���㷽ʽ & "�ʻ����:" & Format(dbl�ʻ����, "0.00")
        stbThis.Panels(C4�����ʻ���Ϣ).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
        stbThis.Panels(C4�����ʻ���Ϣ).Visible = True
    End If
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtԤ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtԤ���_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtԤ���, KeyAscii, m���ʽ
End Sub

Private Sub txtԤ���_LostFocus()
    If IsValidԤ����() = False Then Exit Sub
    
    mdbl��ǰδ�� = mdblDelMoney - (-1 * Val(txtԤ���.Text))
    Call Show�����(-1 * mdbl��ǰδ��)
End Sub

Private Function IsValidԤ����() As Boolean
    '��Ԥ������
    On Error GoTo errHandle
    If txtԤ���.Visible = False Then IsValidԤ���� = True: Exit Function
    
    If txtԤ���.Text = "" Then
        txtԤ���.Text = "0.00"
    ElseIf Not IsNumeric(txtԤ���.Text) And txtԤ���.Text <> "" Then
        ShowMsgbox "��Ч��ֵ��"
        txtԤ���.Text = Format(mdbl����Ԥ��, "0.00")
        zlControl.ControlSetFocus txtԤ���
        Exit Function
    ElseIf Val(txtԤ���.Text) < 0 Then
        ShowMsgbox "Ԥ����˿����Ϊ����"
        zlControl.ControlSetFocus txtԤ���
        txtԤ���.Text = Format(mdbl����Ԥ��, "0.00")
        Exit Function
    ElseIf Val(txtԤ���.Text) > mdbl����Ԥ�� Then
        ShowMsgbox "Ԥ����˿���ܳ������˽��:" & Format(mdbl����Ԥ��, "0.00") & " ��"
        txtԤ���.Text = Format(mdbl����Ԥ��, "0.00")
        zlControl.ControlSetFocus txtԤ���
        Exit Function
    Else
        txtԤ���.Text = Format(Val(txtԤ���.Text), "0.00")
    End If
    IsValidԤ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txtԤ���_GotFocus()
    zlControl.TxtSelAll txtԤ���
End Sub
