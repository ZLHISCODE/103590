VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicDelBalance 
   Caption         =   "�����˷ѽ���"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   15.75
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClinicDelBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   10365
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdExit 
      Caption         =   "����(&X)"
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
      Left            =   8364
      TabIndex        =   27
      Top             =   900
      Width           =   1704
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   36
      ScaleHeight     =   3090
      ScaleWidth      =   7995
      TabIndex        =   16
      Top             =   0
      Width           =   7995
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1260
         Left            =   45
         ScaleHeight     =   1230
         ScaleWidth      =   3060
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1572
         Width           =   3090
         Begin VB.Label lbl�˷Ѻϼ� 
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
            Height          =   612
            Left            =   2028
            TabIndex        =   22
            Top             =   552
            Width           =   1008
         End
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption2 
            Height          =   420
            Left            =   15
            TabIndex        =   21
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   741
            _StockProps     =   6
            Caption         =   "�˷Ѻϼ�"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   15.76
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   2736
         Left            =   3204
         ScaleHeight     =   2700
         ScaleWidth      =   4710
         TabIndex        =   18
         Top             =   90
         Width           =   4740
         Begin VB.ComboBox cbo֧����ʽ 
            BackColor       =   &H8000000F&
            ForeColor       =   &H8000000D&
            Height          =   408
            Left            =   1380
            Style           =   2  'Dropdown List
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   204
            Width           =   1245
         End
         Begin VB.TextBox txt������� 
            Height          =   480
            IMEMode         =   3  'DISABLE
            Left            =   1368
            MaxLength       =   30
            TabIndex        =   6
            Top             =   1332
            Width           =   3225
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
            Height          =   468
            IMEMode         =   3  'DISABLE
            Left            =   2700
            MaxLength       =   12
            TabIndex        =   2
            Top             =   183
            Width           =   1920
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
            Height          =   465
            Left            =   1368
            MaxLength       =   50
            TabIndex        =   8
            Top             =   1992
            Width           =   3210
         End
         Begin VB.TextBox txt�Ҳ� 
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
            Left            =   1368
            Locked          =   -1  'True
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   744
            Width           =   3225
         End
         Begin VB.Label lbl������� 
            AutoSize        =   -1  'True
            Caption         =   "�������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   15.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   108
            TabIndex        =   5
            Top             =   1416
            Width           =   1260
         End
         Begin VB.Label lbl�Ҳ� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ҡ���"
            Height          =   312
            Left            =   372
            TabIndex        =   3
            Top             =   828
            Width           =   996
         End
         Begin VB.Label lblժҪ 
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
            Height          =   312
            Left            =   396
            TabIndex        =   7
            Top             =   2064
            Width           =   960
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��"
            Height          =   312
            Left            =   384
            TabIndex        =   0
            Top             =   240
            Width           =   984
         End
      End
      Begin VB.PictureBox picTotal 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1308
         Left            =   48
         ScaleHeight     =   1275
         ScaleWidth      =   3060
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   90
         Width           =   3090
         Begin XtremeSuiteControls.ShortcutCaption stcCurDelTitle 
            Height          =   450
            Left            =   15
            TabIndex        =   19
            Top             =   30
            Width           =   3045
            _Version        =   589884
            _ExtentX        =   5371
            _ExtentY        =   794
            _StockProps     =   6
            Caption         =   "��ǰӦ��"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   15.76
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label lblδ�˽�� 
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
            Height          =   612
            Left            =   2016
            TabIndex        =   9
            Top             =   588
            Width           =   1008
         End
      End
   End
   Begin VB.Frame fraSplitLeft 
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   8100
      TabIndex        =   13
      Top             =   -84
      Width           =   30
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   14
      Top             =   6120
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3572
            MinWidth        =   882
            Picture         =   "frmClinicDelBalance.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8599
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   2
            Object.Tag             =   "�����շ�Ԥ�������ʾ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   582
            MinWidth        =   1
            Object.Tag             =   "�����շ�������������ʾ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Picture         =   "frmClinicDelBalance.frx":115E
            Key             =   "Calc"
            Object.ToolTipText     =   "������:ALT+?"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1693
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
   Begin VB.PictureBox picBlance 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   2868
      Left            =   84
      ScaleHeight     =   2835
      ScaleWidth      =   10140
      TabIndex        =   15
      Top             =   3096
      Width           =   10176
      Begin VB.CommandButton cmdDel 
         Caption         =   "ɾ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8730
         TabIndex        =   23
         Top             =   75
         Width           =   1080
      End
      Begin VSFlex8Ctl.VSFlexGrid vsBlance 
         Height          =   2295
         Left            =   15
         TabIndex        =   12
         Top             =   495
         Width           =   9930
         _cx             =   17515
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
         Rows            =   7
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmClinicDelBalance.frx":1838
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
      Begin VB.Label lbl���˺ϼ� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "�Ѹ��ϼ�:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4305
         TabIndex        =   11
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label lblDeledInfor 
         Caption         =   "�����������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   10
         Top             =   98
         Width           =   2145
      End
   End
   Begin VB.PictureBox pic��� 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   8232
      ScaleHeight     =   1140
      ScaleWidth      =   2040
      TabIndex        =   24
      Top             =   1656
      Width           =   2040
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Caption         =   "0.0111"
         Height          =   285
         Left            =   135
         TabIndex        =   26
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lbl��� 
         Caption         =   "�������"
         Height          =   315
         Left            =   105
         TabIndex        =   25
         Top             =   90
         Width           =   1455
      End
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
      Height          =   540
      Left            =   8376
      TabIndex        =   28
      Top             =   210
      Width           =   1716
   End
End
Attribute VB_Name = "frmClinicDelBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------
'���������ر���
Public Enum gChargeDelType
    EM_FUN_�˷� = 0
    EM_FUN_���� = 1
End Enum
Private mobjDelBalance As clsCliniDelBalance
Private mbytFunc As gChargeDelType  '0-�շ�;1-����
Private mfrmMain As frmClinicDelAndView
Private mcllDelPro As Collection
Private mlngModule As Long, mstrPrivs As String
Private mcllForceDelToCash As Collection 'ǿ��������Ϣ��Array(����Ա,���������)
'------------------------------------------------------------------------------------------
Private mrsBalance As ADODB.Recordset '��ǰ��������
Private mstr��֧Ʊ As String
Private mblnSingleBalance As Boolean  '��ҽ�����㷽ʽ���⣬�Ƿ�ֻʹ����һ�ֽ��㷽ʽ
    
Private mobjPayCards As Cards
Private mblnNotClick  As Boolean '����������¼�
Private mblnOK As Boolean
Private mblnUnloaded  As Boolean
Private mblnLoad As Boolean
Private mlngR  As Long
'------------------------------------------------------------------------------------------
'�ֲ�����
Private mblnFirst As Boolean
Private mblnUnLoad As Boolean '�Ƿ�Unload����
Private mbln�ѱ��� As Boolean
Private mcur������� As Currency
Private mlngPre֧����ʽ As Long
'----------------------------------------------------------------------------------------------
'ҽ�����
'��ǰ���������ҽ��֧�ֲ���
Private Type TYPE_MedicarePAR
    �����ѽɿ���� As Boolean    '27536
    ���������շ� As Boolean
    �ֱҴ��� As Boolean
End Type
Private mInsurePara As TYPE_MedicarePAR

Private Type TY_BrushCard    'ˢ������
    str���� As String
    str���� As String
    str������ˮ�� As String    '������ˮ��
    str����˵��  As String     '������Ϣ
    str��չ��Ϣ As String    '���׵���չ��Ϣ
    dbl�ʻ���� As Double
End Type
Private mCurBrushCard As TY_BrushCard   '��ǰ��ˢ����Ϣ
Private Type TY_ChargeMoney
    dbl�˷Ѻϼ� As Double
    dbl������� As Double
    dbl����Ӧ�� As Double
    dbl����ҽ���˷� As Double
    dbl���˺ϼ� As Double
    dbl������Ԥ��  As Double
    dbl��ǰδ�� As Double
    dblԤ����� As Double
    dbl������� As Double
    dbl����Ԥ�� As Double
    dblӦ���ۼ� As Double
    dbl�������� As Double
End Type
Private mCurCarge As TY_ChargeMoney
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsOneCard As ADODB.Recordset
Private mrsUsedCards As ADODB.Recordset

Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private mblnCacheKeyReturn As Boolean   '41025:�Ƿ񻺴��˻س���,���ܴ������շѽ���ˢ���б�������˻س�,�����Ҫ�ж�
Private mrsClassMoney As ADODB.Recordset
Private mcllSquareBalance As Collection '���ѿ��˷ѽ�����Ϣ
Private mcllSquareChargeBalance As Collection '���ѿ��շѽ�����Ϣ
Private mcllCurSquareBalance As Collection '��ǰ���ѿ�ˢ����Ϣ
Private mblnNotChange As Boolean
Private mstrTittle As String
Private mblnTurnFee As Boolean

Public Function zlDelCharge(ByVal frmMain As Object, _
    ByVal bytFunc As gChargeDelType, _
    ByVal lngModule As Long, ByVal strPrivs As String, objDelBalance As clsCliniDelBalance, _
    ByVal cllDelPro As Collection, Optional ByVal strDefault���㷽ʽ As String = "", _
    Optional ByVal cllForceDelToCash As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������:��ʾ�����˷ѽ��㴰��
    '���:frmMain-���õ�������
    '       bytFunc-0- �˷�;1-���쳣�˷�
    '       lngModule -ģ���
    '       strPrivs-Ȩ�޴�
    '       objDelBalance-�˷���ؽ�����Ϣ
    '       cllDelPro-�˷�ǰ��Ҫִ�е�SQL
    '       strDefault���㷽ʽ-ȱʡ�Ľ��㷽ʽ
    '       cllForceDelToCash ǿ��������Ϣ��Array(����Ա,���������)
    '����:
    '����:����շ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-12 09:59:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mobjDelBalance = objDelBalance: Set mcllDelPro = cllDelPro
    mblnOK = False
    mblnUnLoad = False: mblnUnloaded = False
    mblnTurnFee = IsTurnFee(mobjDelBalance.AllNos)
    mstrPrivs = strPrivs: mlngModule = lngModule
    Set mfrmMain = frmMain
    mbytFunc = bytFunc
    mblnOK = False
    If cllForceDelToCash Is Nothing Then Set cllForceDelToCash = New Collection
    Set mcllForceDelToCash = cllForceDelToCash
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    Set objDelBalance = mobjDelBalance
    zlDelCharge = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function IsTurnFee(ByVal strNos As String) As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset
    strSQL = "Select 1" & vbNewLine & _
            " From ������˼�¼ A, ������ü�¼ B" & vbNewLine & _
            " Where a.����id = b.Id And b.No In (Select Column_Value From Table(f_Str2list([1])))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "����Ƿ�Ϊ����תסԺ����", strNos)
    If rsTmp.EOF Then
        IsTurnFee = False
    Else
        IsTurnFee = True
    End If
End Function

Private Sub initInsure()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҽ������
    '����:���˺�
    '����:2011-08-21 18:55:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjDelBalance.intInsure = 0 Then Exit Sub
    mInsurePara.���������շ� = gclsInsure.GetCapability(support���������շ�, mobjDelBalance.����ID, mobjDelBalance.intInsure)
    '���˺�:27536 20100119
    mInsurePara.�����ѽɿ���� = gclsInsure.GetCapability(support�����ѽɿ����, mobjDelBalance.����ID, mobjDelBalance.intInsure)
    mInsurePara.�ֱҴ��� = gclsInsure.GetCapability(support�ֱҴ���, mobjDelBalance.����ID, mobjDelBalance.intInsure)
End Sub

Private Sub InitBalanceData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����������
    '����:���˺�
    '����:2012-02-05 16:02:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call ClearBanalce
    With mCurCarge
        .dbl�˷Ѻϼ� = mobjDelBalance.�˷Ѻϼ�
        .dbl������� = 0
        .dbl����ҽ���˷� = 0
        .dbl���˺ϼ� = 0
        .dbl��ǰδ�� = .dbl�˷Ѻϼ� - .dbl����ҽ���˷�
        .dbl������Ԥ�� = 0
        .dbl�������� = 0
    End With
    
    '����ԭ����
    Call Loadԭ����
End Sub

Public Function IsSingleBalance(ByVal lngԭ����ID As Long) As Boolean
    '�жϵ��ݵ�һ�ν��ʳ�ҽ�����㷽ʽ���Ƿ�ֻʹ����һ�ֽ��㷽ʽ
    '��Σ�
    '   lngԭ����ID - ԭ����ID��ҽ�������˵�����һ�����յĽ���ID
    Dim rsBalance As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    '1.��һ�ν��ʵĽ���ID
    strSQL = "Select Distinct m.����id" & vbNewLine & _
            " From ������ü�¼ M, ������ü�¼ N" & vbNewLine & _
            " Where Mod(m.��¼����, 10) = Mod(n.��¼����, 10) And m.No = n.No" & vbNewLine & _
            "       And m.��¼���� = 1 And m.��¼״̬ In (1, 3) And n.����id = [1]"
    '2.��һ�ν���Ľ�����Ϣ
    strSQL = "With ԭʼ����id As(" & strSQL & ")" & vbNewLine & _
            " Select Decode(a.��¼����, 11, '��Ԥ��', a.���㷽ʽ) As ���㷽ʽ, a.��Ԥ��, c.����" & vbNewLine & _
            " From ����Ԥ����¼ A, ԭʼ����id B, ���㷽ʽ C" & vbNewLine & _
            " Where a.����id = b.����id And a.��¼���� In (11, 3) And a.���㷽ʽ = c.����(+) And c.Ӧ�տ� <> 1 And c.Ӧ���� <> 1"
    '3.��ҽ�����㷽ʽ����Ľ��㷽ʽ
    '���㷽ʽ.���ʣ�1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����(�ϰ�),8-���㿨����(�°�),9-����
    strSQL = "Select Distinct ���㷽ʽ From (" & strSQL & ") Where ���� Not In (3, 4, 9)"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSQL, "��ȡ��ҽ���Ľ��㷽ʽ", lngԭ����ID)
    IsSingleBalance = (rsBalance.RecordCount <= 1)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Loadԭ����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ԭ���˵Ľ��㷽ʽ
    '����:���˺�
    '����:2014-07-31 14:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBalance As ADODB.Recordset
    Dim strTemp As String, lngCurRecord As Long
    Dim i As Long, strCardNo As String, lng�����ID As Long
    Dim j As Long, ingCount As Integer
    Dim blnFind As Boolean
    Dim bln��ͨ���� As Boolean
    Dim varTemp As Variant, lng���ѿ�ID As Long
    Dim objCard As Card, dblMoney As Double
    
    On Error GoTo errHandle
    'ԭ���˵Ļ�,�Ƚ��˷ѵĽ����ʾ����
    'ע�⣺���һ��ͨ�����ѿ��������֣�ͬʱ�ÿ�δ�����򰴷ǵ��ֽ��㷽ʽ����104555��
    Set rsBalance = mobjDelBalance.rsBalance
    If rsBalance Is Nothing Then Exit Sub
    If rsBalance.State <> 1 Then Exit Sub
    
    mblnSingleBalance = IsSingleBalance(mobjDelBalance.ԭ����ID)
    
    If mblnSingleBalance Then
        rsBalance.Filter = "����<>2 And �������� <> 9 And �˷�=0 And ���㷽ʽ<>'" & mstr��֧Ʊ & "'"
        rsBalance.Sort = "ID Asc"
        If rsBalance.RecordCount > 0 Then
            rsBalance.MoveFirst
            If RoundEx(mCurCarge.dbl��ǰδ��, 6) <> 0 Then
                'δ�˽��Ϊ��ȱʡ�������㷽ʽû������
                If Val(Nvl(rsBalance!����)) = 1 Then
                    mobjDelBalance.ȱʡ���㷽ʽ = "��Ԥ���"
                Else
                    mobjDelBalance.ȱʡ���㷽ʽ = Trim(Nvl(rsBalance!���㷽ʽ))
                End If
            End If
            '3-һ��ͨ ���������Ҳ���ȫ�˵Ĳ����л��˿ʽ����������ȫ�˵����ڽ���������ǰ����
            '5-���ѿ����������ֵĲ����л��˿ʽ
            If Val(Nvl(rsBalance!����)) = 3 Or Val(Nvl(rsBalance!����)) = 5 Then
                If Val(Nvl(rsBalance!����)) = 3 Then
                    Set objCard = GetPayCard(Val(Nvl(rsBalance!�����ID)), False)
                Else
                    Set objCard = GetPayCard(Val(Nvl(rsBalance!���㿨���)), True)
                End If
                If Not objCard Is Nothing Then
                    If Val(Nvl(rsBalance!����)) = 5 Then
                        cbo֧����ʽ.Enabled = (Val(Nvl(rsBalance!�Ƿ�����)) = 1)
                    End If
                    '��������֣���δ���꣬���ܱ༭�˿ʽ
                    dblMoney = GetOldBalanceMoney(Val(Nvl(rsBalance!����)), objCard)
                    If RoundEx(dblMoney, 6) = 0 Then cbo֧����ʽ.Enabled = True
                End If
            End If
        End If
    Else
        '77873,Ƚ����,2014-9-15
        If mobjDelBalance.�����˷� Then
            '���ر���ȫ�˵����ѿ���һ��ͨ����
            '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
            '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
            rsBalance.Filter = "(�Ƿ�����=0 AND ����=3) or (�Ƿ�����=0 AND ����=5) OR �Ƿ�ȫ��=1"
            rsBalance.Sort = "�Ƿ����� asc,�Ƿ�ȫ�� desc"
            If rsBalance.RecordCount = 0 Then
                rsBalance.Filter = 0
                Exit Sub
            End If
        Else
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            rsBalance.Filter = "����=0 and ��������=1"
            If Not rsBalance.EOF Then
                 mobjDelBalance.ȱʡ���㷽ʽ = Trim(Nvl(rsBalance!���㷽ʽ))
            End If
            rsBalance.Filter = "����<>2 and ����<>4 "
            rsBalance.Sort = "���� desc,�������� desc"
        End If
    
        If rsBalance.RecordCount <> 0 Then rsBalance.MoveFirst
        With rsBalance
            bln��ͨ���� = False
            lngCurRecord = 1
            mobjDelBalance.ȱʡ���㷽ʽ = ""
            Do While Not .EOF
                strTemp = ""
                '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                Select Case Val(Nvl(!����))
                Case 0 '��ͨ����
                    If InStr(gTy_Module_Para.strȱʡ����, Trim(Nvl(!���㷽ʽ))) = 0 Then
                        strTemp = Trim(Nvl(!���㷽ʽ))
                    End If
                    bln��ͨ���� = True
                Case 1 'Ԥ����
                    strTemp = "��Ԥ���"
                    mCurCarge.dbl������Ԥ�� = RoundEx(mCurCarge.dbl������Ԥ�� + Val(Nvl(!��Ԥ��)), 6)
                Case 2 'ҽ��,������
                    'ҽ���Ѿ����˷�ǰ����
                Case 3 'һ��ͨ
                    Set objCard = GetPayCard(Val(Nvl(!�����ID)), False)
                    If objCard Is Nothing Then
                        strTemp = "" '��δ��������ǰ���ж�
                    ElseIf Not (objCard.�Ƿ����� And objCard.�Ƿ�ȱʡ����) Then
                        strTemp = Trim(Nvl(!���㷽ʽ))
                    End If
                Case 4 'һ��ͨ(��)
                Case 5  '���ѿ�
                    strTemp = Trim(Nvl(!���㷽ʽ))
                End Select
                
                If Val(Nvl(!����)) = 0 And Val(Nvl(!��������)) = 1 Then strTemp = "" '�ֽ𲻼�
                If Val(Nvl(rsBalance!��������)) = 9 Then strTemp = ""   '���Ѳ�����
                If strTemp <> "" Then
                    With vsBlance
                        i = 1
                        If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                             blnFind = False
                            If Nvl(rsBalance!����) = 1 Then
                                For j = 1 To .Rows - 1
                                    If Val(.TextMatrix(j, .ColIndex("����"))) = 1 Then
                                        blnFind = True
                                        i = j: Exit For
                                    End If
                                Next
                            ElseIf Nvl(rsBalance!����) = 5 Then
                                For j = 1 To .Rows - 1
                                    If strTemp = Trim(.TextMatrix(j, .ColIndex("֧����ʽ"))) _
                                        And Val(Nvl(rsBalance!���ѿ�ID)) = Val(.TextMatrix(j, .ColIndex("���ѿ�ID"))) Then
                                        blnFind = True
                                        i = j: Exit For
                                    End If
                                Next
                            Else
                                For j = 1 To .Rows - 1
                                    If strTemp = Trim(.TextMatrix(j, .ColIndex("֧����ʽ"))) Then
                                        blnFind = True
                                        i = j: Exit For
                                    End If
                                Next
                            
                            End If
                            If Not blnFind Then
                                .Rows = .Rows + 1
                                .RowPosition(.Rows - 1) = 1
                            End If
                        End If
                        
                        If Not (Val(.TextMatrix(i, .ColIndex("����״̬"))) = 1 And blnFind) Then  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                            strCardNo = Nvl(rsBalance!����)
                            If Nvl(rsBalance!����) = 5 Then
                                lng�����ID = Val(Nvl(rsBalance!���㿨���))
                                If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����,ʣ��δ�˽��)
                                mcllSquareBalance.Add Array(lng�����ID, Val(Nvl(rsBalance!���ѿ�ID)), _
                                 0, strCardNo, "", "", Val(Nvl(rsBalance!�Ƿ�����)), Format(Val(Nvl(rsBalance!��Ԥ��)), "0.00"))
                            Else
                                lng�����ID = Val(Nvl(rsBalance!�����ID))
                            End If
                            .RowData(i) = Nvl(rsBalance!����)
                            .TextMatrix(i, .ColIndex("����")) = Val(Nvl(rsBalance!����))
                            .TextMatrix(i, .ColIndex("��������")) = Val(Nvl(rsBalance!��������))
                            If Nvl(rsBalance!����) = 5 Then
                                .TextMatrix(i, .ColIndex("ɾ����־")) = IIf(Val(Nvl(rsBalance!�Ƿ�����)) = 1, 0, 1) '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                            Else
                                .TextMatrix(i, .ColIndex("ɾ����־")) = 0 '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                            End If
                            .TextMatrix(i, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                            .TextMatrix(i, .ColIndex("�����ID")) = lng�����ID
                            .TextMatrix(i, .ColIndex("���ѿ�ID")) = Val(Nvl(rsBalance!���ѿ�ID))
                            .TextMatrix(i, .ColIndex("֧����ʽ")) = strTemp
                            ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                            .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = lng�����ID & "|" & IIf(Val(Nvl(rsBalance!����)) = 5, 1, 0) & "|" & Val(Nvl(rsBalance!���ƿ�)) & "|" & Val(Nvl(rsBalance!�Ƿ�ȫ��)) & "|" & Val(Nvl(rsBalance!�Ƿ�����)) & "|" & Nvl(rsBalance!���������)
                            .TextMatrix(i, .ColIndex("֧�����")) = FormatEx(-1 * Val(.Cell(flexcpData, i, .ColIndex("֧�����"))) + Val(Nvl(rsBalance!��Ԥ��)), 6, , , 2)
                            .Cell(flexcpData, i, .ColIndex("֧�����")) = FormatEx(Val(.Cell(flexcpData, i, .ColIndex("֧�����"))) + -1 * Val(Nvl(rsBalance!��Ԥ��)), 6)
                            If Nvl(rsBalance!����) <> 1 Then 'Ԥ�����ʾ������롢ժҪ�����š�������ˮ�š�����˵��
                                .TextMatrix(i, .ColIndex("�������")) = Nvl(rsBalance!�������)
                                .TextMatrix(i, .ColIndex("��ע")) = Nvl(rsBalance!ժҪ)
                                .TextMatrix(i, .ColIndex("������ˮ��")) = Nvl(rsBalance!������ˮ��)
                                .TextMatrix(i, .ColIndex("����˵��")) = Nvl(rsBalance!����˵��)
                                .TextMatrix(i, .ColIndex("����")) = IIf(Val(Nvl(rsBalance!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                                .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(Nvl(rsBalance!�Ƿ�����))
                                .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(Nvl(rsBalance!�Ƿ�ȫ��))
                                .TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����")) = Val(Nvl(rsBalance!�Ƿ�ת�ʼ�����))
                                .TextMatrix(i, .ColIndex("���������")) = Nvl(rsBalance!���������)
                                .Cell(flexcpData, i, .ColIndex("����")) = Nvl(rsBalance!����)
                            End If
                            mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + -1 * Val(Nvl(rsBalance!��Ԥ��)), 6)
                        End If
                    End With
                End If
                .MoveNext
            Loop
            mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl�˷Ѻϼ� - mCurCarge.dbl���˺ϼ�, 6)
            
            '77873,Ƚ����,2014-9-15
            '85597,�����˷Ѻ󣬽�ʣ�ಿ��ȫ��ʱ֧����ʽĬ�Ͻ���ȷ
            '86248,�����˷�ʱ��Ϊ֧Ʊ���ڶ��ν�ʣ�ಿ��ȫ��ʱ��Ӧ��Ĭ��Ϊ��֧Ʊ
            With vsBlance
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                        If .Cell(flexcpData, i, .ColIndex("֧�����")) >= 0 Then '�տ��Ĭ����ʾ
                            mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� - .Cell(flexcpData, i, .ColIndex("֧�����")), 6)
                            mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� + .Cell(flexcpData, i, .ColIndex("֧�����")), 6)
                            .TextMatrix(i, .ColIndex("֧�����")) = 0
                            .Cell(flexcpData, i, .ColIndex("֧�����")) = 0
                        End If
                    End If
                Next
                For i = 1 To .Rows - 1
                    If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                        If mCurCarge.dbl��ǰδ�� > 0 Then
                            '93114,ȫ�˿����ֵĿ�ת�ʵ�һ��ͨ����ȱʡΪ��ǰ�˿���
                            If Val(.TextMatrix(i, .ColIndex("�Ƿ�ȫ��"))) = 1 _
                                And (Val(.TextMatrix(i, .ColIndex("����"))) <> 3 _
                                    Or (Val(.TextMatrix(i, .ColIndex("����"))) = 3 _
                                        And (Val(.TextMatrix(i, .ColIndex("�Ƿ�����"))) = 0 _
                                            Or Val(.TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����"))) = 0))) Then
                                Exit For
                            End If
                            If mCurCarge.dbl��ǰδ�� > -1 * .Cell(flexcpData, i, .ColIndex("֧�����")) Then
                                mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� - .Cell(flexcpData, i, .ColIndex("֧�����")), 6)
                                mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� + .Cell(flexcpData, i, .ColIndex("֧�����")), 6)
                                .TextMatrix(i, .ColIndex("֧�����")) = 0
                                .Cell(flexcpData, i, .ColIndex("֧�����")) = 0
                            Else
                                If .TextMatrix(i, .ColIndex("֧�����")) <> 0 Then
                                    '���⴦���շ�ʱ����Ľ����������λС��,��˴˴��˿���ҲҪ���������봦��
                                    '�統ǰδ��30.105���տ�ʱ֧�����Ϊ30.11��������Ƚ����������봦���ͻ���Ϊ�տ�ʱ��30.105
                                    mCurCarge.dbl������� = RoundEx(mCurCarge.dbl�˷Ѻϼ� - Format(mCurCarge.dbl�˷Ѻϼ�, "0.00"), 6)
                                    
                                    .TextMatrix(i, .ColIndex("֧�����")) = Format(.TextMatrix(i, .ColIndex("֧�����")) - (mCurCarge.dbl��ǰδ�� - mCurCarge.dbl�������), "0.00")
                                    .Cell(flexcpData, i, .ColIndex("֧�����")) = Format(.Cell(flexcpData, i, .ColIndex("֧�����")) + (mCurCarge.dbl��ǰδ�� - mCurCarge.dbl�������), "0.00")
                                    mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + (mCurCarge.dbl��ǰδ�� - mCurCarge.dbl�������), 6)
                                    mCurCarge.dbl��ǰδ�� = mCurCarge.dbl�������
                                    Exit For
                                End If
                            End If
                        Else
                            '93114,֧��ת���ҿ����ֵ�һ��ͨ��ȱʡΪ���еĿ�ת�ʽ��
                            '�ų����Ϊ��ģ����Ϊ���ʾҪ�Ƴ���
    '                        If Val(.TextMatrix(i, .ColIndex("����"))) = 3 And Val(.TextMatrix(i, .ColIndex("�Ƿ�ת�ʼ�����"))) = 1 _
    '                            And Val(.TextMatrix(i, .ColIndex("�Ƿ�����"))) = 1 And .TextMatrix(i, .ColIndex("֧�����")) <> 0 Then
    '                            .TextMatrix(i, .ColIndex("֧�����")) = Format(.TextMatrix(i, .ColIndex("֧�����")) - mCurCarge.dbl��ǰδ��, "0.00")
    '                            .Cell(flexcpData, i, .ColIndex("֧�����")) = .Cell(flexcpData, i, .ColIndex("֧�����")) + mCurCarge.dbl��ǰδ��
    '                            mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + mCurCarge.dbl��ǰδ��, 6)
    '                            mCurCarge.dbl��ǰδ�� = 0
    '                        End If
                        End If
                    End If
                Next
                
                i = 1
                Do While True
                    If Val(.TextMatrix(i, .ColIndex("֧�����"))) = 0 Then
                        '�Ƴ����Ϊ���֧���������δ��
                        lng�����ID = Val(.TextMatrix(i, .ColIndex("�����ID")))
                        lng���ѿ�ID = Val(.TextMatrix(i, .ColIndex("���ѿ�ID")))
                        Call ClearReMoveSquareBalance(lng�����ID, lng���ѿ�ID)
                        If .Rows <= 2 Then
                            .Rows = 2
                            .Clear 1
                            .RowData(1) = ""
                            .Cell(flexcpData, 1, 0, .Rows - 1, .COLS - 1) = ""
                            Exit Do
                        Else
                            .RemoveItem i
                        End If
                    Else
                        i = i + 1
                    End If
                    If i > .Rows - 1 Then Exit Do
                Loop
            End With
        End With
    End If
    mobjDelBalance.rsBalance.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


Private Sub ClearBanalce()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '����:���˺�
    '����:2012-02-05 16:02:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurCarge
        .dbl�˷Ѻϼ� = 0
        .dbl������� = 0
        .dbl����ҽ���˷� = 0
        .dbl���˺ϼ� = 0
        .dbl����Ӧ�� = 0
        .dbl��ǰδ�� = 0
        .dbl������Ԥ�� = 0
        .dbl�������� = 0
    End With
    With vsBlance
        .Clear 1: .Rows = 2
    End With
End Sub
Private Sub LoadData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2011-08-20 19:49:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, bln���ѿ� As Boolean, lng�����ID As Long
    Dim strCardNo As String
    Dim blnYb As Boolean
    Dim dbl��Ԥ���� As Double
    
    On Error GoTo errHandle
    
    Call ClearBanalce
    If mobjDelBalance.SaveBilled = False Then Exit Sub
    
    '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    'bytType-��������:0-���ݽ���ID����;1-���ݽ�����Ų���
    Set mrsBalance = zlFromIDGetChargeBalance(1, mobjDelBalance.�������, False, True)
    mrsBalance.Filter = 0
    mrsBalance.Sort = "����,���㷽ʽ"
    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    With mrsBalance
        i = 1: blnYb = False
        dbl��Ԥ���� = 0
        Do While Not .EOF
            Select Case Nvl(!����)
            Case 1 'Ԥ����
                mCurCarge.dbl������Ԥ�� = RoundEx(mCurCarge.dbl������Ԥ�� + Val(Nvl(!��Ԥ��)), 6)
                mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + Val(Nvl(!��Ԥ��)), 6)
                dbl��Ԥ���� = RoundEx(dbl��Ԥ���� + Val(Nvl(!��Ԥ��)), 6)
            Case 2, 3, 5 'ҽ��,һ��ͨ,���ѿ�
                If Nvl(!����) = 2 Then
                    mCurCarge.dbl����ҽ���˷� = RoundEx(mCurCarge.dbl����ҽ���˷� + Nvl(!��Ԥ��, 0), 6)
                    blnYb = True
                End If
'                If Val(Nvl(mrsBalance!У�Ա�־, 0)) = 2 Then
                    With vsBlance
                        If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then
                            .Rows = .Rows + 1
                            i = i + 1
                        End If
                        .RowData(i) = Nvl(mrsBalance!����)
                        If Nvl(mrsBalance!����) = 5 Then
                            lng�����ID = Val(Nvl(mrsBalance!���㿨���))
                        Else
                            lng�����ID = Val(Nvl(mrsBalance!�����ID))
                        End If
                        
                        strCardNo = Nvl(mrsBalance!����)
                        If Nvl(mrsBalance!����) = 5 Then
                            If Val(Nvl(mrsBalance!��Ԥ��)) <= 0 Then
                                If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
                                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
                                mcllSquareBalance.Add Array(lng�����ID, Val(Nvl(mrsBalance!���ѿ�ID)), _
                                Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00"), strCardNo, "", "", Val(Nvl(mrsBalance!�Ƿ�����)), Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00"))
                            Else
                                If mcllSquareChargeBalance Is Nothing Then Set mcllSquareChargeBalance = New Collection
                                'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
                                mcllSquareChargeBalance.Add Array(lng�����ID, Val(Nvl(mrsBalance!���ѿ�ID)), _
                                Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00"), strCardNo, "", "", Val(Nvl(mrsBalance!�Ƿ�����)), Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00"))
                            End If
                        End If
                        .TextMatrix(i, .ColIndex("����")) = Val(Nvl(mrsBalance!����))
                        .TextMatrix(i, .ColIndex("��������")) = Val(Nvl(mrsBalance!��������))
                        .TextMatrix(i, .ColIndex("ɾ����־")) = 1  '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                        .TextMatrix(i, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                        .TextMatrix(i, .ColIndex("�����ID")) = lng�����ID
                        .TextMatrix(i, .ColIndex("���ѿ�ID")) = Val(Nvl(mrsBalance!���ѿ�ID))

                        .TextMatrix(i, .ColIndex("֧����ʽ")) = Nvl(mrsBalance!���㷽ʽ)
                        
                        ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                        .Cell(flexcpData, i, .ColIndex("֧����ʽ")) = lng�����ID & "|" & IIf(Val(Nvl(mrsBalance!����)) = 5, 1, 0) & "|" & Val(Nvl(mrsBalance!���ƿ�)) & "|" & Val(Nvl(mrsBalance!�Ƿ�ȫ��)) & "|" & Val(Nvl(mrsBalance!�Ƿ�����)) & "|" & Nvl(mrsBalance!���������)
                        .TextMatrix(i, .ColIndex("֧�����")) = Format(-1 * Val(Nvl(mrsBalance!��Ԥ��)), "0.00")
                        .Cell(flexcpData, i, .ColIndex("֧�����")) = Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00")
                        .TextMatrix(i, .ColIndex("�������")) = Nvl(mrsBalance!�������)
                        .TextMatrix(i, .ColIndex("��ע")) = Nvl(mrsBalance!ժҪ)
                        .TextMatrix(i, .ColIndex("������ˮ��")) = Nvl(mrsBalance!������ˮ��)
                        .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsBalance!����˵��)
                        .TextMatrix(i, .ColIndex("����")) = IIf(Val(Nvl(mrsBalance!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .TextMatrix(i, .ColIndex("�Ƿ�����")) = Val(Nvl(mrsBalance!�Ƿ�����))
                        .TextMatrix(i, .ColIndex("�Ƿ�ȫ��")) = Val(Nvl(mrsBalance!�Ƿ�ȫ��))
                        .TextMatrix(i, .ColIndex("���������")) = Nvl(mrsBalance!���������)
  
                        
                        .Cell(flexcpData, i, .ColIndex("����")) = Nvl(mrsBalance!����)
                        .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
                        mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + Val(Nvl(mrsBalance!��Ԥ��)), 6)
                    End With
'                End If
            Case Else '0-��ͨ����
                With vsBlance
                   If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" And Nvl(mrsBalance!���㷽ʽ) <> "" Then
                       .Rows = .Rows + 1
                       i = i + 1
                   End If
                   If Trim(Nvl(mrsBalance!���㷽ʽ)) <> "" Then
                        .RowData(i) = Nvl(mrsBalance!����)
                        
                        .TextMatrix(i, .ColIndex("����")) = Val(Nvl(mrsBalance!����))
                        .TextMatrix(i, .ColIndex("��������")) = Val(Nvl(mrsBalance!��������))
                        .TextMatrix(i, .ColIndex("ɾ����־")) = 1  '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                        .TextMatrix(i, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                        .TextMatrix(i, .ColIndex("�����ID")) = 0
                        .TextMatrix(i, .ColIndex("���ѿ�ID")) = 0

                        
                        .TextMatrix(i, .ColIndex("֧����ʽ")) = Nvl(mrsBalance!���㷽ʽ)
                        .TextMatrix(i, .ColIndex("֧�����")) = Format(-1 * Val(Nvl(mrsBalance!��Ԥ��)), "0.00")
                        .Cell(flexcpData, i, .ColIndex("֧�����")) = Format(Val(Nvl(mrsBalance!��Ԥ��)), "0.00")
                        .TextMatrix(i, .ColIndex("�������")) = Nvl(mrsBalance!�������)
                        .TextMatrix(i, .ColIndex("��ע")) = Nvl(mrsBalance!ժҪ)
                        .TextMatrix(i, .ColIndex("������ˮ��")) = Nvl(mrsBalance!������ˮ��)
                        .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsBalance!����˵��)
                        .TextMatrix(i, .ColIndex("����")) = IIf(Val(Nvl(mrsBalance!�Ƿ�����)) = 1, String(Len(strCardNo), "*"), strCardNo)
                        .Cell(flexcpData, i, .ColIndex("����")) = Nvl(mrsBalance!����)
                        .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
                        mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + Val(Nvl(mrsBalance!��Ԥ��)), 6)
                    End If
                End With
            End Select
            .MoveNext
        Loop
    End With
    '�ȼ�����˷Ѻϼ�
    gstrSQL = "" & _
    "   Select B.NO,B.����ID, Nvl(Sum(Nvl(B.Ӧ�ս��, 0)), 0)  As ����Ӧ�պϼ�, " & _
    "       Nvl(Sum(Nvl(B.ʵ�ս��, 0)), 0)  As ����ʵ�պϼ� " & _
    "   From ������ü�¼ B " & _
    "    Where B.����ID=[1] Or B.����ID=[2]" & _
    "    Group by B.NO,B.����ID"
   Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.����ID, mobjDelBalance.����ID)
   With mCurCarge
         .dbl�˷Ѻϼ� = 0:
         .dbl����Ӧ�� = 0
        Do While Not rsTemp.EOF
            .dbl�˷Ѻϼ� = RoundEx(.dbl�˷Ѻϼ� + Val(Nvl(rsTemp!����ʵ�պϼ�)), 6)
            .dbl����Ӧ�� = RoundEx(.dbl����Ӧ�� + Val(Nvl(rsTemp!����Ӧ�պϼ�)), 6)
            rsTemp.MoveNext
        Loop
        .dbl��ǰδ�� = RoundEx(.dbl�˷Ѻϼ� - .dbl���˺ϼ�, 6)
    End With
    Call Loadԭ����
                   
    If dbl��Ԥ���� <> 0 Then
        With vsBlance
            If .Rows = 2 Then .Row = 1
            If .Row < 0 Then .Row = 1
            i = .Row
            If Trim(.TextMatrix(.Row, .ColIndex("֧����ʽ"))) <> "" Then
                .Rows = .Rows + 1
                i = .Rows - 1
            End If
            .RowData(i) = 1
            .TextMatrix(i, .ColIndex("ɾ����־")) = 1   ' �Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
            .TextMatrix(i, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
            .TextMatrix(i, .ColIndex("֧����ʽ")) = "��Ԥ���"
            .TextMatrix(i, .ColIndex("֧�����")) = Format(-1 * mCurCarge.dbl������Ԥ��, "0.00")
            .Cell(flexcpData, i, .ColIndex("֧�����")) = Format(mCurCarge.dbl������Ԥ��, "0.00")
            .TextMatrix(i, .ColIndex("����")) = 1
            
            .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
        End With
    End If
   vsBlance_AfterRowColChange 0, 0, vsBlance.Row, vsBlance.Col
   Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Init�˷ѷ�ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:���˺�
    '����:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, objCard As Card, objCards As Cards
    Dim lngKey As Long
    
    Set mobjPayCards = New Cards
    Set objCards = New Cards
    
    Set rsTemp = mobjDelBalance.rs���㷽ʽ
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
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
            blnFind = False
            For i = 1 To objCards.Count
                If objCards(i).���㷽ʽ = Nvl(rsTemp!����) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!����)) = 3 Or Val(Nvl(rsTemp!����)) = 4 _
                    Or Val(Nvl(rsTemp!����)) = 7 Or Val(Nvl(rsTemp!����)) = 8 _
                    Or Val(Nvl(rsTemp!Ӧ����)) = 1) Then
                    
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
                     If objCard.�������� = 7 And objCard.�ӿ���� <= 0 Then   'һ��ͨδ����ʱ,������
                        mrsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
                        If Not mrsOneCard.EOF Then
                            mobjPayCards.Add objCard, "K" & lngKey
                            lngKey = lngKey + 1
                        End If
                     Else
                        mobjPayCards.Add objCard, "K" & lngKey
                        lngKey = lngKey + 1
                     End If
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '��������
    For i = 1 To objCards.Count
        rsTemp.Filter = "����='" & objCards(i).���㷽ʽ & "'" '���㷽ʽҪ������"����"Ӧ�ó��ϲ���ʹ��
        If Not rsTemp.EOF Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.Count = 0 Then
        MsgBox "û�п��õĽ��㷽ʽ,���ȵ����㷽ʽ���������á�", vbExclamation, gstrSysName
        mblnUnLoad = True: Exit Sub
    End If
    
    '���Ƽ���Ԥ�����
     Set objCard = New Card
     objCard.���� = "Ԥ"
     objCard.�ӿڱ��� = ""
     objCard.�ӿڳ����� = ""
     objCard.�ӿ���� = -1 * lngKey
     objCard.���㷽ʽ = "Ԥ����"
     objCard.���� = "Ԥ����"
     objCard.���� = True
     objCard.ȱʡ��־ = False
     objCard.֧������ = True
     objCard.�������� = "-99"
     mobjPayCards.Add objCard, "K" & lngKey
End Sub

Private Sub StartAndStopԤ���()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ���˿����̬����Ԥ����֧����ʽ
    '����:���˺�
    '����:2014-07-08 15:21:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBalance As ADODB.Recordset
    Dim objCard As Card
    Dim blnStart As Boolean, i As Long, dblMoney As Double
    
    Set rsBalance = mobjDelBalance.rsBalance
    '��Ԥ���
    '114528,��ǰδ��ֻ�������ʱӦ�����˿��Ӧ�ó����տ�
    If RoundEx(mCurCarge.dbl��ǰδ�� - mCurCarge.dbl��������, 6) <= 0 Then
        '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        rsBalance.Filter = "����=1"
        If rsBalance.RecordCount > 0 Then
            dblMoney = 0
            Do While Not rsBalance.EOF
                dblMoney = dblMoney + Nvl(rsBalance!��Ԥ��)
                rsBalance.MoveNext
            Loop
            dblMoney = RoundEx(dblMoney, 6)
            If RoundEx(dblMoney, 6) <> 0 Then
                For i = 1 To mobjPayCards.Count
                   Set objCard = mobjPayCards(i)
                   If objCard.�������� = -99 Then
                      objCard.���㷽ʽ = "��Ԥ���"
                      objCard.���� = "��Ԥ���"
                      objCard.֧������ = True
                   End If
                Next
            End If
        End If
    Else '��Ԥ���
        For i = 1 To mobjPayCards.Count
           Set objCard = mobjPayCards(i)
           If objCard.�������� = -99 Then
              objCard.���㷽ʽ = "��Ԥ���"
              objCard.���� = "��Ԥ���"
              objCard.֧������ = True
           End If
        Next
    End If
    
    blnStart = True
    With vsBlance
        For i = 1 To .Rows - 1
             If Val(.TextMatrix(i, .ColIndex("����"))) = 1 Then
                blnStart = False: Exit For
             End If
        Next
    End With
    If Not blnStart Then
        For i = 1 To mobjPayCards.Count
           Set objCard = mobjPayCards(i)
           If objCard.�������� = -99 Then objCard.֧������ = False
        Next
    End If
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ���ؼ�
    '����:���˺�
    '����:2011-06-13 14:09:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl��� As Double, rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    With vsBlance
        .Cell(flexcpFontBold, 1, 0, 1, .COLS - 1) = True
        .Clear: .Rows = 2: i = 0: .COLS = 18
        .TextMatrix(0, i) = "�����ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "���ѿ�ID": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "��������": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "֧����ʽ": .ColWidth(i) = 2000: i = i + 1
        .TextMatrix(0, i) = "֧�����": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "�������": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "��ע": .ColWidth(i) = 2500: i = i + 1
        .TextMatrix(0, i) = "����": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "������ˮ��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "����˵��": .ColWidth(i) = 1400: i = i + 1
        .TextMatrix(0, i) = "ɾ����־": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "����״̬": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ȫ��": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ�ת�ʼ�����": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "���������": .ColWidth(i) = 0: i = i + 1
        .TextMatrix(0, i) = "�Ƿ���֤": .ColWidth(i) = 0: i = i + 1 '�����ж�Ԥ�����Ƿ�����֤
        
        For i = 0 To .COLS - 1
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = flexAlignLeftCenter
            If .ColKey(i) Like "*ID" Then .ColHidden(i) = True
            Select Case .ColKey(i)
            Case "��������", "����", "ɾ����־", "�Ƿ�����", "�Ƿ�ȫ��", "�Ƿ�ת�ʼ�����", "���������", "����״̬", "�Ƿ���֤"
                .ColHidden(i) = True
            Case "֧�����"
                .ColAlignment(i) = flexAlignRightCenter
            End Select
        Next
    End With
    With mCurCarge
        .dbl������Ԥ�� = 0
        .dbl�˷Ѻϼ� = 0
        .dbl������� = 0
        .dbl����ҽ���˷� = 0
        .dbl���˺ϼ� = 0
        .dbl����Ӧ�� = 0
        .dbl��ǰδ�� = 0
        .dbl������� = 0
        .dbl����Ԥ�� = 0
        .dblԤ����� = 0
    End With
    
    mstr��֧Ʊ = ""
    If mobjDelBalance.rs���㷽ʽ Is Nothing Then
        Set mobjDelBalance.rs���㷽ʽ = Get���㷽ʽ("�շ�")
    ElseIf mobjDelBalance.rs���㷽ʽ.State <> 1 Then
        Set mobjDelBalance.rs���㷽ʽ = Get���㷽ʽ("�շ�")
    End If
    mobjDelBalance.rs���㷽ʽ.Filter = "Ӧ����=1"
    If Not mobjDelBalance.rs���㷽ʽ.EOF Then
         mstr��֧Ʊ = Nvl(mobjDelBalance.rs���㷽ʽ!����)
    End If
    mobjDelBalance.rs���㷽ʽ.Filter = 0
    Call initInsure
    Call Init�˷ѷ�ʽ
    
    If mbytFunc = EM_FUN_�˷� And mcllDelPro.Count <> 0 Then
        '����δ����ʱ,���
        Call InitBalanceData
    Else
        Call LoadData
    End If
    Call SetDeleteVisible '����������ʱɾ����ťӦ�ø��������ʾ
    
    Call Load�˷ѷ�ʽ: Call LoadPatiInfor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetOldBalanceMoney(ByVal int���� As Integer, ByVal objCard As Card) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͣ�ȷ��ԭ���㷽ʽ�Ľ��
    '���:int����-����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    '����:����ԭ������
    '����:���˺�
    '����:2014-07-08 15:49:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, i As Integer, blnFindByList As Boolean
    
    On Error GoTo errHandle
    With mobjDelBalance
        If .rsBalance Is Nothing Then Exit Function
        If .rsBalance.State <> 1 Then Exit Function
        .rsBalance.Filter = ""
        
        '93114���˷�ʱʹ��ת�ʷ�ʽ
        If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.ԭ����ID) Then
            '�����ת�ʽ��
            Do While Not .rsBalance.EOF
                Select Case Val(Nvl(.rsBalance!����))
                Case 0, 1, 4 '��ͨ����,Ԥ����,��һ��ͨ
                    dblMoney = dblMoney + Val(Nvl(.rsBalance!��Ԥ��))
                Case 3, 5 'һ��ͨ,���ѿ�
                    If Val(Nvl(.rsBalance!�Ƿ�����)) = 1 Then
                        dblMoney = dblMoney + Val(Nvl(.rsBalance!��Ԥ��))
                    End If
                End Select
                .rsBalance.MoveNext
            Loop
            
            '��ȥ���˿��ҽ���Ľ��
            For i = 1 To vsBlance.Rows - 1
                Select Case Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("����")))
                Case 0, 1, 4
                    dblMoney = dblMoney - Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("֧�����")))
                Case 3, 5
                    If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("����"))) = 3 And Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("�����ID"))) = objCard.�ӿ���� Then
                        '���б��еĽ��н���
                        blnFindByList = True
                    Else
                        If Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("�Ƿ�����"))) = 1 Then
                            dblMoney = dblMoney - Val(vsBlance.TextMatrix(i, vsBlance.ColIndex("֧�����")))
                        End If
                    End If
                End Select
            Next
            
            If blnFindByList Then
                If dblMoney > -1 * mCurCarge.dbl�˷Ѻϼ� Then dblMoney = -1 * mCurCarge.dbl�˷Ѻϼ�
            Else
                If dblMoney > -1 * mCurCarge.dbl��ǰδ�� Then dblMoney = -1 * mCurCarge.dbl��ǰδ��
            End If
            If dblMoney < 0 Then dblMoney = 0
            GetOldBalanceMoney = RoundEx(dblMoney, 6)
            Exit Function
        End If
       
        '77338,Ƚ����,2014-9-1,û����ȷ��ȡԤ������
        If objCard.�ӿ���� > 0 Then
            If objCard.���ѿ� = False Then 'һ��ͨ
                .rsBalance.Filter = "����=" & int���� & " And �����ID=" & objCard.�ӿ����
            Else '���ѿ�
                .rsBalance.Filter = "����=" & int���� & " And ���㿨���=" & objCard.�ӿ����
            End If
        ElseIf objCard.�������� = 2 And objCard.���㷽ʽ Like "*��" Then '87532
            .rsBalance.Filter = "����=" & int���� & " And ���㷽ʽ='" & objCard.���㷽ʽ & "'"
        Else
            .rsBalance.Filter = "����=" & int����
        End If
        If .rsBalance.EOF Then
            .rsBalance.Filter = 0
            Exit Function
        End If
        .rsBalance.MoveFirst
        Do While Not .rsBalance.EOF
            dblMoney = dblMoney + Val(Nvl(.rsBalance!��Ԥ��))
            .rsBalance.MoveNext
        Loop
        GetOldBalanceMoney = RoundEx(dblMoney, 6)
        .rsBalance.Filter = 0
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Sub SetControlProperty(Optional ByVal blnLoadDefault As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����
    '���:blnLoadDefault-�Ƿ����ȱʡֵ
    '����:���˺�
    '����:2014-07-10 17:49:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngTop As Long, sngSplitHeight As Single, dbl�ֽ� As Double
    Dim bln�ֱ� As Boolean, dblMoney As Double, dblTemp As Double
    Dim bln�˿� As Boolean '��Ҫ��ҽ����ؽ�������˵����շ�
    Dim blnVisible As Boolean, blnEnabled As Boolean
    Dim objCard As Card, intIndex As Integer
    Dim blnDel As Boolean
    
    blnDel = mCurCarge.dbl��ǰδ�� <= 0

    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    
    sngSplitHeight = 80
    lbl���˺ϼ�.Caption = "�Ѹ��ϼ�:" & Format(Abs(mCurCarge.dbl���˺ϼ�), "###0.00;-###0.00;0.00;0.00;")
    
    If objCard.�������� = 1 Then
        If RoundEx(mCurCarge.dbl�������, 6) = RoundEx(mCurCarge.dbl��ǰδ��, 6) Then
            dbl�ֽ� = 0
        Else
            dblMoney = mCurCarge.dbl��ǰδ��
            If mobjDelBalance.intInsure > 0 Then
                If mInsurePara.�ֱҴ��� Then
                    bln�ֱ� = True
                    dbl�ֽ� = CentMoney(CCur(dblMoney))
                Else
                    dbl�ֽ� = Format(dblMoney, "0.00")
                End If
            Else
                bln�ֱ� = True
                dbl�ֽ� = RoundEx(CentMoney(CCur(dblMoney)), 6)
            End If
        End If
        lblδ�˽��.Caption = Format(Abs(dbl�ֽ�), "0.00")
    Else
        lblδ�˽��.Caption = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
    End If
    If blnDel Then
        stcCurDelTitle.Caption = "��ǰӦ��"
        lblδ�˽��.ForeColor = vbRed
        lblPayType.Caption = "��  ��"
        lbl�Ҳ�.Caption = "��  ��"
    Else
        stcCurDelTitle.Caption = "��ǰӦ��"
        lblδ�˽��.ForeColor = vbBlue
        lblPayType.Caption = "��  ��"
        lbl�Ҳ�.Caption = "��  ��"
    End If
    
    '������ҽ�������һ��ͨ���ϰ�һ��ͨ
    '1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
    '77353,Ƚ����,2014-9-1,�˷�ʱ�տ�,�����˴���Ԥ����,����ʹ��Ԥ������нɿ�,ѡ������ʵĽ��㷽ʽʱ��������롹,��ժҪ������������
    blnEnabled = InStr(",1,3,4,5,6,-99,", "," & objCard.�������� & ",") = 0
    txt�������.Enabled = blnEnabled
    txtժҪ.Enabled = blnEnabled
                
    'ȱʡ��������
    If blnLoadDefault Then
        '77324,Ƚ����,2014-9-1,���������˻���������ʱ,Ӧ�ý�ֹ¼���˿���,ֻ�ܰ�����ȡ�Ľ��Ĭ��
        txt�ɿ�.Locked = False
        If objCard.�ӿ���� > 0 Then          '������������ѿ�
            '���ܳ������˽��
            If mCurCarge.dbl��ǰδ�� <= 0 Then
                 '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                dblTemp = GetOldBalanceMoney(IIf(objCard.���ѿ�, 5, 3), objCard)
                If objCard.�Ƿ�ȫ�� Then
                    txt�ɿ�.Text = FormatEx(dblTemp, 6, , , 2)
                    txt�ɿ�.Locked = True
                Else
                    If dblTemp >= Abs(mCurCarge.dbl��ǰδ��) Then
                        txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
                    Else
                        txt�ɿ�.Text = FormatEx(dblTemp, 6, , , 2)
                    End If
                    txt�ɿ�.Locked = objCard.�Ƿ����� = False
                End If
            Else
                txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
            End If
        ElseIf objCard.�������� = 7 And objCard.�ӿ���� <= 0 Then '��һ��ͨ
             '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If mCurCarge.dbl��ǰδ�� <= 0 Then
                dblTemp = GetOldBalanceMoney(4, objCard)
                If dblTemp >= Abs(mCurCarge.dbl��ǰδ��) Then
                    txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
                Else
                    txt�ɿ�.Text = FormatEx(dblTemp, 6, , , 2)
                End If
'                txt�ɿ�.Locked = True
            Else
                txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
            End If
        ElseIf objCard.�������� = -99 Then  '��Ԥ��
             '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If mCurCarge.dbl��ǰδ�� <= 0 Then
                dblTemp = GetOldBalanceMoney(1, objCard)
                If dblTemp >= Abs(mCurCarge.dbl��ǰδ��) Then
                    txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
                Else
                    txt�ɿ�.Text = FormatEx(dblTemp, 6, , , 2)
                End If
            Else
                txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
            End If
        ElseIf objCard.�������� = 1 Then    '�ֽ���
            If gTy_Module_Para.bln�ֽ��˿�ȱʡ��ʽ Then
                txt�ɿ�.Text = Format(Abs(dbl�ֽ�), "0.00")
            Else
                txt�ɿ�.Text = "0.00"
            End If
        ElseIf objCard.�������� = 2 And objCard.���㷽ʽ Like "*��" Then  '��ҽ�����㷽ʽΪ"***��",87532
            If mCurCarge.dbl��ǰδ�� <= 0 Then
                dblTemp = GetOldBalanceMoney(0, objCard)
                If dblTemp >= Abs(mCurCarge.dbl��ǰδ��) Then
                    txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
                Else
                    txt�ɿ�.Text = FormatEx(dblTemp, 6, , , 2)
                End If
            Else
                txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
            End If
        Else
            '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            txt�ɿ�.Text = Format(Abs(mCurCarge.dbl��ǰδ��), "0.00")
        End If
    End If
    
    '�����Ҳ�
    If blnDel Then
        dblTemp = RoundEx(Val(txt�ɿ�.Text) - (-1 * mCurCarge.dbl��ǰδ�� + mCurCarge.dbl��������), 6)
        If dblTemp > 0 Then
            txt�Ҳ�.Text = Format(dblTemp, "0.00")
            txt�Ҳ�.ForeColor = vbRed
        Else
            txt�Ҳ�.Text = ""
        End If
    Else
        dblTemp = Val(txt�ɿ�.Text) - mCurCarge.dbl��ǰδ��
        txt�Ҳ�.ForeColor = lbl�Ҳ�.ForeColor
        If dblTemp > 0 Then
            txt�Ҳ�.Text = Format(dblTemp, "0.00")
        End If
    End If
    Call SetControlColor
End Sub

Private Sub cbo֧����ʽ_Click()
    Dim intIndex As Integer
    Dim objCard As Card, i As Integer
    Dim intSelectIndex As Integer
    
    If mblnFirst Then Exit Sub
    If mblnNotClick Then Exit Sub
    If mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex) Then Exit Sub
    
    '105432
    If mlngPre֧����ʽ > 0 And Val(txt�ɿ�.Text) <> 0 Then
        '��������շѽ��㷽ʽ�оͲ��ü�飬��Ҫ���֧�֡�ת�ʼ����ۡ���
        Set objCard = mobjPayCards(mlngPre֧����ʽ)
        mobjDelBalance.rsBalance.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "' And �˷�=0"
        
        If Not mobjDelBalance.rsBalance.EOF Then
            mblnNotClick = True
            intSelectIndex = cbo֧����ʽ.ListIndex
            cbo֧����ʽ.ListIndex = cbo.FindIndex(cbo֧����ʽ, mlngPre֧����ʽ)
            If ThreeBalanceCheck(Me, mlngModule, mobjPayCards(mlngPre֧����ʽ), _
                  mcllForceDelToCash, cbo֧����ʽ.Text) = False Then mblnNotClick = False: Exit Sub
            cbo֧����ʽ.ListIndex = intSelectIndex
            mblnNotClick = False
        End If
    End If
    
    mlngPre֧����ʽ = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    txt�ɿ�.Text = ""
    If cbo֧����ʽ.ListIndex < 0 Then GoTo SetProperty:
    
    intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    Set objCard = mobjPayCards(intIndex)
    '�л�������Ҫ���
    If objCard.�ӿ���� > 0 And objCard.���ѿ� = False Then
        For i = 1 To mcllForceDelToCash.Count
            If mcllForceDelToCash(i)(1) = objCard.���� Then Exit For
        Next
        If i <= mcllForceDelToCash.Count Then mcllForceDelToCash.Remove i
    End If
    
    If objCard.�������� = 7 And objCard.�ӿ���� <= 0 Then '��һ��ͨ
        If mobjICCard Is Nothing Then
            Set mobjICCard = New clsICCard
            Call mobjICCard.SetParent(Me.hWnd)
            Set mobjICCard.gcnOracle = gcnOracle
        End If
'    ElseIf objCard.���ѿ� Then
'         If IsExistSquare = True Then
'            If MsgBox("�Ѿ�����" & cbo֧����ʽ.Text & ",�Ƿ�ɾ��?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
'                Call ClearSquareBalance
'                Set mcllSquareBalance = Nothing
'            End If
'         End If
    End If
SetProperty:
     Call SetControlProperty(True)
     If txt�ɿ�.Enabled Then txt�ɿ�.SetFocus
End Sub
Private Function IsExistSquare(ByVal lngCardTypeID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�ж��Ƿ�������ѿ�����
    '���:lngCardTypeID-���ѿ����
    '����:
    '����:���ڳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-12 11:28:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, varTemp As Variant
    If mcllSquareBalance Is Nothing Then Exit Function
    For i = 1 To mcllSquareBalance.Count
        varTemp = mcllSquareBalance(i)
        ' array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
        If Val(varTemp(0)) = lngCardTypeID Then
            IsExistSquare = True
            Exit Function
        End If
    Next
End Function

Private Function CheckOneCard(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���һ��ͨ�Ƿ���ȷ
    '����:һ��ͨ��֤��ȷ���һ��ͨ,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-23 17:07:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim CurOneCard As Currency, dblMoney As Double, dblTemp As Double
    Dim strTittle As String, strCardNo As String
    
    If objCard.�������� <> 7 Then CheckOneCard = True: Exit Function
    
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        Set mobjICCard.gcnOracle = gcnOracle
    End If
    If mobjICCard Is Nothing Then
        MsgBox "һ��ͨ�ӿڴ���ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
    
    dblMoney = Val(txt�ɿ�.Text)
    If strTittle = "�ɿ�" Then
        CurOneCard = mobjICCard.GetSpare
        If CurOneCard < dblMoney Then
            MsgBox "������֧��,����!" & vbCrLf & vbCrLf & _
            "   �� ��  ��" & Format(CurOneCard, "0.00") & vbCrLf & _
            "   ����֧��" & FormatEx(Val(txt�ɿ�.Text), 6), vbInformation, gstrSysName
            Exit Function
        End If
        stbThis.Panels(4).Text = Format(CurOneCard, "0.00")
        stbThis.Panels(4).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(CurOneCard, "0.00")
        CheckOneCard = True
        Exit Function
    End If
    
     '�˿���
    If mobjDelBalance.rsBalance Is Nothing Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelBalance.rsBalance.State <> 1 Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    mobjDelBalance.rsBalance.Filter = "����=4"
    If mobjDelBalance.rsBalance.EOF Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo = "" Then
        MsgBox "һ��ͨ����ʧ��,�뽫IC�����ڶ�������", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    If strCardNo <> Nvl(mobjDelBalance.rsBalance!����) Then
        MsgBox "��ǰ������ۿ�Ų�һ��,���ܽ����˷�.", vbInformation, gstrSysName
        Exit Function
    End If
    
    dblTemp = Format(Val(Nvl(mobjDelBalance.rsBalance!��Ԥ��)), "0.00")
    If RoundEx(dblMoney, 6) <> Format(dblTemp, "0.00") Then
        MsgBox "һ��ͨ�������ȫ��,����!" & vbCrLf & vbCrLf & _
        "   ������" & Format(dblTemp, "0.00") & vbCrLf & _
        "   ����֧��" & Format(dblMoney, "0.00"), vbInformation, gstrSysName
        Exit Function
    End If
    CheckOneCard = True
End Function

Private Function CheckThreeSwapValied(ByVal objCard As Card, Optional dblDelMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������֤
    '���:objCard-������
    '     dblDelMoney-�˿���
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:00:34
    '˵��:ͬ����֤�˽ӿں�ˢ����Ʒ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl�ʻ���� As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln���� As Boolean
    Dim strBalanceIDs As String
    
    On Error GoTo errHandle
    If objCard Is Nothing Then
        If GetCurCard(objCard) = False Then
            MsgBox "��ǰ" & lblPayType.Caption & "��ʽδѡ��,��ѡ��!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then CheckThreeSwapValied = True: Exit Function
    
    If mCurCarge.dbl��ǰδ�� <= 0 Or dblDelMoney <> 0 Then
        strTittle = "�˿�"
    Else
        strTittle = "�ɿ�"
    End If
    
    mCurBrushCard = strBrushCard
    If dblDelMoney = 0 Then
        If Val(txt�ɿ�.Text) = 0 Then
            MsgBox strTittle & "���δ����,����!", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        End If
    End If
    
    If strTittle = "�ɿ�" Then
        If Abs(Val(txt�ɿ�.Text)) > Format(Abs(mCurCarge.dbl��ǰδ��), "0.00") And Val(txt�ɿ�.Text) <> 0 Then
            MsgBox strTittle & "���ܴ��ڱ���δ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
            Exit Function
        End If
        Set cllSquareBalance = Nothing
        Set mcllCurSquareBalance = Nothing
        
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
           '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
        
        dblMoney = Val(txt�ɿ�.Text)
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
            objCard.�ӿ����, objCard.���ѿ�, _
            mobjDelBalance.����, mobjDelBalance.�Ա�, mobjDelBalance.����, dblMoney, _
            mCurBrushCard.str����, mCurBrushCard.str����, _
            False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>0</CZLX></IN>") = False Then Exit Function
            '����ǰ,һЩ���ݼ��
            'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal strCardTypeID As Long, ByVal strCardNo As String, _
            ByVal dblMoney As Double, ByVal strNOs As String, _
            Optional ByVal strXMLExpend As String
            'mobjDelBalance.strNOs:��������ʱ,û�����ʱ,����Ϊ��.
            If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, objCard.�ӿ����, _
                objCard.���ѿ�, mCurBrushCard.str����, dblMoney, mobjDelBalance.CurDelNos, strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
            '���:frmMain-���õ�������
            '        lngModule-ģ���
            '        strCardNo-����
            '        strExpand-Ԥ����Ϊ��,�Ժ���չ
            '����:dblMoney-�����ʻ����
            If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, objCard.�ӿ����, _
                  mCurBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�) = False Then Exit Function
        
        stbThis.Panels(4).Text = Format(dbl�ʻ����, "0.00")
        stbThis.Panels(4).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
        mCurBrushCard.dbl�ʻ���� = RoundEx(dbl�ʻ����, 2)
        If dbl�ʻ���� <> 0 And dbl�ʻ���� < dblMoney Then
            MsgBox objCard.���㷽ʽ & "���ʻ�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        CheckThreeSwapValied = True
        Exit Function
    End If
    
    '�˿���
    If mobjDelBalance.rsBalance Is Nothing Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,���ܽ����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelBalance.rsBalance.State <> 1 Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,���ܽ����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '93114���˷�ʱʹ��ת�ʷ�ʽ
    If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.ԭ����ID) Then
        If dblDelMoney <> 0 Then
            dblMoney = dblDelMoney
        Else
            dblMoney = Val(txt�ɿ�.Text)
        End If
        dblTemp = GetOldBalanceMoney(3, objCard)
        
        If RoundEx(dblTemp, 6) < RoundEx(dblMoney, 6) Then
            MsgBox "ע��:" & vbCrLf & "   ������˿��������" & objCard.���� & "�Ŀ��˽����飡" & vbCrLf & _
                   "   ���˽��:" & Format(dblTemp, "###0.00;-###0.00;;") & vbCrLf & _
                   "   ��ǰ�˿�:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
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
         If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.�ӿ����, _
             objCard.���ѿ�, mobjDelBalance.����, mobjDelBalance.�Ա�, _
             mobjDelBalance.����, dblMoney, mCurBrushCard.str����, mCurBrushCard.str����, _
             True, True, bln����, True, Nothing, False, False, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
    
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
        '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��2-����ҵ��;3-�����˷�ҵ��4-�����˷�ҵ��
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
        strXMLExpend = "<IN><CZLX>4</CZLX></IN>"
        If gobjSquare.objSquareCard.zltransferAccountsCheck(Me, mlngModule, objCard.�ӿ����, _
            mCurBrushCard.str����, dblMoney, mobjDelBalance.ԭ����ID, strXMLExpend) = False Then Exit Function
    Else
        '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        mobjDelBalance.rsBalance.Filter = "����=3 And �����ID=" & objCard.�ӿ����
        If mobjDelBalance.rsBalance.EOF Then
            MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���㷽ʽ & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        If dblDelMoney <> 0 Then
            dblMoney = dblDelMoney
        Else
            dblMoney = Val(txt�ɿ�.Text)
        End If
        dblTemp = 0
        With mobjDelBalance.rsBalance
            Do While Not .EOF
                dblTemp = dblTemp + Val(Nvl(!��Ԥ��))
                .MoveNext
            Loop
            mobjDelBalance.rsBalance.MoveFirst
            dblTemp = RoundEx(dblTemp, 6)
        End With
    
        If dblTemp < dblMoney Then
            MsgBox "ע��:" & vbCrLf & "   ������˿��������" & objCard.���� & "�Ŀ��˽����飡" & vbCrLf & _
                   "   ���˽��:" & Format(dblTemp, "###0.00;-###0.00;;") & vbCrLf & _
                   "   ��ǰ�˿�:" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        If objCard.�Ƿ�ȫ�� And Not objCard.�Ƿ����� Then
            If dblTemp <> dblMoney Then
                MsgBox "ע��:" & vbCrLf & objCard.���� & "�����˿�ʱ������ȫ�ˣ�" & vbCrLf & _
                "  ʣ��δ��:" & Format(dblTemp, "0.00") & vbCrLf & _
                "  ��ǰ���:" & Format(dblMoney, "0.00"), vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
        
        mCurBrushCard.str���� = Nvl(mobjDelBalance.rsBalance!����)
        mCurBrushCard.str������ˮ�� = Nvl(mobjDelBalance.rsBalance!������ˮ��)
        mCurBrushCard.str����˵�� = Nvl(mobjDelBalance.rsBalance!����˵��)
        
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
        strXMLExpend = mfrmMain.GetDelXMLExpend()
        strBalanceIDs = "3|" & mobjDelBalance.ԭ����ID '& IIf(mobjDelBalance.����ID = 0, "", "," & mobjDelBalance.����ID)
        If gobjSquare.objSquareCard.zlReturnCheck(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, _
            strBalanceIDs, dblMoney, mCurBrushCard.str������ˮ��, mCurBrushCard.str����˵��, strXMLExpend) = False Then Exit Function
        
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
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.�ӿ����, _
                objCard.���ѿ�, mobjDelBalance.����, mobjDelBalance.�Ա�, _
                mobjDelBalance.����, dblMoney, mCurBrushCard.str����, mCurBrushCard.str����, _
                True, True, bln����, True, Nothing, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        End If
        
    End If
    CheckThreeSwapValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPrepayMoneyIsValied(ByVal objCard As Card, Optional ByVal intType As Integer, Optional ByVal dblMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���Ԥ�����������Ƿ�Ϸ�
    '����:�Ϸ�,����true,���򷵻�False
    '����:
    '   intTppe:0-���㷽ʽѡ��Ԥ���� 1-�����б�Ĭ��Ԥ����
    '����:���˺�
    '����:2014-07-08 18:18:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTittle As String, i As Long, str���㷽ʽ As String
    Dim int����  As Integer, dblTemp As Double
    
    On Error GoTo errHandle
    If objCard.�������� <> -99 Then CheckPrepayMoneyIsValied = True: Exit Function
    
    If intType = 0 Then
        Call txt�ɿ�_LostFocus
        strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "��", "��")
        dblMoney = Val(txt�ɿ�.Text)
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox "δ����" & strTittle & "Ԥ������!", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        strTittle = IIf(dblMoney <= 0, "��", "��")
    End If

    If strTittle = "��" Then
        Dim str����IDs As String
        If zlDatabase.PatiIdentify(Me, glngSys, mobjDelBalance.����ID, dblMoney, mlngModule, 1, , IIf(-1 * gdblԤ��������鿨 >= dblMoney, False, True), True, str����IDs, _
            (gdblԤ��������鿨 <> 0), (gdblԤ��������鿨 = 2)) = False Then Exit Function
        mobjDelBalance.����IDs = str����IDs
        CheckPrepayMoneyIsValied = True
        Exit Function
    End If
    
    dblTemp = RoundEx(GetOldBalanceMoney(1, objCard), 6)
    If dblMoney > dblTemp Then
        MsgBox "��Ԥ����ܳ����շѽ����ʣ��Ԥ���" & FormatEx(dblTemp, 6) & "����", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If intType = 0 Then
        With vsBlance
            For i = .Rows - 1 To 1 Step -1
                str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
                ' 0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                int���� = Val(.TextMatrix(i, .ColIndex("����")))
                If int���� = 1 And str���㷽ʽ <> "" Then
                    MsgBox "�Ѿ�ʹ����" & str���㷽ʽ & ",�����ٴ�ʹ��Ԥ����" & lblPayType.Caption & "!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        End With
    End If
        
    '��Ԥ����
    If gbytԤ����˷��鿨 = 0 Then CheckPrepayMoneyIsValied = True: Exit Function
    If Not zlDatabase.PatiIdentify(Me, glngSys, mobjDelBalance.����ID, dblMoney, , , , , True, , , (gbytԤ����˷��鿨 = 2)) Then Exit Function
    CheckPrepayMoneyIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckCashValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ֽ�֧����ʽ��һЩ�Ϸ�����
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer, dblMoney As Double
    Dim strTittle As String
    
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
   
    On Error GoTo errHandle
    
    intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    If objCard.�������� <> 1 Then CheckCashValied = True: Exit Function
    dblMoney = Val(txt�ɿ�.Text)
    If strTittle = "�ɿ�" Then
        Select Case gTy_Module_Para.byt�ɿ����
        Case 1, 3 '1-�ಡ�ɿ�;3�����˽ɿ��ۼ�
            If RoundEx(mCurCarge.dbl��ǰδ�� - mCurCarge.dbl��������, 2) > 0 And RoundEx(dblMoney, 6) = 0 Then
               If MsgBox("ע��:" & vbCrLf & "    �ò���δ����ɿ���,�Ƿ�����շ�? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            End If
        Case 2  '2-�շ�ʱ����Ҫ����ɿ���
            If RoundEx(mCurCarge.dbl��ǰδ�� - mCurCarge.dbl��������, 2) > 0 And RoundEx(dblMoney, 6) = 0 Then
                MsgBox "ע��:" & vbCrLf & _
                "    �ò���δ����ɿ���,���ܽ����շ�!", vbInformation + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        Case Else   ',0-�������нɿ�������ۼƿ���
            'ҽ������ɿ���:Ҫ�ɶ�δ��ʱ,�Խɿ���Ϊ������������,��Ϊ��ǿ������0�����ɿ��
            If mobjDelBalance.intInsure <> 0 And Not mInsurePara.���������շ� And _
                RoundEx(mCurCarge.dbl��ǰδ�� - mCurCarge.dbl��������, 2) > 0 And RoundEx(dblMoney, 6) = 0 Then
                '���˺�:27536 20100119
                If mInsurePara.�����ѽɿ���� = False Then
                    MsgBox "������:" & vbCrLf & vbTab & "��ҽ�����˵ķ���δȫ�����㣬��ע����ȡ���˽ɿ", vbInformation, gstrSysName
                End If
            End If
        End Select
        If RoundEx(dblMoney, 6) <> 0 Then
            If Val(txt�Ҳ�.Text) < 0 Then
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        CheckCashValied = True
        Exit Function
    End If
    '�˿�
'    If dblMoney = 0 Then
'        MsgBox "δ�����˿��", vbInformation, gstrSysName
'        Exit Function
'    End If
    If dblMoney < Abs(Val(lblδ�˽��.Caption)) And RoundEx(dblMoney, 6) <> 0 Then
        MsgBox "������˿���㣡", vbInformation, gstrSysName
        Exit Function
    End If
    CheckCashValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckChequeValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���֧Ʊ֧����ʽ��һЩ�Ϸ�����
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer, dblMoney As Double
    Dim strTittle As String
    
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
   
    On Error GoTo errHandle
    
    intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    If objCard.�������� <> 2 Or Not objCard.���㷽ʽ Like "*֧Ʊ*" Then CheckChequeValied = True: Exit Function
    
    dblMoney = Val(txt�ɿ�.Text)
    
    If strTittle = "�ɿ�" Then
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox "δ����ɿ��", vbInformation, gstrSysName
            Exit Function
        End If
        CheckChequeValied = True
        Exit Function
    End If
    '�˿�
    If RoundEx(dblMoney, 6) = 0 And Not mblnTurnFee Then
        MsgBox "δ�����˿��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckChequeValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function CheckOtherValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���֧Ʊ֧����ʽ��һЩ�Ϸ�����
    '����:���ݺϷ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:21:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer, dblMoney As Double
    Dim strTittle As String, dblTemp As Double
    
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
   
    On Error GoTo errHandle
    
    intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    If intIndex <= 0 Then Exit Function
    Set objCard = mobjPayCards(intIndex)
    
    If objCard.�ӿ���� > 0 Or objCard.���㷽ʽ Like "*֧Ʊ*" Or objCard.�������� = -99 Or objCard.�������� = 1 Then CheckOtherValied = True: Exit Function
    
    dblMoney = Val(txt�ɿ�.Text)
    

    If strTittle = "�ɿ�" Then
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox "δ����ɿ��", vbInformation, gstrSysName
            Exit Function
        End If
        If dblMoney > RoundEx(mCurCarge.dbl��ǰδ��, 2) Then
            MsgBox "ע��:" & vbCrLf & "    ����Ľɿ��������δ֧���Ľ��,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        CheckOtherValied = True
        Exit Function
    End If
    
    '�˿�
    If RoundEx(dblMoney, 6) = 0 And Not mblnTurnFee Then
        MsgBox "δ�����˿��", vbInformation, gstrSysName
        Exit Function
    End If
    If dblMoney > RoundEx(Abs(mCurCarge.dbl��ǰδ��), 2) Then
        MsgBox "ע��:" & vbCrLf & "    ������˿�������˿��˽��,���ܼ���!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    If objCard.�������� = 2 And objCard.���㷽ʽ Like "*��" Then '87532
        dblTemp = RoundEx(GetOldBalanceMoney(0, objCard), 6)
        If dblMoney > dblTemp Then
            MsgBox "ע�⣺" & vbCrLf & "   ������˿�������� " & objCard.���㷽ʽ & " �Ŀ��˽����飡" & vbCrLf & _
                   "   ���˽�" & Format(dblTemp, "###0.00;-###0.00;;") & vbCrLf & _
                   "   ��ǰ�˿" & Format(dblMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckOtherValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
        

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����������ݵ���Ч��,������Ч,����true,���򷵻�False
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-08-13 16:30:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strTittle As String, i As Long, str���㷽ʽ As String
    Dim int���� As Integer, objCard As Card
    
    On Error GoTo errHandle
    If Not CheckTextLength("�������", txt�������) Then Exit Function
    If Not CheckTextLength("ժҪ", txtժҪ) Then Exit Function
    
    If mbytFunc = EM_FUN_�˷� And mcllDelPro.Count > 0 Then
        If mfrmMain.CheckSelectItemCanDel(mobjDelBalance.CurDelNos) = False Then Exit Function
    End If
    
    '�������
    If mbytFunc = EM_FUN_���� Then
        If zlIsCheckExistErrBill(mobjDelBalance.�������) = False Then
            MsgBox "��ǰ�쳣�����ѱ������㲻�ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
        If zlCheckOtherSessionDoing(mobjDelBalance.�������) Then
            MsgBox "��ǰ�������������շѴ����н��д����㲻�ܼ�����", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    If GetCurCard(objCard) = False Then
        MsgBox "��ǰ" & lblPayType.Caption & "��ʽδѡ��,��ѡ��!", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    '93114
    If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.ԭ����ID) = False Or objCard.�ӿ���� <= 0 Then
        If CheckIsExistCashValied(objCard) = False Then Exit Function
    End If
    
    '�������ĺϷ���
    If mCurCarge.dbl��ǰδ�� <= 0 Then
        strTittle = "�˿�"
    Else
        strTittle = "�ɿ�"
    End If
    
    If Not IsNumeric(txt�ɿ�.Text) And txt�ɿ�.Text <> "" Then
        MsgBox strTittle & "��������Ч��ֵ��", vbInformation, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    End If
    If Val(txt�ɿ�.Text) < 0 Then
        MsgBox strTittle & "�������븺����", vbInformation, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
    End If
    
    If Abs(Val(txt�ɿ�.Text)) > 999999999 Then
        MsgBox "����Ľɿ������,����ܳ���-999999999��999999999!", vbOKOnly, gstrSysName
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�: Exit Function
        Exit Function
    End If
    
    If txt�������.Text <> "" Then
        If zlCommFun.ActualLen(txt�������) > 30 Then
            MsgBox "������������������30���ַ��� 15�����֣�", vbInformation, gstrSysName
            If txt�������.Enabled And txt�������.Visible Then txt�������.SetFocus
            zlControl.TxtSelAll txt�������: Exit Function
        End If
        If InStr(txt�������, "'") > 0 Then
            MsgBox "������뺬�зǷ��ַ�(������)��", vbInformation, gstrSysName
            If txt�������.Enabled And txt�������.Visible Then txt�������.SetFocus
            zlControl.TxtSelAll txt�������: Exit Function
        End If
    End If
    If txtժҪ.Text <> "" Then
        If zlCommFun.ActualLen(txtժҪ) > 50 Then
            MsgBox "ժҪ�����������50���ַ��� 25�����֣�", vbInformation, gstrSysName
            If txtժҪ.Enabled And txtժҪ.Visible Then txtժҪ.SetFocus
            zlControl.TxtSelAll txtժҪ: Exit Function
        End If
        If InStr(txtժҪ, "'") > 0 Then
            MsgBox "ժҪ���зǷ��ַ�(������)��", vbInformation, gstrSysName
            If txtժҪ.Enabled And txtժҪ.Visible Then txtժҪ.SetFocus
            zlControl.TxtSelAll txtժҪ: Exit Function
        End If
    End If
    With vsBlance
        For i = .Rows - 1 To 1 Step -1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            ' 0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            If objCard.���㷽ʽ = str���㷽ʽ And int���� <> 1 And (int���� <> 5 Or (int���� = 5 And .Cell(flexcpData, i, .ColIndex("֧�����")) < 0)) Then
                'Ԥ������Ԥ�����麯�����д˼��
                MsgBox objCard.���㷽ʽ & " �Ѿ�����,��������" & objCard.���㷽ʽ & "����" & Replace(lblPayType.Caption, " ", "") & "!", vbOKOnly + vbDefaultButton1, gstrSysName
                Exit Function
            End If
        Next
    End With
        
    If CheckInterfaceNumIsValied(objCard) = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    
    '1.һ��ͨˢ��
    If CheckOneCard(objCard) = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    
'    '2.�������׼��
'    If CheckThreeSwapValied(objCard) = False Then
'        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
'        zlControl.TxtSelAll txt�ɿ�
'        Exit Function
'    End If
    
    '3.���ѿ����
    '�˷�
    If CheckSquareDelValied(objCard) = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    '�շ�
    If CheckSquareBalanceValied(objCard) = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    
    '3.���Ԥ�����Ƿ�Ϸ�
    If CheckPrepayMoneyIsValied(objCard) = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    
    '4.�ֽ�ʽ�ļ��
    If CheckCashValied = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    '5.���֧Ʊ�������
    If CheckChequeValied = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If
    '6.�����շѷ�ʽ���
    If CheckOtherValied = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Function
    End If

    '��鵱ǰ�����Ƿ�������ִ�����,��Ҫ�ǲ���ԭ����м��
    '��ֹ��������Ա����:
    '45186
    If mobjDelBalance.����ID <> 0 Then
        gstrSQL = "" & _
        "   Select  1  From ����Ԥ����¼ A " & _
        "   Where   A.����ID=[1] and nvl(A.У�Ա�־,0)<>0 and Rownum =1 and A.��¼״̬=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.����ID)
        If rsTemp.EOF Then
            '�����Ǳ�����ִ��,������Ҫ����Ƿ�����ִ��
            gstrSQL = "Select ��¼״̬, ����Ա����,����״̬ From ������ü�¼ Where ����ID=[1] And rownum=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mobjDelBalance.����ID)
            If Not rsTemp.EOF Then
                If Val(Nvl(rsTemp!��¼״̬)) <> 1 Then
                    MsgBox "�õ����Ѿ�����������Ա����,�����ٽ����շ�!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                If Val(Nvl(rsTemp!����״̬)) <> 1 Then
                    MsgBox "�ô��շ��Ѿ��������շ�,�����ٽ����շ�!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                If Nvl(rsTemp!����Ա����) <> UserInfo.���� Then
                    MsgBox "�õ��ݲ��Ǳ����շѵ�,������ȡ��������Ա�ĵ���!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo֧����ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub ClearReMoveSquareBalance(ByVal lng�����ID As Long, Optional ByVal lng���ѿ�ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƴ�ָ�������ѿ�����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-08-12 12:10:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, varTemp As Variant
    If mcllSquareBalance Is Nothing Then Exit Sub
    j = 1
    Do While True
        If j > mcllSquareBalance.Count Then Exit Do
        varTemp = mcllSquareBalance(j)
        If Val(varTemp(0)) = lng�����ID _
            And (lng���ѿ�ID = 0 Or (lng���ѿ�ID <> 0 And Val(varTemp(1)) = lng���ѿ�ID)) Then
            mcllSquareBalance.Remove j
        Else
            j = j + 1
        End If
    Loop
    If mcllSquareBalance.Count = 0 Then Set mcllSquareBalance = Nothing
End Sub

Private Sub cmdDel_Click()
    Dim int���� As Integer, dblMoney As Double
    Dim lngCardTypeID As Long, lng���ѿ�ID As Long
    Dim objCard As Card
    Dim blnǿ������ As Boolean
    Dim str��������� As String
    
    'ɾ����صķ���
    With vsBlance
        If .Row < 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("ɾ����־"))) = 1 Then Exit Sub
        
        '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        int���� = Val(.TextMatrix(.Row, .ColIndex("����")))
        lngCardTypeID = Val(.TextMatrix(.Row, .ColIndex("�����ID")))
        lng���ѿ�ID = Val(.TextMatrix(.Row, .ColIndex("���ѿ�ID")))
        str��������� = Trim(.TextMatrix(.Row, .ColIndex("���������")))
        
        If int���� = 3 And Val(.TextMatrix(.Row, .ColIndex("֧�����"))) <> 0 Then
            '105432
            Set objCard = GetPayCard(lngCardTypeID, False, False)
            If ThreeBalanceCheck(Me, mlngModule, objCard, mcllForceDelToCash, _
                str���������, blnǿ������) = False Then Exit Sub
        End If
        
        mobjDelBalance.ԭ���� = False
        If int���� = 5 Then
            Set objCard = GetPayCard(lngCardTypeID, True, False)
            If objCard Is Nothing Then
                MsgBox "ע��:" & vbCrLf & "δ�ҵ�ָ�������ѿ�,����ɾ��!", vbInformation, gstrSysName
                Exit Sub
            End If
            If objCard.�Ƿ����� = 0 Then
                MsgBox "ע��:" & vbCrLf & "    " & objCard.���㷽ʽ & "��֧������,����ɾ��!", vbInformation, gstrSysName
                Exit Sub
            End If
            Call ClearSquareBalance(lngCardTypeID, lng���ѿ�ID) '������ѿ�����
            Call ClearReMoveSquareBalance(lngCardTypeID, lng���ѿ�ID)
        Else
            If lngCardTypeID = 0 Then
                Set objCard = GetPayCard(Trim(.TextMatrix(.Row, .ColIndex("֧����ʽ"))), False, False)
            Else
                Set objCard = GetPayCard(lngCardTypeID, False, False)
            End If
            dblMoney = Val(.Cell(flexcpData, .Row, .ColIndex("֧�����")))
            mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� + dblMoney, 6)
            mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� - dblMoney, 6)
            If .Rows <= 2 Then
                .Clear 1
                .RowData(1) = ""
                .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
            Else
                vsBlance.RemoveItem .Row
            End If
        End If
    End With
    Call Set�˷ѷ�ʽ(IIf(mCurCarge.dbl��ǰδ�� <= 0, 2, 3), , , blnǿ������)
    Call Load�˷ѷ�ʽ(blnǿ������)
    Call SetDeleteVisible
    Call SetControlProperty(True)
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdExit_Click()
    mblnOK = False
    Call ExcuteMainReshData
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim blnUnload As Boolean
   
    '���ݽ��水�˻س���
    If mblnCacheKeyReturn Then mblnCacheKeyReturn = False: Exit Sub
    '�ٴ�������
    If isValied = False Then Exit Sub
    If txt�ɿ�.Text <> "0.00" Then
        'LED��ʾ
        Call ShowLedInfor
    End If
    If Not Executeԭ���� Then Exit Sub
    '2.�������׼��
    '93114����isValied()�зŵ�����������Ϊ������˷��б���ȱʡ������������ʱ����ѡ�����������������ôˢ����Ϣ���ᱻ����
    If CheckThreeSwapValied(Nothing) = False Then
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        zlControl.TxtSelAll txt�ɿ�
        Exit Sub
    End If
    
    If ExecuteDelete(blnUnload) = False Then Exit Sub
    If blnUnload Then
        'ˢ����������Ϣ
        ExcuteMainReshData
        Unload Me
    End If
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

Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ʾ״̬
    '����:���˺�
    '����:2014-07-08 19:12:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTemp As Boolean
    
    If mbytFunc = EM_FUN_�˷� Then
        'ҽ�����н����˵�,���ҽ����,��ʾ����շ�
        cmdOK.Visible = True
        'ҽ�������˽����,�����˳�
        cmdExit.Visible = mobjDelBalance.SaveBilled = False
        Exit Sub
     End If
     If mbytFunc = EM_FUN_���� Then
        cmdExit.Caption = "�˳�(&E)"
        cmdOK.Visible = True: cmdExit.Visible = True
     End If
End Sub
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call StartAndStopԤ���
    Call cbo֧����ʽ_Click
    Call SetControlProperty
    Call Set�˷ѷ�ʽ(IIf(mCurCarge.dbl��ǰδ�� <= 0, 2, 3)): Call Load�˷ѷ�ʽ
    Call SetCtrlVisible
    mblnLoad = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Select Case KeyCode
    Case vbKeyAdd, vbKeyF4
        If gTy_Module_Para.blnʹ�üӼ��л� = False And KeyCode = vbKeyAdd Then Exit Sub
        If Me.ActiveControl Is txt�ɿ� And cbo֧����ʽ.Enabled Then
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
        If Me.ActiveControl Is txt�ɿ� And cbo֧����ʽ.Enabled Then
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
        If cmdOK.Visible And cmdOK.Enabled Then
            cmdOK.SetFocus
            cmdOK_Click
        End If
    Case vbKeyReturn
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(":'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    'ѡ������������Ƿ����˻س�����
    mblnCacheKeyReturn = (GetAsyncKeyState(VK_RETURN) And &H1) <> 0
    mstrTittle = "�����˷ѽ���"
    
    
    RestoreWinState Me, App.ProductName, mstrTittle
    Call SetWindowsSize
    Set mrsOneCard = GetOneCard
    zlControl.CboSetWidth cbo֧����ʽ.hWnd, cbo֧����ʽ.Width * 2
    mblnFirst = True: mblnLoad = True
    mblnUnLoad = False
    zlControl.PicShowFlat picTotal, -1, , taCenterAlign
    zlControl.PicShowFlat Picture1, -1, , taCenterAlign
    zlControl.PicShowFlat picPay, -1, , taCenterAlign
    Call InitFace
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    'If Me.Width < 10530 Then Me.Width = 10530
    'If Me.Height < 7035 Then Me.Height = 7035
    With picBlance
        .Width = ScaleWidth - .Left * 2
        .Height = ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    With mCurCarge
           .dbl������Ԥ�� = 0
           .dbl�˷Ѻϼ� = 0
           .dbl������� = 0
           .dbl����ҽ���˷� = 0
           .dbl���˺ϼ� = 0
           .dbl����Ӧ�� = 0
           .dbl��ǰδ�� = 0
           .dbl������� = 0
           .dbl����Ԥ�� = 0
           .dblԤ����� = 0
    End With
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
    
    Set mrsClassMoney = Nothing
    With mCurBrushCard
        .dbl�ʻ���� = 0
        .str������ˮ�� = ""
        .str����˵�� = ""
        .str���� = ""
        .str��չ��Ϣ = ""
        .str���� = ""
    End With
    Set mrsUsedCards = Nothing
    SaveWinState Me, App.ProductName, mstrTittle
End Sub

 

 
Private Sub picBlance_Resize()
    Err = 0: On Error Resume Next
    With vsBlance
        .Left = picBlance.ScaleLeft
        .Width = picBlance.ScaleWidth
        .Height = picBlance.ScaleHeight - .Top
    End With
End Sub
 
Private Sub LoadPatiInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '����:���˺�
    '����:2011-08-13 10:52:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    stbThis.Panels(2).Text = mobjDelBalance.����
    Set rsTemp = GetMoneyInfo(mobjDelBalance.����ID, 0, False, 1, False, 0, True)
    Dim dbl������� As Double
    With mCurCarge
        .dblԤ����� = 0
        .dbl������� = 0
        Do While Not rsTemp.EOF
            .dblԤ����� = .dblԤ����� + Val(Nvl(rsTemp!Ԥ�����))
            .dbl������� = .dbl������� + Val(Nvl(rsTemp!�������))
            If Nvl(rsTemp!����, 0) = 1 Then
                dbl������� = Val(Nvl(rsTemp!Ԥ�����)) - Val(Nvl(rsTemp!�������))
            End If
            rsTemp.MoveNext
        Loop
        .dbl����Ԥ�� = .dblԤ����� - .dbl�������
    End With
    If RoundEx(mCurCarge.dbl����Ԥ��, 6) = 0 And RoundEx(dbl�������, 6) = 0 Then
        stbThis.Panels(3).Visible = False
    Else
        stbThis.Panels(3).Visible = True
        stbThis.Panels(3).Text = "Ԥ��:" & Format(mCurCarge.dbl����Ԥ��, "0.00") & _
            IIf(dbl������� > 0, "(������:" & Format(dbl�������, "0.00") & ")", "")
    End If
    
    lbl�˷Ѻϼ�.Caption = Format(Abs(mCurCarge.dbl�˷Ѻϼ�), "###0.00;-###0.00;0.00;0.00;")
End Sub

Private Sub LedVoiceSpeak()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���˺�
    '����:2011-08-13 16:38:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'If mCurBrushCard.int���� <> 1 Then Exit Sub
    If gblnLED = False Then Exit Sub
    
    If mobjDelBalance.intInsure <> 0 Then Exit Sub
'    If mCurCarge.dbl�˷Ѻϼ� = 0 Then Exit Sub
'    If mCurCarge.dbl��ǰδ�� = 0 Then Exit Sub

    If mCurCarge.dbl��ǰδ�� < 0 Then
'        zl9LedVoice.Speak "#21 " & Format(-1 * lblδ�˽��.Caption, "0.00")
    Else
        zl9LedVoice.Speak "#21 " & Format(lblδ�˽��.Caption, "0.00")
    End If
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
Private Function zlGetClassMoney(ByRef lng������� As Long, ByRef rsMoney As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ,��ʼ��֧�����(�շ����,ʵ�ս��)
    '����:���سɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-06-10 17:52:18
    '����:38841
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    On Error GoTo errHandle

'    If Not mrsClassMoney Is Nothing Then
'        Set rsMoney = mrsClassMoney: zlGetClassMoney = True: Exit Function
'    End If
    If lng������� = 0 Then
        Call mfrmMain.zlGetClassMoney(rsMoney)
        zlGetClassMoney = True: Exit Function
    End If
    '��ʼ�����ݽṹ
    Set mrsClassMoney = New ADODB.Recordset
    mrsClassMoney.Fields.Append "�շ����", adVarChar, 10, adFldIsNullable
    mrsClassMoney.Fields.Append "���", adDouble, , adFldIsNullable
    mrsClassMoney.CursorLocation = adUseClient
    mrsClassMoney.LockType = adLockOptimistic
    mrsClassMoney.CursorType = adOpenStatic
    mrsClassMoney.Open
    strSQL = "" & _
    "   Select  A.�շ����,nvl(sum(ʵ�ս��) ,0) as ���   " & _
    "   From ������ü�¼ A,(Select ����ID From ����Ԥ����¼ where �������=[1] ) B " & _
    "   Where A.����ID=B.����ID " & _
    "   Group by �շ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�������)

    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            mrsClassMoney.Find "�շ����='" & Nvl(!�շ����, "��") & "'", , adSearchForward, 1
            If mrsClassMoney.EOF Then mrsClassMoney.AddNew
            mrsClassMoney!�շ���� = Nvl(!�շ����, "��")
            mrsClassMoney!��� = Val(Nvl(mrsClassMoney!���)) + Val(Nvl(!���))
            mrsClassMoney.Update
            .MoveNext
        Loop
    End With
    Set rsMoney = mrsClassMoney
    zlGetClassMoney = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt�ɿ�_Change()
    Call Show�����
    Call SetControlProperty
End Sub

Private Sub txt�ɿ�_GotFocus()
    Dim strTittle As String
    'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
    Select Case strTittle
    Case "�ɿ�"
        If gTy_Module_Para.byt�ɿ���� = 1 _
            Or gTy_Module_Para.byt�ɿ���� = 3 _
            Or gTy_Module_Para.byt�ɿ���� = 2 Then
            If Val(txt�ɿ�.Text) = 0 And Me.ActiveControl Is txt�ɿ� Then txt�ɿ�.Text = ""
        End If
    Case "�˿�"
    End Select
    Call SetControlProperty(True)
    '�Զ����ۻ��ֹ�����ʱ���ȼ�����
    If Not mbln�ѱ��� Then Call LedVoiceSpeak
    zlControl.TxtSelAll txt�ɿ�
End Sub

Private Sub ShowLedInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾLed��Ϣ
    '����:���˺�
    '����:2011-08-13 15:25:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, objCard As Card
    Dim strTittle As String
    If gblnLED = False Then Exit Sub
    If mCurCarge.dbl�˷Ѻϼ� = 0 Then Exit Sub
    
    Call GetCurCard(objCard)
    If mCurCarge.dbl��ǰδ�� <= 0 Then
        strTittle = "�˿�"
    Else
        strTittle = "�ɿ�"
    End If
    
    Select Case strTittle
    Case "�ɿ�"
        'ֻ�н��ֲ���ʾ
        If objCard.�������� = 1 Then
            zl9LedVoice.DispCharge mCurCarge.dbl��ǰδ��, Val(txt�ɿ�.Text), Val(txt�Ҳ�.Text)
        Else
            Call zl9LedVoice.DisplayBank( _
                "�ϼ�:" & lbl�˷Ѻϼ�.Caption & "Ԫ,Ӧ��:" & lblδ�˽��.Caption & "Ԫ", _
                "����:" & txt�ɿ�.Text & "Ԫ" & IIf(Val(txt�Ҳ�.Text) = 0, "", ",����:" & Val(txt�Ҳ�.Text) & "Ԫ"))
        End If
        zl9LedVoice.Speak "#22 " & Val(txt�ɿ�.Text)
        zl9LedVoice.Speak "#23 " & Val(txt�Ҳ�.Text)
        zl9LedVoice.Speak "#3"
    Case "�˿�"
        'ֻ�н��ֲ���ʾ
        If objCard.�������� = 1 Then
            zl9LedVoice.DispCharge mCurCarge.dbl��ǰδ��, -1 * Val(txt�ɿ�.Text), -1 * Val(txt�Ҳ�.Text)
        Else
            Call zl9LedVoice.DisplayBank( _
                "�ϼ�:" & lbl�˷Ѻϼ�.Caption & "Ԫ,Ӧ��:" & lblδ�˽��.Caption & "Ԫ", _
                "����:" & txt�ɿ�.Text & "Ԫ" & IIf(Val(txt�Ҳ�.Text) = 0, "", ",����:" & Val(txt�Ҳ�.Text) & "Ԫ"))
        End If
'        zl9LedVoice.Speak "#22 " & -1 * Val(txt�ɿ�.Text)
'        zl9LedVoice.Speak "#23 " & -1 * Val(txt�Ҳ�.Text)
    End Select
'    zl9LedVoice.Speak "#3"
End Sub

Private Sub LedDisplayBank(ByVal blnLedAsked As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�ѽ�����Ϣ
    '���:blnLedAsked-�Ƿ��ѱ���
    '����:���˺�
    '����:2011-12-15 13:40:46
    '����:52117
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl���ʺϼ� As Double, i As Long
    Dim strҽ�� As String, str�������� As String, str��һ��ͨ As String, str��ͨ���� As String
    Dim varPara  As Variant, str���㷽ʽ As String
    Dim strTittle As String
    If Not gblnLED Then Exit Sub
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
    With vsBlance
        '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
        For i = 1 To .Rows - 1
            'ҽ������
            If .TextMatrix(i, .ColIndex("֧����ʽ")) <> "" Then
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                Case 2 'ҽ��
                    strҽ�� = strҽ�� & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("֧�����"))), "0.00")
                Case 3 '�����ӿڽ���
                    str�������� = str�������� & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("֧�����"))), "0.00")
                Case 4   ' һ��ͨ����
                    str��һ��ͨ = str��һ��ͨ & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("֧�����"))), "0.00")
                Case Else
                    str��ͨ���� = str��ͨ���� & "||" & .TextMatrix(i, .ColIndex("֧����ʽ")) & ":" & Format(Val(.Cell(flexcpData, i, .ColIndex("֧�����"))), "0.00")
                End Select
            End If
        Next
    End With
     
    str���㷽ʽ = ""
    If strҽ�� <> "" Then str���㷽ʽ = str���㷽ʽ & "||ҽ������:||�ʻ����:" & Format(mcur�������, "0.00") & strҽ��
    If str�������� <> "" Then str���㷽ʽ = str���㷽ʽ & "||һ��ͨ����:" & str��������
    If str��һ��ͨ <> "" Then str���㷽ʽ = str���㷽ʽ & "||һ��ͨ����(��):" & str��һ��ͨ
    If str��ͨ���� <> "" Then str���㷽ʽ = str���㷽ʽ & "||��������:" & str��ͨ����
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
        If str���㷽ʽ > "" Then str���㷽ʽ = Mid(str���㷽ʽ, 2)
        zl9LedVoice.DisplayBank varPara(0), varPara(1), varPara(2), varPara(3), varPara(4), varPara(5), varPara(6), varPara(7), varPara(8), varPara(9), str���㷽ʽ
    End Select
    
    If blnLedAsked = False Then
        If strTittle = "�˿�" Then
'            zl9LedVoice.Speak "#21 " & Format(-1 * Val(lblδ�˽��.Caption), "0.00")
        Else
            zl9LedVoice.Speak "#21 " & Format(Val(lblδ�˽��.Caption), "0.00")
        End If
    End If
End Sub

Private Function Check�ɿ�() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ɿ���
    '����:����Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 10:30:38
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, intIndex As Integer
    Dim strTittle As String
    
    On Error GoTo errHandle
    intIndex = cbo֧����ʽ.ItemData(cbo֧����ʽ.ListIndex)
    If intIndex <= 0 Then Exit Function
    
    Set objCard = mobjPayCards(intIndex)
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
    
    If txt�ɿ�.Text <> "" Then
        If Abs(Val(txt�ɿ�.Text)) > 999999999 Then
            MsgBox "����Ľɿ������,����ܳ���-999999999��99999999!", vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If Val(txt�ɿ�.Text) = 0 Then
            If (objCard.�ӿ���� >= 0 Or objCard.�������� <> 1) _
                Or (objCard.�������� = 7 And objCard.�ӿ���� <= 0) Then
                '��Ҫ�ų������ӿڽ���
                MsgBox "δ����" & strTittle & "���,������" & objCard.���㷽ʽ & "֧��,����!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
        Check�ɿ� = True
        Exit Function
    End If
    If CheckCashValied = False Then Exit Function
    Check�ɿ� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    Dim objCard As Card, strTittle As String
    
    If GetCurCard(objCard) = False Then Exit Sub
    
    zlControl.TxtCheckKeyPress txt�ɿ�, KeyAscii, m���ʽ
    If KeyAscii <> 13 Then Exit Sub
    If mblnCacheKeyReturn = True Then mblnCacheKeyReturn = False
    KeyAscii = 0
    If Check�ɿ� = False Then Exit Sub
     
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
    
    
    'ֻ�Խɿ���Ϊ�շѽ�������ʱ,��������ɿ��0
    If gTy_Module_Para.byt�ɿ���� = 1 _
        Or gTy_Module_Para.byt�ɿ���� = 3 _
        Or gTy_Module_Para.byt�ɿ���� = 2 Then
        If txt�ɿ�.Text = "" Then Exit Sub
    End If
    
    If objCard.�������� <> 1 Then
        If (objCard.���㷽ʽ Like "*֧Ʊ*" Or _
            objCard.���㷽ʽ Like "*��*") And objCard.�ӿ���� <= 0 Then
            zlCommFun.PressKey vbKeyTab
            Exit Sub
        End If
        Call cmdOK_Click
        Call txt�ɿ�_GotFocus
        Exit Sub
    End If
    
    If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
    Select Case strTittle
    Case "�ɿ�"
        If txt�ɿ�.Text <> "0.00" Then
            If Val(txt�Ҳ�.Text) >= 0 Then
                 Call cmdOK_Click: Exit Sub
            End If
            MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
            txt�ɿ�.SetFocus: zlControl.TxtSelAll txt�ɿ�
            Exit Sub
        End If
    Case "�˿�"
    End Select
    Call cmdOK_Click
End Sub


Private Sub txt�ɿ�_LostFocus()
    Dim objCard As Card
    Dim dblTemp As Double
    
    If GetCurCard(objCard) = False Then
        Set objCard = New Card
    End If
    If mCurCarge.dbl��ǰδ�� <= 0 Then
        '��ǰ������С��Ԥ����ʣ��δ�˽��ʱ������Ϊ��λС��
        dblTemp = GetOldBalanceMoney(1, objCard)
        If dblTemp > Val(txt�ɿ�.Text) Then
            txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
        Else
            txt�ɿ�.Text = FormatEx(Val(txt�ɿ�.Text), 6, , , 2)
        End If
    Else
        txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
    End If
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
    zlCommFun.OpenIme True
End Sub
Private Sub txtժҪ_LostFocus()
    zlCommFun.OpenIme False
End Sub
Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cmdOK.Visible And cmdOK.Enabled Then cmdOK.SetFocus
End Sub
Private Sub txt�Ҳ�_GotFocus()
    zlControl.TxtSelAll txt�Ҳ�
End Sub

Private Function ChargeDelOver(ByVal str�˷ѽ��� As String, _
    ByVal dblԤ��� As Double, ByRef dbl��֧Ʊ�� As Double, _
    ByVal cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�շ����
    '���:blnNotCommit-�Ƿ�û�н��������ύ�����ʱ���ύ����(ԭ���Ƕ���ͨ���˽���һ���ύ)
    '����:���˺�
    '����:2011-08-15 15:50:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl�ɿ� As Double, dbl�Ҳ� As Double
    Dim cllPro As Collection, objCard As Card
    Dim strSQL As String, i As Long
     
    On Error GoTo errHandle
    
    If GetCurCard(objCard) = False Then Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If objCard.�������� = 1 Then
        dbl�ɿ� = Val(txt�ɿ�.Text)
        dbl�Ҳ� = Val(txt�Ҳ�.Text)
    End If
    
    If dbl�ɿ� = 0 Then
        dbl�ɿ� = 0: dbl�Ҳ� = 0
    End If
    
    '����֮ǰ,�ȴ�������
    'Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    '  --��������_In:
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
    '  --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����Ԥ��_In: ������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����Ԥ��_In: ������
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     ����Ԥ��_In: ������
    strSQL = strSQL & "" & 1 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & str�˷ѽ��� & "',"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "" & dblԤ��� & ","
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "" & "NULL" & ","
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & GetForceDelToCashNote(mcllForceDelToCash) & "',"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    strSQL = strSQL & "" & dbl�ɿ� & ","
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    strSQL = strSQL & "" & dbl�Ҳ� & ","
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    strSQL = strSQL & "" & mCurCarge.dbl�������� & ","
    '  ����˷�_In   Number := 0,
    '0-δ����˷�;1-�쳣����˷�;2-����˷�
    strSQL = strSQL & "2,"
    '77141,Ƚ����,2014-8-26,������ò����շ�/�˷Ѻ�,û�н�����Ϣ
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
    strSQL = strSQL & "null,"
    '  ʣ��תԤ��_In Number:=0,
    strSQL = strSQL & "0,"
    '  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
    strSQL = strSQL & "'" & Trim(cbo֧����ʽ.Text) & "',"
    '  ��Ԥ������ids_In Varchar2 := Null
    strSQL = strSQL & "'" & mobjDelBalance.����IDs & "')"
    zlAddArray cllPro, strSQL
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    mobjDelBalance.�ɿ� = dbl�ɿ�: mobjDelBalance.�Ҳ� = dbl�Ҳ�
    Set cllBillPro = New Collection
    
    ChargeDelOver = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Show�����()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ�����
    '����:���˺�
    '����:2014-07-09 18:44:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dbl��֧Ʊ�� As Double
    Dim dblʣ���� As Double, dblTemp As Double
    Dim objCard As Card, strTittle As String
    
    On Error GoTo errHandle
    
    If GetCurCard(objCard) = False Then Exit Sub
    
    strTittle = IIf(mCurCarge.dbl��ǰδ�� <= 0, "�˿�", "�ɿ�")
    
    mCurCarge.dbl�������� = 0
    
    dblMoney = IIf(strTittle = "�˿�", -1, 1) * Val(txt�ɿ�.Text)
    dblʣ���� = RoundEx(mCurCarge.dbl��ǰδ�� - dblMoney, 6)
    
    If RoundEx(mCurCarge.dbl�������, 6) = RoundEx(mCurCarge.dbl��ǰδ��, 6) Then
        mCurCarge.dbl�������� = mCurCarge.dbl��ǰδ��
    Else
        If objCard.�������� = -99 Then
            mCurCarge.dbl�������� = mCurCarge.dbl�˷Ѻϼ� - mCurCarge.dbl���˺ϼ� - RoundEx(mCurCarge.dbl��ǰδ��, 2)
        ElseIf objCard.�������� = 1 Then
            '�ֽ�
            dblTemp = IIf(dblMoney = 0, dblʣ����, mCurCarge.dbl��ǰδ��): dblʣ���� = 0
            If mobjDelBalance.intInsure > 0 Then  '����:43855
                If mInsurePara.�ֱҴ��� Then
                    dblMoney = CentMoney(CCur(dblTemp))
                Else
                    dblMoney = Format(dblTemp, "0.00")
                End If
            Else
                 dblMoney = CentMoney(CCur(dblTemp))
            End If
            mCurCarge.dbl�������� = mCurCarge.dbl�˷Ѻϼ� - mCurCarge.dbl���˺ϼ� - dblMoney
        Else
            mCurCarge.dbl�������� = mCurCarge.dbl�˷Ѻϼ� - mCurCarge.dbl���˺ϼ� - RoundEx(mCurCarge.dbl��ǰδ��, 2)
        End If
    End If
    
    mCurCarge.dbl�������� = RoundEx(mCurCarge.dbl��������, 6)
    pic���.Visible = mCurCarge.dbl�������� <> 0
    lbl����.Caption = FormatEx(mCurCarge.dbl��������, 6, , , 2)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function CheckMulitInterfaceNum() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ͬʱ�����������Ͻӿ�(��������)
    '����:�����������Ͻӿڵ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-07 15:07:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer, i As Long, int���� As Integer, str���㷽ʽ As String
    Dim varData As Variant, strErrMsg As String
    Dim objCard As Card, strTittle As String
    
    On Error GoTo errHandle
    strErrMsg = ""
    
    If GetCurCard(objCard) = False Then Exit Function
    If objCard.�������� = -99 Or objCard.�ӿ���� <= 0 Then
        CheckMulitInterfaceNum = True: Exit Function
    End If
   'ҽ����һ���ӿ�
   If mobjDelBalance.intInsure <> 0 Then intCount = intCount + 1: strErrMsg = strErrMsg & "ҽ������:" & mobjDelBalance.ҽ��������
   With vsBlance
        For i = 1 To .Rows - 1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            'rowdata:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If InStr("34", int����) > 0 Then
                If int���� = 4 Then intCount = intCount + 1
                If int���� = 3 Then '�����ӿ�
                    intCount = intCount + 1: strErrMsg = strErrMsg & vbCrLf & str���㷽ʽ & ":" & .Cell(flexcpData, i, .ColIndex("֧�����"))
                End If
            End If
        Next
    End With
    If intCount > 2 Then
        Call MsgBox("ע��:" & vbCrLf & "   ��ϵͳĿǰֻ֧���������½ӿ�,�����Ѿ��������½ӿڽ���:" & vbCrLf & strErrMsg, vbInformation + vbOKOnly, gstrSysName)
        Exit Function
    End If
    CheckMulitInterfaceNum = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ExecuteDelete(Optional ByRef blnUnload As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������
    '���:blnUnload-�Ƿ��շ���ɣ��˳��󣬽�Unload����
    '����:�˷ѳɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-10 09:53:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDel As Boolean, blnHaveMoney As Boolean
    Dim dblMoney As Double, dbl��֧Ʊ�� As Double
    Dim objCard As Card, strTittle As String, str�˷ѽ��� As String
    Dim dblʣ���� As Double, dblTemp As Double, dblԤ��� As Double
    Dim j As Long, i As Long, strCardNo As String
    Dim cllBalance As Collection
    
    On Error GoTo errHandle
    blnUnload = False
    If CheckMulitInterfaceNum = False Then Exit Function
    
    If GetCurCard(objCard) = False Then
        MsgBox lblPayType.Caption & "��ʽδѡ��!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    blnDel = mCurCarge.dbl��ǰδ�� <= 0
    mobjDelBalance.�˷ѽ��� = ""
    
    dblMoney = IIf(blnDel, -1, 1) * Val(txt�ɿ�.Text)
    
    dbl��֧Ʊ�� = 0
    dblʣ���� = mCurCarge.dbl��ǰδ�� - dblMoney - mCurCarge.dbl��������
    
    
    If objCard.�������� = -99 Then
        mobjDelBalance.�˷ѽ��� = mobjDelBalance.�˷ѽ��� & "|" & IIf(blnDel, "��Ԥ����:", "��Ԥ��:") & dblMoney
    ElseIf objCard.�������� = 1 Then
        If RoundEx(mCurCarge.dbl�������, 6) = RoundEx(mCurCarge.dbl��ǰδ��, 6) Then
            dblMoney = 0
        Else
            dblTemp = IIf(dblMoney = 0, dblʣ����, mCurCarge.dbl��ǰδ��): dblʣ���� = 0
            If mobjDelBalance.intInsure > 0 Then
                If gclsInsure.GetCapability(support�ֱҴ���, , mobjDelBalance.intInsure) Then
                    dblMoney = CentMoney(CCur(dblTemp))
                Else
                    dblMoney = Format(dblTemp, "0.00")
                End If
            Else
                dblMoney = CentMoney(CCur(dblTemp))
            End If
        End If
        
        If Val(txt�ɿ�.Text) <> 0 Then
            mobjDelBalance.�˷ѽ��� = mobjDelBalance.�˷ѽ��� & "|�ɿ�:" & IIf(blnDel, -1, 1) * Val(txt�ɿ�.Text) & ":1"
            mobjDelBalance.�˷ѽ��� = mobjDelBalance.�˷ѽ��� & "|�Ҳ�:" & IIf(blnDel, -1, 1) * Val(txt�Ҳ�.Text) & ":2"
        End If
        mobjDelBalance.�˷ѽ��� = mobjDelBalance.�˷ѽ��� & "|" & objCard.���㷽ʽ & ":" & dblMoney
        
    ElseIf objCard.���㷽ʽ Like "*֧Ʊ*" Then
        mobjDelBalance.�˷ѽ��� = mobjDelBalance.�˷ѽ��� & "|" & objCard.���㷽ʽ & ":" & dblMoney
        If blnDel = False Then
            '����:58344
            '����Ƿ�ǰ֧�����Ϊ����,�Ǹ���ʱ,��Ҫ���Ѳ���Ա(��Ҫ��ҽ������ʱ���ܴ��ڱ����ݵķ���)
            If RoundEx(dblʣ����, 2) < 0 Then
                If mstr��֧Ʊ = "" Then
                    MsgBox "�ڽ��㷽ʽ��û������Ӧ����Ľ��㷽ʽ,���ܽ�����֧Ʊ����", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                dbl��֧Ʊ�� = -1 * Val(txt�Ҳ�.Text)
                mobjDelBalance.�˷ѽ��� = mobjDelBalance.�˷ѽ��� & "|" & mstr��֧Ʊ & ":" & -1 * dbl��֧Ʊ�� & ":2"
            End If
        Else
            If RoundEx(dblʣ����, 2) > 0 Then
                MsgBox objCard.���㷽ʽ & "����ȫ��!", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
        End If
    Else
        mobjDelBalance.�˷ѽ��� = mobjDelBalance.�˷ѽ��� & "|" & objCard.���㷽ʽ & ":" & dblMoney
    End If
    Call Show�����
    
    If objCard.�������� = 1 Then
        '���ܴ���10��Ǯ
        If Abs(mCurCarge.dbl��������) > 1.5 Then
            Call MsgBox("������,�����Ƿ���ȷ!", vbInformation + vbOKOnly, gstrSysName)
            Exit Function
        End If
    End If
    
    If RoundEx(dblʣ����, 2) <> 0 Then blnHaveMoney = True
    If blnHaveMoney = False And dblMoney = 0 Then GoTo GoOver:
     
    If blnDel Then
        '���ϰ�һ��ͨ
        If ExecuteOneCardDelInterface(objCard, -1 * dblMoney, mcllDelPro) = False Then Exit Function
        '������������
        If ExecuteThreeSwapDelInterface(objCard, -1 * dblMoney, mcllDelPro) = False Then Exit Function
    Else
        '���ϰ�һ��֧ͨ��
        If ExecuteOneCardPayInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
        '��һ��֧ͨ��(��������)
        If ExecuteThreeSwapPayInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
        '�����ѿ�֧��
        If ExecuteSquarePayInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
    End If
    Call SetCtrlVisible
    
GoOver:
    If Not blnHaveMoney Then
         
         dblԤ��� = 0: str�˷ѽ��� = Get�˷ѽ���(dblMoney, dblԤ���)
        '�����ѿ���
        If ExecuteSquareDelInterface(mcllSquareBalance, mcllDelPro) = False Then Exit Function
        Set mcllSquareBalance = Nothing 'ִ�гɹ�����ռ���
        If ChargeDelOver(str�˷ѽ���, dblԤ���, dbl��֧Ʊ��, mcllDelPro) = False Then Exit Function
        mblnOK = True: ExecuteDelete = True: mblnUnloaded = True
        blnUnload = True
        Exit Function
    End If
    If objCard.�������� = 1 Then
       '�ֽ�
        ExecuteDelete = True: Exit Function
    End If
    
    Err = 0: On Error GoTo errHandle:
    With vsBlance
        If objCard.���ѿ� Then
            Call AddSquareBalance(objCard, blnDel)
        Else
            If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                .Rows = .Rows + 1
                .RowPosition(.Rows - 1) = 1
            End If
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            .RowData(1) = 0
            strCardNo = mCurBrushCard.str����
                
            If objCard.�������� = -99 Then
                .TextMatrix(1, .ColIndex("֧����ʽ")) = IIf(blnDel, "��Ԥ���", "��Ԥ���")
                .RowData(1) = 1
                .TextMatrix(1, .ColIndex("ɾ����־")) = 0  '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                .TextMatrix(1, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                .TextMatrix(1, .ColIndex("�Ƿ���֤")) = 1
            ElseIf objCard.�ӿ���� > 0 Then
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = objCard.�ӿ���� & "|" & 3 & "|" & objCard.���ƿ� & "|" & objCard.�Ƿ�ȫ�� & "|" & objCard.�Ƿ����� & "|" & objCard.����
                .RowData(1) = 3
                strCardNo = gobjSquare.objSquareCard.zlGetCardNODencode(mCurBrushCard.str����, objCard.�ӿ����, objCard.���ѿ�)
                .TextMatrix(1, .ColIndex("ɾ����־")) = 1  '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                .TextMatrix(1, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                .Cell(flexcpBackColor, 1, 0, 1, .COLS - 1) = Me.BackColor
            ElseIf objCard.�������� = 7 And objCard.�ӿ���� <= 0 Then '��һ��ͨ
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = objCard.�ӿ���� & "|" & 3 & "|" & objCard.���ƿ� & "|" & objCard.�Ƿ�ȫ�� & "|" & objCard.�Ƿ����� & "|" & objCard.����
                .TextMatrix(1, .ColIndex("ɾ����־")) = 1  '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                .TextMatrix(1, .ColIndex("����״̬")) = 1  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                .Cell(flexcpBackColor, 1, 0, 1, .COLS - 1) = Me.BackColor
                .RowData(1) = 4
            Else
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                .TextMatrix(1, .ColIndex("ɾ����־")) = 0  '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                .TextMatrix(1, .ColIndex("����״̬")) = 0  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
            End If
            .TextMatrix(1, .ColIndex("����")) = Val(.RowData(1))
            .TextMatrix(1, .ColIndex("��������")) = objCard.��������
            .TextMatrix(1, .ColIndex("�����ID")) = objCard.�ӿ����
            .TextMatrix(1, .ColIndex("���ѿ�ID")) = 0
            
            .TextMatrix(1, .ColIndex("֧�����")) = FormatEx(-1 * dblMoney, 6, , , 2)
            .Cell(flexcpData, 1, .ColIndex("֧�����")) = FormatEx(dblMoney, 6)
            .TextMatrix(1, .ColIndex("�������")) = IIf(txt�������.Visible, Trim(txt�������.Text), "")
            .TextMatrix(1, .ColIndex("��ע")) = Trim(txtժҪ.Text)
            
            If objCard.�ӿ���� > 0 Then
                .TextMatrix(1, .ColIndex("����")) = IIf(objCard.�������Ĺ��� <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("����")) = mCurBrushCard.str����
                .TextMatrix(1, .ColIndex("������ˮ��")) = mCurBrushCard.str������ˮ��
                .TextMatrix(1, .ColIndex("����˵��")) = mCurBrushCard.str����˵��
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(1, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(1, .ColIndex("�Ƿ�ת�ʼ�����")) = IIf(objCard.�Ƿ�ת�ʼ�����, 1, 0)
                .TextMatrix(1, .ColIndex("���������")) = objCard.����
            End If
            
            mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + dblMoney, 6)
            mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� - dblMoney, 6)
        End If
        
        '�Ƴ���ǰ���㷽ʽ
        If Not objCard.���ѿ� Or (objCard.���ѿ� And blnDel) Then
            Call Set�˷ѷ�ʽ(IIf(mCurCarge.dbl��ǰδ�� <= 0, 2, 3))
            Call Load�˷ѷ�ʽ
        Else
            Call SetControlProperty(True)
            txt�ɿ�.Text = ""
        End If
        
        cbo֧����ʽ.Enabled = True 'ֻʹ��ҽ�ƿ������ѿ����㣬�˷�ʱ֧����ʽ���Ǳ������˵�
        If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
        Call LedDisplayBank(False)
    End With
    Call SetDeleteVisible
    ExecuteDelete = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt�Ҳ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt�Ҳ�_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbl�Ҳ�.Caption <> "�Ҳ�" Then Exit Sub
    zlCommFun.ShowTipInfo txt�Ҳ�.hWnd, "", False
End Sub

Private Sub vsBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
    If NewRow < 0 Then Exit Sub
    Call SetDeleteVisible
End Sub
Private Sub SetDeleteVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ɾ���ؼ���visible����
    '����:���˺�
    '����:2014-07-10 11:26:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    
    With vsBlance
        If .Row < 0 Then
            blnEdit = False
        Else
             '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
           blnEdit = (Val(.TextMatrix(.Row, .ColIndex("�Ƿ�����"))) = 1 And InStr(1, "54", Val(.RowData(.Row))) <> 0) _
                Or (Val(.RowData(.Row)) = 0 And .TextMatrix(.Row, .ColIndex("֧����ʽ")) <> "") _
                Or InStr(1, "13", Val(.RowData(.Row))) > 0
           blnEdit = blnEdit And Val(.TextMatrix(.Row, .ColIndex("ɾ����־"))) <> 1    '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
        End If
    End With
    cmdDel.Visible = blnEdit
End Sub

Private Sub SetWindowsSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ô����С
    '����:���˺�
    '����:2014-07-10 11:27:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
   If OS.IsDesinMode Then Exit Sub
    '��С����ߴ�
    With gWinRect
        .MaxW = Me.Width
        .MaxH = Screen.Height * Screen.TwipsPerPixelY
        .MinH = Me.Height
        .MinW = Me.Width
    End With
    glngOld = GetWindowLong(hWnd, GWL_WNDPROC)
    Call SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SetWindowResizeWndMessage)
End Sub

Private Sub SetControlColor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����ɫ
    '����:���˺�
    '����:2014-07-10 11:32:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
    txt�ɿ�.BackColor = IIf(txt�ɿ�.Enabled, &H80000005, Me.BackColor)
    txt�Ҳ�.BackColor = Me.BackColor
    txt�������.BackColor = IIf(txt�������.Enabled, &H80000005, Me.BackColor)
    txtժҪ.BackColor = IIf(txtժҪ.Enabled, &H80000005, Me.BackColor)
End Sub
Public Function Get�˷ѽ���(ByVal dblCurDelMoney As Double, _
    ByRef dblԤ��� As Double) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�˷ѽ�������
    '���:dblCurDelMoney-��ǰ�˷ѽ��
    '����:dblԤ���-���ر���֧����Ԥ��
    '����:�շ��ý��㷽ʽ,��ʽ����:
    '       ���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
    '����:���˺�
    '����:2014-07-10 11:33:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str���㷽ʽ As String, i As Integer, int���� As Integer
    Dim str�˷ѽ��� As String, objCard As Card
    Dim dblMoney As Double, blnDel As Double
    
    
    '���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
    '�շ����
    blnDel = IIf(mCurCarge.dbl��ǰδ�� <= 0, True, False)
    str�˷ѽ��� = ""
    With vsBlance
        dblԤ��� = 0
        For i = .Rows - 1 To 1 Step -1
            str���㷽ʽ = Trim(.TextMatrix(i, .ColIndex("֧����ʽ")))
            ' 0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            
            If str���㷽ʽ <> "" Then
                Select Case int����
                Case 0 '��ͨ����
                  If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0 Then
                    str�˷ѽ��� = str�˷ѽ��� & "||" & str���㷽ʽ
                    str�˷ѽ��� = str�˷ѽ��� & "|" & Val(.Cell(flexcpData, i, .ColIndex("֧�����")))
                    str�˷ѽ��� = str�˷ѽ��� & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("�������"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("�������"))))
                    str�˷ѽ��� = str�˷ѽ��� & "|" & IIf(Trim(.TextMatrix(i, .ColIndex("��ע"))) = "", " ", Trim(.TextMatrix(i, .ColIndex("��ע"))))
                  End If
                Case 1 'Ԥ���
                     dblԤ��� = Val(.Cell(flexcpData, i, .ColIndex("֧�����")))
                End Select
            End If
        Next
        
        If GetCurCard(objCard) = False Then Exit Function
        dblMoney = dblCurDelMoney
        If RoundEx(dblMoney, 6) <> 0 And objCard.�ӿ���� <= 0 Then
            If objCard.�������� <> -99 Then
                str�˷ѽ��� = str�˷ѽ��� & "||" & objCard.���㷽ʽ
                If objCard.�������� = 1 Then
                    '�ֽ�
                    str�˷ѽ��� = str�˷ѽ��� & "|" & dblMoney
                    str�˷ѽ��� = str�˷ѽ��� & "| "
                    str�˷ѽ��� = str�˷ѽ��� & "| "
                Else
                    str�˷ѽ��� = str�˷ѽ��� & "|" & dblMoney
                    str�˷ѽ��� = str�˷ѽ��� & "|" & IIf(Trim(txt�������) = "", " ", Trim(txt�������))
                    str�˷ѽ��� = str�˷ѽ��� & "|" & IIf(Trim(txtժҪ) = "", " ", Trim(txtժҪ))
                End If
            Else
                 dblԤ��� = RoundEx(dblԤ��� + dblMoney, 6)
            End If
        End If
    End With
    If str�˷ѽ��� <> "" Then str�˷ѽ��� = Mid(str�˷ѽ���, 3)
    Get�˷ѽ��� = str�˷ѽ���
End Function

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
Private Function ExecuteOneCardPayInterface(ByVal objCard As Card, ByVal dblMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�ϰ汾)
    '���:lng�������-��������Ž��д���
    '     dblMoney-����֧�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 10:42:15
    '˵��:�ӿ��ڲ������������
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, strҽԺ���� As String
    Dim i As Long, strSQL As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim intCardType As Integer, strSwapNO As String
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�������� <> 7 Then ExecuteOneCardPayInterface = True: Exit Function

    mrsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
    If mrsOneCard.EOF Then
        MsgBox objCard.���㷽ʽ & "δ����,���ڡ������������á�����������!", vbInformation, gstrSysName
        ExecuteOneCardPayInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '����֮ǰ,�ȴ�������
    'Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    '  --��������_In:
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
    '  --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����Ԥ��_In: ������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & objCard.���㷽ʽ & "|" & dblMoney & "| | " & "',"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str���� & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str����˵�� & "')"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ����˷�_In   Number := 0,
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
    '  ʣ��תԤ��_In Number:=0
    zlAddArray cllPro, strSQL
    
    'һ��ͨ����
    blnTrans = True
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If Not mobjICCard.PaymentSwap(dblMoney, dbl���, intCardType, Val("" & mrsOneCard!ҽԺ����), mCurBrushCard.str����, mCurBrushCard.str������ˮ��, mobjDelBalance.����ID, mobjDelBalance.����ID) Then
        gcnOracle.RollbackTrans
        MsgBox objCard.���㷽ʽ & "����ʧ��!", vbOKOnly, gstrSysName
        Exit Function
    End If
    gstrSQL = "Zl_һ��ͨ����_Update(" & 0 & ",'" & objCard.���㷽ʽ & "','" & mCurBrushCard.str���� & "','" & intCardType & "','" & strSwapNO & "'," & dbl��� & "," & mobjDelBalance.������� & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    blnTrans = False
    ExecuteOneCardPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
 End Function
 Private Function ExecuteOneCardDelInterface(ByVal objCard As Card, ByVal dblDelMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��ͨ�˷ѽӿ�(�ϰ�)
    '���:cllBillPro-���浥�ݵ�SQL
    '����:���˺�
    '����:2014-07-10 10:36:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String 'ҽԺ����
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String
    Dim str���㷽ʽ As String
    Dim cllPro As Collection, blnTrans As Boolean
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�������� <> 7 Then ExecuteOneCardDelInterface = True: Exit Function

    mrsOneCard.Filter = "���㷽ʽ='" & objCard.���㷽ʽ & "'"
    If mrsOneCard.EOF Then
        MsgBox objCard.���㷽ʽ & "δ����,���ڡ������������á�����������!", vbInformation, gstrSysName
        ExecuteOneCardDelInterface = False: Exit Function
    End If
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    On Error GoTo errHandle
    If mobjDelBalance.rsBalance Is Nothing Then Exit Function
    If mobjDelBalance.rsBalance.State <> 1 Then Exit Function
    mobjDelBalance.rsBalance.Filter = "����=4"
    If mobjDelBalance.rsBalance.RecordCount = 0 Then Exit Function
    With mobjDelBalance.rsBalance
        .MoveFirst
        Do While Not .EOF
            dblMoney = dblMoney + Val(Nvl(mobjDelBalance.rsBalance!��Ԥ��))
            .MoveNext
        Loop
        .MoveFirst
    End With
    dblMoney = RoundEx(dblMoney, 6)
    If RoundEx(dblMoney, 6) = 0 Then Exit Function
    
    If dblDelMoney <> dblMoney Then
        MsgBox objCard.���㷽ʽ & " ����ȫ��!" & vbCrLf & "ԭ������:" & Format(dblMoney, "0.00") & vbCrLf & " ���˿���:" & Format(dblDelMoney, "0.00"), vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    'һ��ͨ(��):ֻ��ʹ��һ��
    With mobjDelBalance.rsBalance
        strCardNo = Nvl(!����)
        str���㷽ʽ = Nvl(!���㷽ʽ)
        
        '���㷽ʽ|������|�������|����ժҪ||..
        str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblMoney
        str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(Nvl(!�������)) = "", " ", Trim(Nvl(!�������)))
        str���㷽ʽ = str���㷽ʽ & "| "
        
        'Zl_�����˷ѽ���_Modify
        strSQL = "Zl_�����˷ѽ���_Modify("
        '  ��������_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  ����id_In     ������ü�¼.����id%Type,
         
        strSQL = strSQL & "" & mobjDelBalance.����ID & ","
        '  ����id_In     ����Ԥ����¼.����id%Type,
        strSQL = strSQL & "" & mobjDelBalance.����ID & ","
        '  ���㷽ʽ_In   Varchar2,
        strSQL = strSQL & "'" & str���㷽ʽ & "',"
        '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
        strSQL = strSQL & "NULL,"
        '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  ����_In       ����Ԥ����¼.����%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
        strSQL = strSQL & "'" & Nvl(!������ˮ��) & "',"
        '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
        strSQL = strSQL & "'" & Nvl(!����˵��) & "')"
        '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
        '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
        '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
        '  ����˷�_In   Number := 0,
        '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null
    End With
    zlAddArray cllPro, strSQL
    
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    Err = 0: On Error GoTo ErrRoll:
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "һ��ͨ�˷ѽ��׵���ʧ��,���ܼ����˷Ѳ�����", vbExclamation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    ExecuteOneCardDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Private Function ExecuteThreeSwapPayInterface(objCard As Card, ByVal dblMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�����ӿ�)
    '���:lng�������-��������Ž��д���
    '     dblMoney-����֧�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long, strSQL As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '��һ��֧ͨ��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapPayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '����֮ǰ,�ȴ�������
    'Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    '  --��������_In:
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
    '  --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����Ԥ��_In: ������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & objCard.���㷽ʽ & "|" & dblMoney & "| | " & "',"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & objCard.�ӿ���� & ","
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str���� & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str����˵�� & "')"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ����˷�_In   Number := 0,
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
    '  ʣ��תԤ��_In Number:=0
    zlAddArray cllPro, strSQL
    
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
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    str����IDs = mobjDelBalance.����ID
    str����IDs = str����IDs & IIf(mobjDelBalance.����ID <> 0, "," & mobjDelBalance.����ID, "")
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, _
         str����IDs, _
        mobjDelBalance.CurDelNos, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    mCurBrushCard.str������ˮ�� = strSwapGlideNO
    mCurBrushCard.str����˵�� = strSwapMemo
    If objCard.���ѿ� = False Then
        Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    End If
    Call zlAddThreeSwapSQLToCollection(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '��������������Ϣ
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    
    '77156,Ƚ����,2014-8-26,��ͨ����ʹ�����п��˷Ѻ󣬻����Ե�����ذ�ť���²������˷ѵ��쳣����
    mobjDelBalance.SaveBilled = True
    ExecuteThreeSwapPayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function ExecuteThreeSwapDelInterface(ByVal objCard As Card, ByVal dblDelMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:һ��֧ͨ��(�����ӿ�)
    '���:lng�������-��������Ž��д���
    '     dblMoney-����֧�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim dblMoney As Double, str���㷽ʽ   As String
    Dim strTemp As String, strXMLExpend As String, strSwapExtendInfor As String
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    Err = 0: On Error GoTo Errhand:
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� Then ExecuteThreeSwapDelInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '���㷽ʽ|������|�������|����ժҪ||..
    str���㷽ʽ = objCard.���㷽ʽ & "|" & -1 * dblDelMoney
    str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(txt�������.Text) = "", " ", Trim(txt�������.Text))
    str���㷽ʽ = str���㷽ʽ & "|" & IIf(Trim(txtժҪ.Text) = "", " ", Trim(txtժҪ.Text))

    'Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    '  --��������_In:
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
    '  --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����Ԥ��_In: ������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  -- �����_In:��������ʱ,����
    '  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷�
    '  -- ԭ����ID_IN:ԭ����ʱ,����(���ԭ����δ����ʱ,�������һ�ν���Ϊ׼)
    
    strSQL = strSQL & "" & 2 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & objCard.�ӿ���� & ","
        '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str���� & "',"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str������ˮ�� & "',"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "'" & mCurBrushCard.str����˵�� & "')"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ����˷�_In   Number := 0,
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null
    zlAddArray cllPro, strSQL
    
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    str����IDs = mobjDelBalance.����ID & IIf(mobjDelBalance.����ID <> 0, "," & mobjDelBalance.����ID, "")
    
    '81489,Ƚ����,2015-1-22,�˷Ѵ������ID
    strSwapExtendInfor = "3|" & str����IDs: strTemp = strSwapExtendInfor
    
    '93114���˷�ʱʹ��ת�ʷ�ʽ
    If CheckThreeSwapCanTransfer(objCard, mobjDelBalance.ԭ����ID) Then
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
        '                                <CZLX>��������</CZLX> //0��NULL:������ҵ��;1-�������˷�ҵ��2-����ҵ��;3-�����˷�ҵ��4-�����˷�ҵ��
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
        strXMLExpend = "<IN><CZLX>4</CZLX></IN>"
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.�ӿ����, mCurBrushCard.str����, _
            mobjDelBalance.ԭ����ID, dblDelMoney, mCurBrushCard.str������ˮ��, mCurBrushCard.str����˵��, strSwapExtendInfor, strXMLExpend) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, mCurBrushCard.str������ˮ��, mCurBrushCard.str����˵��, cllUpdate, 2)
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
        '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
        '       dblMoney-�˿���
        '       strSwapNo-������ˮ��(�ۿ�ʱ�Ľ�����ˮ��)
        '       strSwapMemo-����˵��(�ۿ�ʱ�Ľ���˵��)
        '       strSwapExtendInfor-���룬�����˷ѵĳ���ID��
        '                           ��ʽ:�շ�����1|ID1,ID2��IDn||�շ�����n|ID1,ID2��IDn
        '                           �շ�����:1-Ԥ����,2-����,3-�շ�,4-�Һ�,5-ҽ�ƿ��տ�
        '       strSwapExtendInfor-���������׵���չ��Ϣ
        '           ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n ÿ����Ŀ�в��ܰ���|�ַ�
        If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, _
            "3|" & mobjDelBalance.ԭ����ID, dblDelMoney, mCurBrushCard.str������ˮ��, mCurBrushCard.str����˵��, strSwapExtendInfor) = False Then gcnOracle.RollbackTrans: Exit Function
        'Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, strCardNO, strSwapNO, strSwapMemo, cllUpdate, 2)
    End If
    
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
    End If
    
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, , True
    Set cllBillPro = New Collection
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    
    '77156,Ƚ����,2014-8-26,��ͨ����ʹ�����п��˷Ѻ󣬻����Ե�����ذ�ť���²������˷ѵ��쳣����
    mobjDelBalance.SaveBilled = True
    ExecuteThreeSwapDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'Private Sub SumSquareBalance(cllSquareCard As Collection)
'    '�����ѿ�ID�Ϳ��Ž��з�����ͣ�����µļ���
'    Dim cllTemp1 As New Collection, cllTemp2 As New Collection
'    Dim strCards As String, strCard As String
'    Dim varCard As Variant, varTemp As Variant
'    Dim dblSumMoney As Double
'    Dim j As Integer, i As Integer
'
'    On Error GoTo errHandle:
'    If cllSquareCard Is Nothing Then Exit Sub
'    If cllSquareCard.Count = 0 Then Exit Sub
'    '���ϲ���ֱ�Ӹ�ֵ
'    For i = 1 To cllSquareCard.Count
'        cllTemp1.Add cllSquareCard(i)
'        cllTemp2.Add cllSquareCard(i)
'    Next
'
'    Set cllSquareCard = New Collection
'    For i = 1 To cllTemp1.Count
'        'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
'        varCard = cllTemp1(i): dblSumMoney = 0
'        strCard = varCard(0) & "|" & varCard(1) & "|" & varCard(3)
'        If InStr(strCards & "||", "||" & strCard & "||") = 0 Then
'            strCards = strCards & "||" & strCard
'            For j = 1 To cllTemp2.Count
'                varTemp = cllTemp2(j)
'                If strCard = varTemp(0) & "|" & varTemp(1) & "|" & varTemp(3) Then
'                    dblSumMoney = dblSumMoney + Val(varTemp(2))
'                End If
'            Next
'            cllSquareCard.Add Array(varCard(0), varCard(1), RoundEx(dblSumMoney, 6), varCard(3), varCard(4), varCard(5), varCard(6))
'        End If
'    Next
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Sub

Private Function ExecuteSquareDelInterface(ByVal cllSquareCard As Collection, _
    ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѿ�֧��
    '���:lng�������-��������Ž��д���
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '     cllSquareCard-�����˵����ѿ���(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean, dblDelMoney As Double
    Dim str����IDs As String, i As Long, varTemp As Variant
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strCardNo As String, dblMoney As Double, str���㷽ʽ  As String
    Dim objCard As Card
 
    '�����ѿ������˿�,����true
    If cllSquareCard Is Nothing Then ExecuteSquareDelInterface = True: Exit Function
    If cllSquareCard.Count = 0 Then ExecuteSquareDelInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    Err = 0: On Error GoTo Errhand:
    '�ֶ�:����,��¼����,���㷽ʽ,ժҪ,�����ID,���������,���ƿ�,���㿨���,�������,����,������ˮ��, ����˵��,�������,У�Ա�־,ҽ��,���ѿ�id
    '     �Ƿ�����,�Ƿ�ȫ��,�Ƿ�����,��Ԥ��
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�)
    If mobjDelBalance.rsBalance Is Nothing Then Exit Function
    If mobjDelBalance.rsBalance.State <> 1 Then Exit Function
    
    For i = 1 To cllSquareCard.Count
        'array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
        varTemp = cllSquareCard(i): dblDelMoney = Val(varTemp(2))
        'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByRef objCard As Card)
        If gobjSquare.objSquareCard.zlGetCard(Val(varTemp(0)), True, objCard) = False Then Exit Function
        
        mobjDelBalance.rsBalance.Filter = "����=5 And ���㿨���=" & Val(varTemp(0)) & " And ���ѿ�ID=" & Val(varTemp(1))
        If mobjDelBalance.rsBalance.RecordCount = 0 Then
            MsgBox "δ�ҵ�" & objCard.���㷽ʽ & "��ԭʼ�����¼�����ܽ����˿������", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        With mobjDelBalance.rsBalance
            .MoveFirst
            Do While Not .EOF
                dblMoney = dblMoney + Val(Nvl(mobjDelBalance.rsBalance!��Ԥ��))
                .MoveNext
            Loop
            .MoveFirst
        End With
        dblMoney = RoundEx(dblMoney, 6)
        If RoundEx(dblMoney, 6) = 0 Then
            MsgBox objCard.���㷽ʽ & " ���Ѿ�ȫ�����꣬�������ˣ�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        If RoundEx(Val(varTemp(2)), 6) > RoundEx(dblMoney, 6) Then
            MsgBox objCard.���㷽ʽ & " ���˿������ԭʼ�����" & vbCrLf & "ԭ������:" & Format(dblMoney, "0.00") & vbCrLf & " ���˿���:" & Format(Val(varTemp(2)), "0.00"), vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
        
        '�����ID|����|���ѿ�ID|���ѽ��||.
        str���㷽ʽ = str���㷽ʽ & "||" & Val(varTemp(0))
        str���㷽ʽ = str���㷽ʽ & "|" & varTemp(3)
        str���㷽ʽ = str���㷽ʽ & "|" & Val(varTemp(1))
        str���㷽ʽ = str���㷽ʽ & "|" & -1 * dblDelMoney
    Next
    If str���㷽ʽ <> "" Then str���㷽ʽ = Mid(str���㷽ʽ, 3)
    
    'Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    '  --��������_In:
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
    '  --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����Ԥ��_In: ������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  -- �����_In:��������ʱ,����
    '  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷�
    '  -- ԭ����ID_IN:ԭ����ʱ,����(���ԭ����δ����ʱ,�������һ�ν���Ϊ׼)
    strSQL = strSQL & "" & 4 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & str���㷽ʽ & "',"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL)"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ����˷�_In   Number := 0,
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null
    zlAddArray cllPro, strSQL
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    Set cllBillPro = New Collection
    
    '77156,Ƚ����,2014-8-26,��ͨ����ʹ�����п��˷Ѻ󣬻����Ե�����ذ�ť���²������˷ѵ��쳣����
    mobjDelBalance.SaveBilled = True
    ExecuteSquareDelInterface = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Executeԭ����() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ��ԭ���˹���(ֻ���������ӿ�)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-31 14:49:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objCard As Card, lng�����ID As Long, varTemp As Variant
    Dim dblMoney As Double, int���� As Integer, varTemp1 As Variant, j As Integer
    Dim strCardTypeIDs As String, cllBalance As New Collection, dblTemp As Double
    
    On Error GoTo errHandle
    With vsBlance
        For i = 1 To .Rows - 1
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            ''�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
            lng�����ID = Val(.TextMatrix(i, .ColIndex("�����ID")))
            int���� = Val(.TextMatrix(i, .ColIndex("����")))
            varTemp = .TextMatrix(i, .ColIndex("֧����ʽ"))
            If (Val(.TextMatrix(i, .ColIndex("ɾ����־"))) = 0 Or Val(.TextMatrix(i, .ColIndex("����״̬"))) = 0) Then
                Select Case int����
                Case 1 'Ԥ����
                    If Val(.TextMatrix(i, .ColIndex("�Ƿ���֤"))) = 0 Then
                        Set objCard = GetPayCard(Trim(.TextMatrix(i, .ColIndex("֧����ʽ"))), False)
                        dblMoney = RoundEx(Val(.Cell(flexcpData, i, .ColIndex("֧�����"))), 4)
                        If CheckPrepayMoneyIsValied(objCard, 1, dblMoney) = False Then Exit Function
                        .TextMatrix(i, .ColIndex("�Ƿ���֤")) = 1
                        Call SetDeleteVisible
                    End If
                Case 3 'һ��ͨ
                    '֤���������ӿ�������༭,���,��Ҫ�����ʱ,ԭ���˿�
                    Set objCard = GetPayCard(lng�����ID, False)
                    If objCard Is Nothing Then
                        MsgBox "ע��:" & vbCrLf & varTemp & " ������Ч��֧����ʽ�����ܽ����˿", vbInformation, gstrSysName
                        '�����˳�������Ϳ��������������
                        If cmdExit.Visible = False And cmdDel.Visible = False Then cmdExit.Visible = True
                        Exit Function
                    End If
                    '���Ϸ���
                    dblMoney = RoundEx(-1 * Val(.Cell(flexcpData, i, .ColIndex("֧�����"))), 4)
                    If CheckThreeSwapValied(objCard, dblMoney) = False Then Exit Function
                    '������������
                    If ExecuteThreeSwapDelInterface(objCard, dblMoney, mcllDelPro) = False Then Exit Function
                    .TextMatrix(i, .ColIndex("ɾ����־")) = 1
                    .TextMatrix(i, .ColIndex("����״̬")) = 1
                    .Cell(flexcpBackColor, i, 0, i, .COLS - 1) = Me.BackColor
                    Call SetDeleteVisible
                Case 4 'һ��ͨ(��)
                Case 5 '���ѿ�
                
                Case Else
                End Select
            End If
        Next
    End With
    
    '���ѿ������˷�
    If Not mcllSquareBalance Is Nothing Then
        For i = 1 To mcllSquareBalance.Count
            cllBalance.Add mcllSquareBalance(i)
        Next
        
        strCardTypeIDs = ""
        For i = 1 To cllBalance.Count
            varTemp = cllBalance(i)
            lng�����ID = Val(varTemp(0))
            If InStr(1, strCardTypeIDs & ",", "," & lng�����ID & ",") = 0 Then
                Set objCard = GetPayCard(lng�����ID, True, False)
                If objCard Is Nothing Then
                    MsgBox "ע��:" & vbCrLf & " ������Ч��֧����ʽ�����ܽ����˿", vbInformation, gstrSysName
                    '�����˳�������Ϳ��������������
                    If cmdExit.Visible = False And cmdDel.Visible = False Then cmdExit.Visible = True
                    Exit Function
                End If
                dblMoney = 0
                dblTemp = 0
                For j = 1 To cllBalance.Count
                    varTemp1 = cllBalance(j)
                    'array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����,ʣ��δ�˽��)
                    If lng�����ID = Val(varTemp1(0)) Then
                        If UBound(varTemp1) >= 7 Then
                            dblMoney = dblMoney + Val(varTemp1(7))
                        Else
                            dblMoney = dblMoney + Val(varTemp1(2))
                        End If
                        dblTemp = dblTemp + Val(varTemp1(2))
                    End If
                Next
                dblMoney = RoundEx(dblMoney, 6): dblTemp = RoundEx(dblTemp, 6)
                If RoundEx(dblMoney, 6) <> 0 And RoundEx(dblTemp, 6) = 0 Then
                    '�п��ܲ���ȫ����,�Խ����б��н��Ϊ׼
                    dblMoney = 0
                    For j = 1 To vsBlance.Rows - 1
                        '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                        If Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("����"))) = 5 _
                            And Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("�����ID"))) = lng�����ID Then
                            dblMoney = dblMoney + -1 * Val(vsBlance.Cell(flexcpData, j, vsBlance.ColIndex("֧�����")))
                        End If
                    Next
                    dblMoney = RoundEx(dblMoney, 6)
                    If CheckSquareDelValied(objCard, 0, dblMoney) = False Then Exit Function
                
                    dblTemp = 0
                    For j = 1 To mcllSquareBalance.Count
                        varTemp1 = mcllSquareBalance(j)
                        'array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����)
                        If lng�����ID = Val(varTemp1(0)) Then
                            dblTemp = dblTemp + Val(varTemp1(2))
                        End If
                    Next
                    dblTemp = RoundEx(dblTemp, 6)
                    If RoundEx(dblTemp, 6) <> RoundEx(dblMoney, 6) Then
                        Set objCard = GetPayCard(lng�����ID, True)
                        Call AddSquareBalance(objCard, True)
                        Call MsgBox("ע�⣺" & vbCrLf & objCard.���㷽ʽ & "֧������뵱ǰˢ����һ�£������������˿��" & vbCrLf & _
                                    "  ԭ֧����" & Format(dblMoney, "0.00") & vbCrLf & _
                                    "  ��ǰˢ����" & Format(dblTemp, "0.00"), vbInformation + vbOKOnly, gstrSysName)
                        mobjDelBalance.ԭ���� = False
                        Call StartAndStopԤ���
                        Call SetDeleteVisible
                        Call SetControlProperty(True)
                        Exit Function
                    End If
                End If
            End If
            strCardTypeIDs = strCardTypeIDs & "," & lng�����ID
        Next
      
    End If
    'If ExecuteSquareDelInterface(mcllSquareBalance, mcllDelPro) = False Then Exit Function
    Call SetDeleteVisible
    Executeԭ���� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetPayCard(ByVal strCardType As String, ByVal bln���ѿ� As Boolean, Optional bln������ As Boolean = True) As Card
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����ID
    '���:lngCardTypeID-�����ID
    '����:����Card����
    '����:���˺�
    '����:2014-07-31 15:11:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    Dim lngCardTypeID As Long
    On Error GoTo errHandle
    If Not IsNumeric(strCardType) Then
        For Each objCard In mobjPayCards
            If objCard.�ӿ���� <= 0 And objCard.���㷽ʽ = strCardType Then
                Set GetPayCard = objCard
                Exit Function
            End If
        Next
        Exit Function
    End If
    lngCardTypeID = Val(strCardType)
    For Each objCard In mobjPayCards
        If objCard.�ӿ���� = lngCardTypeID And objCard.���ѿ� = bln���ѿ� Then
            Set GetPayCard = objCard
            Exit Function
        End If
    Next
    If bln������ = False Then
        If Not gobjSquare.objSquareCard Is Nothing Then
            'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln���ѿ� As Boolean, ByRef objCard As Card)
            If gobjSquare.objSquareCard.zlGetCard(lngCardTypeID, bln���ѿ�, objCard) = False Then Exit Function
            Set GetPayCard = objCard
        End If
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckSquareDelValied(ByVal objCard As Card, Optional ByVal lng���ѿ�ID As Long, Optional dblDelMoney As Double) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѿ��˷Ѽ��
    '���:objCard-������
    '     dblDelMoney-�˿���
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:00:34
    '˵��:ͬ����֤�˽ӿں�ˢ����Ʒ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblTemp As Double
    Dim cllSquareBalance As Collection, cllBalance As Collection
    Dim dblToTal As Double, dblBrushMoney As Double
    Dim varData As Variant, varTemp As Variant, i As Integer, j As Integer
    Dim strBalances As String, dblRestMoney As Double
    Dim lng���ѿ� As Long, str���� As String
    
    On Error GoTo errHandle
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� = False Then CheckSquareDelValied = True: Exit Function
    '�˷�
    If Not (mCurCarge.dbl��ǰδ�� <= 0 Or dblDelMoney <> 0) Then CheckSquareDelValied = True: Exit Function
    If dblDelMoney = 0 Then
        If Val(txt�ɿ�.Text) = 0 Then
            MsgBox "δ�����˷ѽ����飡", vbInformation + vbOKOnly, gstrSysName
             Exit Function
        Else
            dblDelMoney = Val(txt�ɿ�.Text)
        End If
    End If
     
    '�˿���
    If mobjDelBalance.rsBalance Is Nothing Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If mobjDelBalance.rsBalance.State <> 1 Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
    If lng���ѿ�ID <> 0 Then
        mobjDelBalance.rsBalance.Filter = "����=5 And ���㿨���=" & objCard.�ӿ���� & " And ���ѿ�ID=" & lng���ѿ�ID
    Else
        mobjDelBalance.rsBalance.Filter = "����=5 And ���㿨���=" & objCard.�ӿ����
    End If
    
    If mobjDelBalance.rsBalance.EOF Then
        MsgBox "ע��:" & vbCrLf & "  δ�ҵ�ԭʼ�Ľ����¼,����ʹ��" & objCard.���� & "�����˿�!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    Set cllSquareBalance = New Collection
    Set cllBalance = New Collection
    dblTemp = dblDelMoney: dblToTal = 0
    With mobjDelBalance.rsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblToTal = dblToTal + Val(Nvl(!��Ԥ��))
            
            lng���ѿ� = Val(Nvl(!���ѿ�ID)): str���� = Nvl(!����)
            If InStr(strBalances & ",", "," & objCard.�ӿ���� & "|" & lng���ѿ� & "|" & str���� & ",") = 0 Then
                '���ӿ���š����ѿ�ID��������ʣ��δ�˽��
                dblRestMoney = 0
                j = .AbsolutePosition
                .MoveFirst
                Do While Not .EOF
                    If Val(Nvl(!���㿨���)) = objCard.�ӿ���� _
                        And Val(Nvl(!���ѿ�ID)) = lng���ѿ� And Nvl(!����) = str���� Then
                        dblRestMoney = dblRestMoney + Val(Nvl(!��Ԥ��))
                    End If
                    .MoveNext
                Loop
                .Move j - 1, adBookmarkFirst
                
                'ʣ��δ�˽��
                dblRestMoney = RoundEx(dblRestMoney, 6)
                '��ˢ�����
                dblBrushMoney = GetSquareBrushMoney(objCard.�ӿ����, lng���ѿ�, str����)
                
                If dblRestMoney <> 0 Then
                    'array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����,ʣ��δ�˽��)
                    cllSquareBalance.Add Array(objCard.�ӿ����, lng���ѿ�, dblBrushMoney, str����, "", "", 0, dblRestMoney)
                     
                    If dblTemp > dblRestMoney And dblTemp <> 0 Then
                        cllBalance.Add Array(objCard.�ӿ����, lng���ѿ�, dblRestMoney, str����, "", "", 0)
                        dblTemp = dblTemp - dblRestMoney
                    ElseIf dblTemp <> 0 Then
                        cllBalance.Add Array(objCard.�ӿ����, lng���ѿ�, dblTemp, str����, "", "", 0)
                        dblTemp = 0
                    End If
                End If
                dblTemp = RoundEx(dblTemp, 6)
                strBalances = strBalances & "," & objCard.�ӿ���� & "|" & lng���ѿ� & "|" & str����
            End If
            .MoveNext
        Loop
    End With
    dblToTal = RoundEx(dblToTal, 6)
    
    If RoundEx(dblToTal, 6) < RoundEx(dblDelMoney, 6) Then
        MsgBox "ע��:" & vbCrLf & "   ������˿��������" & objCard.���㷽ʽ & "�Ŀ��˽��,����!" & vbCrLf & _
               "   ���˽��:" & Format(dblToTal, "###0.00;-###0.00;;") & vbCrLf & _
               "   ��ǰ�˿�:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If RoundEx(dblToTal, 6) <> RoundEx(dblDelMoney, 6) Then
        If objCard.�Ƿ�ȫ�� And Not objCard.�Ƿ����� Then
            MsgBox "ע��:" & vbCrLf & "   " & objCard.���㷽ʽ & "��֧������,����ȫ��,����!" & vbCrLf & _
                   "   ���˽��:" & Format(dblToTal, "###0.00;-###0.00;;") & vbCrLf & _
                   "   ��ǰ�˿�:" & Format(dblDelMoney, "###0.00;-###0.00;;"), vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    If gbln���ѿ��˷��鿨 Then
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
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, objCard.�ӿ����, _
            objCard.���ѿ�, mobjDelBalance.����, mobjDelBalance.�Ա�, _
            mobjDelBalance.����, dblDelMoney, "", "", _
            True, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        Set cllBalance = cllSquareBalance
    End If
    
    If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
    '�����ѿ��˷ѽ�����Ϣ�������Ƴ���ǰ�����ļ�¼
    j = 1
    Do While True
        If j > mcllSquareBalance.Count Then Exit Do
        varTemp = mcllSquareBalance(j)
        'array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����)
        If objCard.�ӿ���� = Val(varTemp(0)) Then 'And varData(3) = varTemp(3)
            mcllSquareBalance.Remove j
        Else
           j = j + 1
        End If
    Loop
    
    '��ˢ����֤��ĵ�ǰ�����ļ�¼��ӵ����ѿ��˷ѽ�����Ϣ������
    dblTemp = 0
    For i = 1 To cllBalance.Count
        varData = cllBalance(i)
        dblTemp = Val(varData(2)) + dblTemp
        mcllSquareBalance.Add varData
    Next
    CheckSquareDelValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSquareBrushMoney(ByVal lngCardTypeID As Long, ByVal lng���ѿ�ID As Long, ByVal strCardNo As String) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ѿ���ˢ�����
    '���:lngCardTypeId-���ѿ��ӿڱ��
    '     lng���ѿ�ID-���ѿ�ID
    '     strCardNo-����
    '����:
    '����:����ˢ�����
    '����:���˺�
    '����:2014-08-12 11:51:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, varTemp As Variant
    Dim dblMoney As Double
    If mobjDelBalance.ԭ���� Then Exit Function
    If mcllSquareBalance Is Nothing Then Exit Function
    dblMoney = 0
    'array(�����ID,���ѿ�ID,ˢ�����, ����,����,�������,�Ƿ�����)
    For j = 1 To mcllSquareBalance.Count
        varTemp = mcllSquareBalance(j)
        If Val(varTemp(0)) = lngCardTypeID And _
           ((lng���ѿ�ID = varTemp(1) And lng���ѿ�ID <> 0) _
             Or varTemp(3) = strCardNo) Then
             
            dblMoney = dblMoney + Val(varTemp(2))
        End If
    Next
    GetSquareBrushMoney = RoundEx(dblMoney, 6)
End Function

Private Sub ClearSquareBalance(ByVal lngCardTypeID As Long, Optional ByVal lng���ѿ�ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ѿ�����
    '����:���˺�
    '����:2014-08-12 10:39:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, j As Long
    With vsBlance
        j = 1
        Do While j <= .Rows - 1
            If Val(.TextMatrix(j, .ColIndex("����"))) = 5 _
                And Val(.TextMatrix(j, .ColIndex("ɾ����־"))) = 0 _
                And Val(.TextMatrix(j, .ColIndex("�����ID"))) = lngCardTypeID _
                And (lng���ѿ�ID = 0 Or (lng���ѿ�ID <> 0 And Val(.TextMatrix(j, .ColIndex("���ѿ�ID"))) = lng���ѿ�ID)) Then
                dblMoney = Val(.Cell(flexcpData, j, .ColIndex("֧�����")))
                
                mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� - dblMoney, 6)
                mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� + dblMoney, 6)
                If .Rows <= 2 Then
                    .Rows = 2
                   .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
                   .Cell(flexcpText, 1, 0, 1, .COLS - 1) = ""
                   .RowData(1) = ""
                   j = 2
                Else
                    .RemoveItem j
                End If
            Else
                j = j + 1
            End If
        Loop
    End With
End Sub

Private Function CheckIsExistCashValied(objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ƿ�������ֵ��������㼰���ѿ�
    '���:
    '����:
    '����:�������������ݺϷ���,����True,���򷵻�False
    '����:���˺�
    '����:2014-08-12 18:18:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsBalance As ADODB.Recordset
    Dim lngCardTypeID As Long, strCardTypeIDs As String
    Dim j As Long, blnFind As Boolean, bln���ѿ� As Boolean
    Dim int���� As Integer, lngID As Long, dblTemp As Double
    Dim blnǿ������ As Boolean
    
    On Error GoTo errHandle
    If mblnSingleBalance Then CheckIsExistCashValied = True: Exit Function '���ü��
    If mCurCarge.dbl��ǰδ�� >= 0 Then CheckIsExistCashValied = True: Exit Function '86915
    Set rsBalance = mobjDelBalance.rsBalance
    If rsBalance Is Nothing Then CheckIsExistCashValied = True: Exit Function
    If rsBalance.State <> 1 Then CheckIsExistCashValied = True: Exit Function
    
    rsBalance.Filter = "(����=3 And �Ƿ�����=0) Or (����=5 And �Ƿ�����=0)"
    If rsBalance.RecordCount = 0 Then CheckIsExistCashValied = True: Exit Function
    
    With rsBalance
        If .RecordCount <> 0 Then .MoveFirst
        strCardTypeIDs = ""
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(!�����ID)): bln���ѿ� = False
            If lngCardTypeID = 0 Then lngCardTypeID = Val(Nvl(!���㿨���)): bln���ѿ� = True
            
            blnǿ������ = False
            If lngCardTypeID > 0 Then
                For j = 1 To mcllForceDelToCash.Count
                    If mcllForceDelToCash(j)(1) = Nvl(!���������) Then blnǿ������ = True: Exit For
                Next
            End If
            
            
            If blnǿ������ = False And lngCardTypeID > 0 And InStr(1, strCardTypeIDs & "||", "||" & lngCardTypeID & "," & IIf(bln���ѿ�, 1, 0) & "||") = 0 Then
                '�����Ƿ��ڽ�����Ϣ���Ƿ����
                blnFind = False
                For j = 1 To vsBlance.Rows - 1
                    lngID = Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("�����ID")))
                     '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                    int���� = Val(vsBlance.TextMatrix(j, vsBlance.ColIndex("����")))
                    If bln���ѿ� Then
                        If lngID = lngCardTypeID And int���� = 5 Then
                            blnFind = True: Exit For
                        End If
                    Else
                        If lngID = lngCardTypeID And int���� = 3 Then
                            blnFind = True: Exit For
                        End If
                    End If
                Next
                If Not objCard Is Nothing Then
                    If objCard.�ӿ���� = lngCardTypeID And objCard.���ѿ� = bln���ѿ� Then blnFind = True
                End If
                
                If blnFind = False Then
                    j = .AbsolutePosition
                    dblTemp = 0
                    '����Ƿ������꣬������ֱ������(���ܵ�һ���˷����˹�)
                    Do While Not .EOF
                        If bln���ѿ� Then
                            If Val(Nvl(!����)) = 5 And Val(Nvl(!���㿨���)) = lngCardTypeID Then
                                dblTemp = dblTemp + Val(Nvl(!��Ԥ��))
                            End If
                        Else
                            If Val(Nvl(!����)) = 3 And Val(Nvl(!�����ID)) = lngCardTypeID Then
                                dblTemp = dblTemp + Val(Nvl(!��Ԥ��))
                            End If
                        End If
                        .MoveNext
                    Loop
                    dblTemp = RoundEx(dblTemp, 6)
                    .Move j - 1, adBookmarkFirst
                    If dblTemp <> 0 Then
                        MsgBox Nvl(rsBalance!���㷽ʽ) & " �������֣�����ȫ�ˣ�", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                End If
                strCardTypeIDs = strCardTypeIDs & "||" & lngCardTypeID & "," & IIf(bln���ѿ�, 1, 0)
            End If
            .MoveNext
        Loop
    End With
    CheckIsExistCashValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Private Sub AddSquareBalance(ByVal objCard As Card, ByVal blnDel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ѿ�֧����ʽ�����㷽ʽ�б�
    '����:���˺�
    '����:2014-08-12 18:18:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllBalance As New Collection
    Dim j As Integer, dblMoney As Double, strCardNo As String
    
    With vsBlance
      '�����ԭʼ�����ѿ�����,�������˷�
        Call ClearSquareBalance(objCard.�ӿ����)
        If blnDel Then
            If mcllSquareBalance Is Nothing Then Set mcllSquareBalance = New Collection
            Set cllBalance = mcllSquareBalance
        Else
            If mcllCurSquareBalance Is Nothing Then Set mcllCurSquareBalance = New Collection
            Set cllBalance = mcllCurSquareBalance
        End If
        
        For j = 1 To cllBalance.Count
            If objCard.�ӿ���� = Val(cllBalance(j)(0)) Then
                If Not blnDel Then
                    If mcllSquareChargeBalance Is Nothing Then Set mcllSquareChargeBalance = New Collection
                    mcllSquareChargeBalance.Add cllBalance(j)
                End If
                '��ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
                If Not (.Rows = 2 And Trim(.TextMatrix(1, .ColIndex("֧����ʽ"))) = "") Then
                    .Rows = .Rows + 1
                    .RowPosition(.Rows - 1) = 1
                End If
                
                '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                dblMoney = Val(IIf(blnDel, -1, 1) * cllBalance(j)(2))
                .RowData(1) = 5
                .TextMatrix(1, .ColIndex("����")) = 5
                .TextMatrix(1, .ColIndex("��������")) = objCard.��������
                .TextMatrix(1, .ColIndex("ɾ����־")) = IIf(blnDel, 0, 1) '�Ƿ�����༭:1-��ֹ�༭;0-����ֹ�༭
                .TextMatrix(1, .ColIndex("����״̬")) = IIf(blnDel, 0, 1)  '�Ƿ��ѽ���:1-�ѽ���;0-δ����
                .TextMatrix(1, .ColIndex("�����ID")) = objCard.�ӿ����
                .TextMatrix(1, .ColIndex("���ѿ�ID")) = Val(cllBalance(j)(1))
                .TextMatrix(1, .ColIndex("֧����ʽ")) = objCard.���㷽ʽ
                 ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                .Cell(flexcpData, 1, .ColIndex("֧����ʽ")) = Val(cllBalance(j)(0)) & "|" & 1 & "|" & IIf(objCard.���ƿ�, 1, 0) & _
                                                            "|" & IIf(objCard.�Ƿ�ȫ��, 1, 0) & "|" & IIf(objCard.�Ƿ�����, 1, 0) & "|" & objCard.����
                strCardNo = Trim(cllBalance(j)(3))
                .TextMatrix(1, .ColIndex("����")) = IIf(objCard.�������Ĺ��� <> "", String(Len(strCardNo), "*"), strCardNo)
                .Cell(flexcpData, 1, .ColIndex("����")) = strCardNo
                .TextMatrix(1, .ColIndex("֧�����")) = Format(-1 * dblMoney, "0.00")
                .Cell(flexcpData, 1, .ColIndex("֧�����")) = Format(dblMoney, "0.00")
                .TextMatrix(1, .ColIndex("�������")) = ""
                .TextMatrix(1, .ColIndex("��ע")) = ""
                .TextMatrix(1, .ColIndex("�Ƿ�����")) = IIf(objCard.�Ƿ�����, 1, 0)
                .TextMatrix(1, .ColIndex("�Ƿ�ȫ��")) = IIf(objCard.�Ƿ�ȫ��, 1, 0)
                .TextMatrix(1, .ColIndex("�Ƿ�ת�ʼ�����")) = IIf(objCard.�Ƿ�ת�ʼ�����, 1, 0)
                .TextMatrix(1, .ColIndex("���������")) = objCard.����
                
                mCurCarge.dbl���˺ϼ� = RoundEx(mCurCarge.dbl���˺ϼ� + dblMoney, 6)
                mCurCarge.dbl��ǰδ�� = RoundEx(mCurCarge.dbl��ǰδ�� - dblMoney, 6)
            End If
        Next
    End With
End Sub

Private Function CheckSquareBalanceValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѿ����㽻�׼��
    '���:objCard-������
    '����:���׺Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2014-07-08 18:00:34
    '˵��:ͬ����֤�˽ӿں�ˢ����Ʒ��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dblMoney As Double, dblTemp As Double
    Dim rsMoney As ADODB.Recordset, strXMLExpend As String
    Dim strTittle As String, dbl�ʻ���� As Double
    Dim strBrushCard As TY_BrushCard, cllSquareBalance As Collection
    Dim strExpand As String, bln���� As Boolean
    
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� = False Then CheckSquareBalanceValied = True: Exit Function
    
    On Error GoTo errHandle
    If mCurCarge.dbl��ǰδ�� <= 0 Then CheckSquareBalanceValied = True: Exit Function
    
    If Val(txt�ɿ�) = 0 Then
        MsgBox strTittle & "���δ����,����!", vbInformation + vbOKOnly, gstrSysName
         Exit Function
    End If
    If Abs(Val(txt�ɿ�.Text)) > Format(Abs(mCurCarge.dbl��ǰδ��), "0.00") And Val(txt�ɿ�.Text) <> 0 Then
        MsgBox strTittle & "���ܴ��ڱ���δ�����:" & Format(mCurCarge.dbl��ǰδ��, "0.00") & " ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ȼ���Ӧ�Ľӿ�
    If zlGetClassMoney(0, rsMoney) = False Then Exit Function
    
     '�������ѿ���ˢ����Ϣ
    Set cllSquareBalance = mcllSquareChargeBalance
    Set mcllCurSquareBalance = Nothing
    
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
    '       varSquareBalance- Collection����,���ص�ǰˢ������(array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����))
     
    dblMoney = Val(txt�ɿ�.Text)
    If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, rsMoney, _
            objCard.�ӿ����, objCard.���ѿ�, _
            mobjDelBalance.����, mobjDelBalance.�Ա�, mobjDelBalance.����, dblMoney, _
            mCurBrushCard.str����, mCurBrushCard.str����, _
            False, True, False, False, cllSquareBalance, False, False, "<IN><CZLX>0</CZLX></IN>") = False Then Exit Function
        Set mcllCurSquareBalance = cllSquareBalance
        '����ǰ,һЩ���ݼ��
        'zlPaymentCheck(frmMain As Object, ByVal lngModule As Long, _
        ByVal strCardTypeID As Long, ByVal strCardNo As String, _
        ByVal dblMoney As Double, ByVal strNOs As String, _
        Optional ByVal strXMLExpend As String
        'mobjDelBalance.strNOs:��������ʱ,û�����ʱ,����Ϊ��.
        If gobjSquare.objSquareCard.zlPaymentCheck(Me, mlngModule, objCard.�ӿ����, _
            objCard.���ѿ�, mCurBrushCard.str����, dblMoney, mobjDelBalance.CurDelNos, strXMLExpend) = False Then Exit Function
        '    zlGetAccountMoney(ByVal frmMain As Object, ByVal lngModule As Long, _
        '    ByVal strCardTypeID As Long, _
        '    ByVal strCardNo As String, strExpand As String, dblMoney As Double
        '���:frmMain-���õ�������
        '        lngModule-ģ���
        '        strCardNo-����
        '        strExpand-Ԥ����Ϊ��,�Ժ���չ
        '����:dblMoney-�����ʻ����
        If gobjSquare.objSquareCard.zlGetAccountMoney(Me, mlngModule, objCard.�ӿ����, _
              mCurBrushCard.str����, strExpand, dbl�ʻ����, objCard.���ѿ�) = False Then Exit Function
    
        stbThis.Panels(4).Text = Format(dbl�ʻ����, "0.00")
        stbThis.Panels(4).ToolTipText = objCard.���㷽ʽ & "���ʻ����:" & Format(dbl�ʻ����, "0.00")
        mCurBrushCard.dbl�ʻ���� = RoundEx(dbl�ʻ����, 2)
        If RoundEx(dbl�ʻ����, 6) <> 0 And dbl�ʻ���� < dblMoney Then
            MsgBox objCard.���㷽ʽ & "���ʻ�����!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        '�Ѿ�������֧�����
        If RoundEx(dblMoney, 6) <> Val(txt�ɿ�.Text) Then
            txt�ɿ�.Text = FormatEx(dblMoney, 6, , , 2)
        End If
        CheckSquareBalanceValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteSquarePayInterface(objCard As Card, ByVal dblMoney As Double, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ѿ�֧��
    '���:lng�������-��������Ž��д���
    '     dblMoney-����֧�����
    '     cllBillPro-���ݹ���(ִ��������,�Ա�����´νӿ�ʱ�ظ�ִ��)
    '����:ִ�гɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2014-07-09 18:14:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim cllPro As Collection, blnTrans As Boolean
    Dim str����IDs As String, i As Long, strSQL As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim str���ѿ�����  As String, j As Long
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection
    
    '�����ѿ�֧��,ֱ�ӷ���
    If objCard.�ӿ���� <= 0 Or objCard.���ѿ� = False Then ExecuteSquarePayInterface = True: Exit Function
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    str���ѿ����� = ""  '�����ID|����|���ѿ�ID|���ѽ��||....
    If mcllCurSquareBalance Is Nothing Then Exit Function
    If mcllCurSquareBalance.Count = 0 Then Exit Function
    For j = 1 To mcllCurSquareBalance.Count
        ' array(�����ID,���ѿ�ID,ˢ�����,����,����,�������,�Ƿ�����)
        str���ѿ����� = str���ѿ����� & "||" & Val(mcllCurSquareBalance(j)(0))
        str���ѿ����� = str���ѿ����� & "|" & mcllCurSquareBalance(j)(3)
        str���ѿ����� = str���ѿ����� & "|" & Val(mcllCurSquareBalance(j)(1))
        str���ѿ����� = str���ѿ����� & "|" & Val(mcllCurSquareBalance(j)(2))
    Next
    If str���ѿ����� <> "" Then str���ѿ����� = Mid(str���ѿ�����, 3)
    
    '����֮ǰ,�ȴ�������
    'Zl_�����˷ѽ���_Modify
    strSQL = "Zl_�����˷ѽ���_Modify("
    '  ��������_In   Number,
    '  --��������_In:
    '  --   0-ԭ����
    '  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
    '  --   1-��ͨ�˷ѷ�ʽ:
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
    '  --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
    '  --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
    '  --   2.�������˷ѽ���:
    '  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
    '  --     ����Ԥ��_In: ������
    '  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
    '  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    '  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    '  --   4-���ѿ�����:
    '  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
    '  --     ����Ԥ��_In: ������
    '  --     ����֧Ʊ��_In:������
    strSQL = strSQL & "" & 4 & ","
    '  ����id_In     ������ü�¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & mobjDelBalance.����ID & ","
    '  ���㷽ʽ_In   Varchar2,
    strSQL = strSQL & "'" & str���ѿ����� & "',"
    '  ��Ԥ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  �����id_In   ����Ԥ����¼.�����id%Type := Null,
    strSQL = strSQL & "" & objCard.�ӿ���� & ","
    '  ����_In       ����Ԥ����¼.����%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
    strSQL = strSQL & "NULL,"
    '  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
    strSQL = strSQL & "NULL)"
    '  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
    '  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
    '  �����_In   ������ü�¼.ʵ�ս��%Type := Null,
    '  ����˷�_In   Number := 0,
    '  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
    '  ʣ��תԤ��_In Number:=0
    zlAddArray cllPro, strSQL
    
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
    '        strBalanceIDs-����ID,����ö��ŷ���
    '        strPrepayNos-��Ԥ��ʱ��Ч. Ԥ�����ݺ�,����ö��ŷ���
    '       strCardNo-����
    '       dblMoney-֧�����
    '����:strSwapGlideNO-������ˮ��
    '       strSwapMemo-����˵��
    '       strSwapExtendInfor-������չ��Ϣ: ��ʽΪ:��Ŀ����1|��Ŀ����2||��||��Ŀ����n|��Ŀ����n
    '����:�ۿ�ɹ�,����true,���򷵻�Flase
    '˵��:
    '   ��������Ҫ�ۿ�ĵط����øýӿ�,Ŀǰ�滮��:�շ��ң��Һ���;������ѯ��;ҽ������վ��ҩ���ȡ�
    '   һ����˵���ɹ��ۿ�󣬶�Ӧ�ô�ӡ��صĽ���Ʊ�ݣ����Է��ڴ˽ӿڽ��д���.
    '   �ڿۿ�ɹ��󣬷��ؽ�����ˮ�ź���ر�ע˵���������������������Ϣ�����Է��ڽ���˵�����Ա��˷�.
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    str����IDs = mobjDelBalance.����ID
    str����IDs = str����IDs & IIf(mobjDelBalance.����ID <> 0, "," & mobjDelBalance.����ID, "")
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, Me.Caption, True
    
    If gobjSquare.objSquareCard.zlPaymentMoney(Me, mlngModule, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, _
         str����IDs, _
        mobjDelBalance.CurDelNos, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    mCurBrushCard.str������ˮ�� = strSwapGlideNO
    mCurBrushCard.str����˵�� = strSwapMemo
    If objCard.���ѿ� = False Then
        Call zlAddUpdateSwapSQL(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    End If
    '��չ������Ϣ
    Call zlAddThreeSwapSQLToCollection(False, str����IDs, objCard.�ӿ����, objCard.���ѿ�, mCurBrushCard.str����, strSwapExtendInfor, cllThreeSwap)
    zlExecuteProcedureArrAy cllUpdate, Me.Caption, True, True
    gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    '��������������Ϣ
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    blnTrans = False
    
    '77156,Ƚ����,2014-8-26,��ͨ����ʹ�����п��˷Ѻ󣬻����Ե�����ذ�ť���²������˷ѵ��쳣����
    mobjDelBalance.SaveBilled = True
    ExecuteSquarePayInterface = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Load�˷ѷ�ʽ(Optional ByVal blnǿ������ As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����˷ѷ�ʽ
    '����:���˺�
    '����:2014-09-01 16:05:44
    '˵��:
    '   ȱʡ�˿ʽ����:
    '       ���տ����ʱ��ҽ������Ϊȱʡ���˿ʽ , ������ڶ��, �����¹���ȱʡ:
    '       1)�����ʻ�:���������ʻ���,ȱʡ�������ʻ�,���ڶ�������ʻ�֧��,��ȱʡΪ��һ�������ʻ�.
    '       2)�տ���㷽ʽ��ֻ����һ�ַ�ҽ�����㷽ʽ��, ��ȱʡΪ�ý��㷽ʽ
    '       3)�տ���㷽ʽ�д���ȱʡ�Ľ��㷽ʽ, ����ȱʡ��Ϊ׼
    '       4)���ֽ�Ϊ׼
    '   blnǿ������=true:ȱʡΪ�ֽ�
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card, blnChargeUsed As Boolean
    Dim i As Long, str���㷽ʽ As String, strTemp As String
    Dim blnSetedIndex As Boolean
      
    mblnNotClick = True
    mlngPre֧����ʽ = 0
    
    Call StartAndStopԤ���

    With cbo֧����ʽ
        .Clear
        For i = 1 To mobjPayCards.Count
            Set objCard = mobjPayCards(i)
            If objCard.֧������ = True And InStr(str���㷽ʽ & "|", "|" & objCard.���㷽ʽ & "|") = 0 Then
                '�����˻���֧����ʽ��ʾΪҽ�ƿ����ƣ�������ʾ���㷽ʽ
                If objCard.�ӿ���� > 0 And Not objCard.���ѿ� Then
                    .AddItem objCard.����: str���㷽ʽ = str���㷽ʽ & "|" & objCard.����
                Else
                    .AddItem objCard.���㷽ʽ: str���㷽ʽ = str���㷽ʽ & "|" & objCard.���㷽ʽ
                End If
                .ItemData(.NewIndex) = i
            End If
        Next
        '����ȱʡֵ
        For i = .ListCount - 1 To 0 Step -1
            Set objCard = mobjPayCards(.ItemData(i))
            mobjDelBalance.rsBalance.Filter = "��������=2"
            strTemp = ""
            If mobjDelBalance.rsBalance.RecordCount > 0 Then
                mobjDelBalance.rsBalance.MoveFirst
                strTemp = Nvl(mobjDelBalance.rsBalance!���㷽ʽ)
            End If
            If mCurCarge.dbl��ǰδ�� < 0 Then '�˷�
                If mblnSingleBalance Then
                    If mobjDelBalance.ȱʡ���㷽ʽ = objCard.���㷽ʽ Then
                        If objCard.�ӿ���� > 0 Then
                            '�����ʻ���,�����ȱʡ����,��ȱʡ�������ʻ�
                            If Not (objCard.�Ƿ����� And objCard.�Ƿ�ȱʡ����) Then .ListIndex = i
                        ElseIf InStr(gTy_Module_Para.strȱʡ����, objCard.���㷽ʽ) = 0 Then
                            'û������ȱʡ����,��ȱʡΪ�ý��㷽ʽ
                            .ListIndex = i
                        End If
                    End If
                Else
                    '�����ʻ�:���������ʻ���,�����ȱʡ����,��ȱʡ�������ʻ�,���ڶ�������ʻ�֧��,��ȱʡΪ��һ�������ʻ�.
                    If objCard.�ӿ���� > 0 Then
                        If Not (objCard.�Ƿ����� And objCard.�Ƿ�ȱʡ����) And blnSetedIndex = False Then
                            '93114���ɷ�ʱδʹ�õĲ�ȱʡ
                            Call CheckThreeSwapCanTransfer(objCard, mobjDelBalance.ԭ����ID, blnChargeUsed)
                            If blnChargeUsed Then .ListIndex = i: blnSetedIndex = True
                        End If
                    End If
                    '��ʹ��Ԥ�����ȱʡԤ����
                    If objCard.�������� = -99 And .ListIndex < 0 Then .ListIndex = i
                    '�տ���㷽ʽ�д���һ�ַ�ҽ�����㷽ʽ��,���û������ȱʡ����,��ȱʡΪ�ý��㷽ʽ
                    If InStr(gTy_Module_Para.strȱʡ����, objCard.���㷽ʽ) = 0 Then
                        If strTemp = objCard.���㷽ʽ And .ListIndex < 0 Then .ListIndex = i
                    End If
                End If
                '�տ���㷽ʽ�д���ȱʡ�Ľ��㷽ʽ,����ȱʡ��Ϊ׼
                If objCard.ȱʡ��־ And .ListIndex < 0 Then .ListIndex = i
                '���ֽ�Ϊ׼
            Else
                If objCard.ȱʡ��־ And .ListIndex < 0 Then .ListIndex = i
                If objCard.�������� = 1 And .ListIndex < 0 Then .ListIndex = i
                If mobjDelBalance.ȱʡ���㷽ʽ = objCard.���㷽ʽ Then .ListIndex = i
            End If
        Next
        If gstr���㷽ʽ <> "" And .ListIndex < 0 Then
            For i = 0 To .ListCount - 1
                If .List(i) = gstr���㷽ʽ Then
                    .ListIndex = i: Exit For
                End If
            Next
        End If
        If blnǿ������ Then
            For i = .ListCount - 1 To 0 Step -1
                Set objCard = mobjPayCards(.ItemData(i))
                If objCard.�������� = 1 Then .ListIndex = i: Exit For
            Next
        End If
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Call cbo֧����ʽ_Click
End Sub

Private Sub Set�˷ѷ�ʽ(ByVal bytType As Byte, Optional ByVal objCard As Card, Optional ByVal bln���� As Boolean, _
    Optional ByVal blnǿ������ As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ָ�����˷ѷ�ʽ�Ƿ�����
    '����:
    '   bytType:1=���ݴ��뿨�������ý��㷽ʽ�Ƿ����
    '           2=�����˷ѽ��㷽ʽ�б�ͽ������������˷ѽ��㷽ʽ�Ƿ����
    '           3=�����˷ѽ��㷽ʽ�б������շѽ��㷽ʽ�Ƿ����
    '����:���˺�
    '����:2014-09-01 16:14:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTemp As Card
    Dim i As Long, j As Long, blnFind As Boolean, dblMoney As Double
    Dim rsTemp As ADODB.Recordset, blnDefault As Boolean
    
    On Error GoTo Errhand
    Select Case bytType
        Case 1
            If objCard Is Nothing Then Exit Sub
            For Each objTemp In mobjPayCards
                If objTemp.�ӿ���� = objCard.�ӿ���� And objTemp.���ѿ� = objCard.���ѿ� Then
                    objTemp.֧������ = bln����
                End If
            Next
        Case 2
            Set rsTemp = mobjDelBalance.rsBalance
            For i = 1 To mobjPayCards.Count
                Set objTemp = mobjPayCards(i)
                dblMoney = 0: blnFind = False: blnDefault = False
                '�жϽ��㷽ʽʣ��δ�˽��
                With rsTemp
                    rsTemp.Filter = 0
                    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
                    Do While Not rsTemp.EOF
                        '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                        If Nvl(rsTemp!����) <> 1 Then 'Ԥ�����ͨ��"���㷽ʽ"�����ж�
                            If Nvl(rsTemp!���㷽ʽ) = objTemp.���㷽ʽ Then dblMoney = dblMoney + Nvl(rsTemp!��Ԥ��)
                        End If
                        rsTemp.MoveNext
                    Loop
                End With
                dblMoney = RoundEx(dblMoney, 6)
                '�ж��Ƿ����˷ѽ����б���
                For j = 1 To vsBlance.Rows - 1
                    If vsBlance.TextMatrix(j, vsBlance.ColIndex("֧����ʽ")) = objTemp.���㷽ʽ Then
                        blnFind = True: Exit For
                    End If
                Next
                '��������:1-�ֽ���㷽ʽ,2-������ҽ������,3-ҽ�������ʻ�,4-ҽ������ͳ��,5-���տ���,6-�����ۿ�,7-һ��ͨ����,8-���㿨����
                If objTemp.�������� = 1 Or objTemp.�������� = 2 Then blnDefault = True
                If blnǿ������ Then
                    'ǿ������ʱ������ת�ʼ�����
                    objTemp.֧������ = (RoundEx(dblMoney, 6) <> 0 Or blnDefault) And Not blnFind
                Else
                    objTemp.֧������ = (RoundEx(dblMoney, 6) <> 0 Or blnDefault Or objTemp.�Ƿ�ת�ʼ�����) And Not blnFind
                End If
            Next
        Case 3
            For i = 1 To mobjPayCards.Count
                Set objTemp = mobjPayCards(i)
                '�ж��Ƿ����˷ѽ����б��У�����ˢ������ѿ�
                blnFind = False
                For j = 1 To vsBlance.Rows - 1
                    '����:0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
                    With vsBlance
                        If .TextMatrix(j, .ColIndex("֧����ʽ")) = objTemp.���㷽ʽ And _
                            (Val(.TextMatrix(j, .ColIndex("����"))) <> 5 Or _
                            (Val(.TextMatrix(j, .ColIndex("����"))) = 5 And .Cell(flexcpData, j, .ColIndex("֧�����")) < 0)) Then
                            blnFind = True: Exit For
                        End If
                    End With
                Next
                objTemp.֧������ = Not blnFind
            Next
        Case Else
    
    End Select
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CheckInterfaceNumIsValied(ByVal objCard As Card) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ӿ������Ƿ񳬹�2������
    '����:δ����2������,����true,���򷵻�False
    '����:���˺�
    '����:2012-02-27 15:23:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long, varData As Variant
    Dim strNames As String, i As Long
    
    On Error GoTo errHandle
    
    lngCount = IIf(mobjDelBalance.intInsure <> 0, 1, 0)   'ҽ����һ������
'    If objCard.�ӿ���� <= 0 Or (objCard.���ѿ� And objCard.���ƿ�) Then CheckInterfaceNumIsValied = True: Exit Function
    With vsBlance
        strNames = vbCrLf & IIf(mobjDelBalance.intInsure <> 0, "ҽ������", "")
        For i = 1 To .Rows - 1
            '0-��ͨ����;1-Ԥ����;2-ҽ��,3-һ��ͨ;4-һ��ͨ(��);5-���ѿ�
            If Val(.RowData(i)) = 3 Or Val(.RowData(i)) = 4 Or Val(.RowData(i)) = 5 Then
                ' ҽ�ƿ����ID|���ѿ�(1, 0) |���ƿ�|�Ƿ�ȫ��|�Ƿ�����|�ӿ�����
                varData = Split(.Cell(flexcpData, i, .ColIndex("֧����ʽ")) & "|||||", "|")
                If Val(varData(0)) <> 0 Then
                    If Val(varData(1)) <> 1 Then
                        lngCount = lngCount + 1
                        If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 1 Then
                            strNames = strNames & vbCrLf & varData(5)
                        End If
                    ElseIf Val(varData(2)) = 0 Then
                        '���ѿ�Ҳ�ǽӿڵ�,�������������ӿ�
                        lngCount = lngCount + 1
                        If Val(.TextMatrix(i, .ColIndex("����״̬"))) = 1 Then
                            strNames = strNames & vbCrLf & varData(5)
                        End If
                    End If
                End If
            End If
        Next
    End With
    If lngCount = 2 Then
        If objCard.�ӿ���� <= 0 Or (objCard.���ѿ� And objCard.���ƿ�) Then CheckInterfaceNumIsValied = True: Exit Function
        MsgBox "  ϵͳ��ֻ֧���������ڵĽӿڣ�������ˢ�����ѣ����飡" & vbCrLf & "����Ϊ��ǰ�Ѿ�ˢ�Ľӿڣ�" & vbCrLf & strNames, vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf lngCount > 2 Then
        MsgBox "  ϵͳ��ֻ֧���������ڵĽӿڣ�������ˢ�����ѣ����飡" & vbCrLf & "����Ϊ��ǰ�Ѿ�ˢ�Ľӿڣ�" & vbCrLf & strNames, vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    CheckInterfaceNumIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function CheckThreeSwapCanTransfer(ByVal objCard As Card, ByVal lng����ID As Long, _
    Optional ByRef blnChargeUsed As Boolean) As Boolean
    '����������Ƿ��ʹ��ת�ʷ�ʽ�˿�
    '�����:93114
    '˵����
    '   ��ʹ��ת�ʹ��ܵ�������1.֧��ת�ʼ����ۣ�2.�ڽɷ�ʱδʹ�û����ڽɷ�ʱʹ������������
    '   �ڽɷ�ʱʹ���˲������ֵ�������ֻ��ԭ���˻�
    Dim strSQL As String
    
    On Error GoTo errHandle
    blnChargeUsed = False
    If objCard Is Nothing Then Exit Function
    If objCard.�ӿ���� <= 0 Then Exit Function
    
    If mrsUsedCards Is Nothing Then
        '���棬��ֹ������ѯ���ݿ�
        strSQL = _
            "Select Nvl(a.�����id,a.���㿨���) As �����id" & vbNewLine & _
            " From ����Ԥ����¼ A," & vbNewLine & _
            "      (Select m.����id" & vbNewLine & _
            "        From ������ü�¼ M, ������ü�¼ N" & vbNewLine & _
            "        Where m.��¼���� = n.��¼���� And m.No = n.No And n.����id = [1] And m.��¼���� = 1) B" & vbNewLine & _
            " Where a.����id = b.����id And a.��¼״̬ In (1, 3) And (Nvl(a.�����id, 0) <> 0 Or Nvl(a.���㿨���, 0) <> 0)"
        Set mrsUsedCards = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    End If
    mrsUsedCards.Filter = "�����id=" & objCard.�ӿ����
    If mrsUsedCards.EOF Then
        CheckThreeSwapCanTransfer = True
    Else
        blnChargeUsed = True '�ɷ�ʱʹ����
        CheckThreeSwapCanTransfer = objCard.�Ƿ�����
    End If
    
    If objCard.�Ƿ�ת�ʼ����� = False Then CheckThreeSwapCanTransfer = False
    'ǿ������ʱ��������ʹ��ת�ʼ�����
    If mcllForceDelToCash.Count > 0 Then CheckThreeSwapCanTransfer = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



