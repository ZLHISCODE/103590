VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmNewCheckMain 
   BackColor       =   &H80000005&
   Caption         =   "ҩƷ�̵����"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12690
   Icon            =   "frmNewCheckMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9000
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   2
      Top             =   7560
      Width           =   2175
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   4
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   1680
         TabIndex        =   6
         Top             =   37
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "��������"
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   37
         Width           =   720
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1605
      Left            =   10320
      TabIndex        =   21
      Top             =   1800
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   2831
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame fraCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   1335
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   12615
      Begin VB.CommandButton cmd���� 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   10200
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   840
         Width           =   1100
      End
      Begin VB.TextBox Txt����� 
         Height          =   300
         Left            =   11160
         MaxLength       =   8
         TabIndex        =   31
         Top             =   120
         Width           =   1365
      End
      Begin VB.TextBox Txt������ 
         Height          =   300
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   29
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton CmdҩƷ 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   300
         Left            =   12240
         TabIndex        =   37
         Top             =   510
         Width           =   255
      End
      Begin VB.TextBox TxtҩƷ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   8970
         MaxLength       =   50
         ScrollBars      =   3  'Both
         TabIndex        =   36
         Top             =   510
         Width           =   3255
      End
      Begin VB.CheckBox ChkҩƷ 
         BackColor       =   &H80000003&
         Caption         =   "ҩƷ"
         Height          =   300
         Left            =   8280
         TabIndex        =   35
         Top             =   510
         Width           =   735
      End
      Begin VB.TextBox txt����NO 
         Height          =   300
         Left            =   6450
         MaxLength       =   8
         TabIndex        =   27
         Top             =   120
         Width           =   1605
      End
      Begin VB.TextBox txt��ʼNo 
         Height          =   300
         Left            =   4560
         MaxLength       =   8
         TabIndex        =   25
         Top             =   120
         Width           =   1605
      End
      Begin VB.CheckBox chkStrike 
         BackColor       =   &H80000003&
         Caption         =   "��������"
         Enabled         =   0   'False
         Height          =   300
         Left            =   8280
         TabIndex        =   41
         Top             =   907
         Width           =   1095
      End
      Begin VB.CommandButton cmdȷ�� 
         Caption         =   "ȷ��(&S)"
         Height          =   350
         Left            =   11400
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   840
         Width           =   1100
      End
      Begin VB.ComboBox cbo����� 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   900
         Width           =   1560
      End
      Begin VB.ComboBox cboδ��� 
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   510
         Width           =   1560
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1560
         TabIndex        =   23
         Text            =   "cboStock"
         Top             =   120
         Width           =   1920
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Index           =   0
         Left            =   4560
         TabIndex        =   33
         Top             =   503
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Index           =   0
         Left            =   6465
         TabIndex        =   34
         Top             =   503
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp��ʼʱ�� 
         Height          =   315
         Index           =   1
         Left            =   4560
         TabIndex        =   39
         Top             =   900
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin MSComCtl2.DTPicker dtp����ʱ�� 
         Height          =   315
         Index           =   1
         Left            =   6465
         TabIndex        =   40
         Top             =   900
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy��MM��dd��"
         Format          =   200736771
         CurrentDate     =   36263
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "�����"
         Height          =   180
         Left            =   10560
         TabIndex        =   30
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "������"
         Height          =   180
         Left            =   8280
         TabIndex        =   28
         Top             =   180
         Width           =   540
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   6225
         TabIndex        =   26
         Top             =   180
         Width           =   180
      End
      Begin VB.Label LblNO 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "No"
         Height          =   180
         Left            =   3660
         TabIndex        =   24
         Top             =   180
         Width           =   180
      End
      Begin VB.Label lbl����� 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "����˵���"
         Height          =   180
         Left            =   420
         TabIndex        =   20
         Top             =   960
         Width           =   900
      End
      Begin VB.Label lblδ��� 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "δ��˵���"
         Height          =   180
         Left            =   420
         TabIndex        =   19
         Top             =   570
         Width           =   900
      End
      Begin VB.Label lblStock 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "��      ��"
         Height          =   180
         Left            =   420
         TabIndex        =   22
         Top             =   180
         Width           =   900
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   6225
         TabIndex        =   11
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "�������"
         Height          =   180
         Index           =   1
         Left            =   3660
         TabIndex        =   10
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   6225
         TabIndex        =   9
         Top             =   570
         Width           =   180
      End
      Begin VB.Label lblʱ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "��������"
         Height          =   180
         Index           =   0
         Left            =   3660
         TabIndex        =   8
         Top             =   570
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Height          =   5415
      Left            =   1680
      ScaleHeight     =   5355
      ScaleWidth      =   8475
      TabIndex        =   1
      Top             =   1800
      Width           =   8535
      Begin VB.CommandButton Cmd���� 
         Caption         =   "����(&V)"
         Height          =   350
         Left            =   7320
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1100
      End
      Begin VB.PictureBox picSeparate_s 
         BorderStyle     =   0  'None
         Height          =   370
         Left            =   0
         MousePointer    =   7  'Size N S
         ScaleHeight     =   375
         ScaleWidth      =   7935
         TabIndex        =   13
         Top             =   2520
         Width           =   7935
         Begin VB.Label lbl2 
            AutoSize        =   -1  'True
            Caption         =   "����ϼƣ�"
            Height          =   180
            Left            =   1680
            TabIndex        =   18
            Top             =   120
            Width           =   1080
         End
         Begin VB.Label lbl1 
            AutoSize        =   -1  'True
            Caption         =   "�̵���ϼƣ�"
            Height          =   180
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label lbl3 
            AutoSize        =   -1  'True
            Caption         =   "�������ϼƣ�"
            Height          =   180
            Left            =   3000
            TabIndex        =   16
            Top             =   120
            Width           =   1440
         End
         Begin VB.Label lblSum�ɱ���� 
            AutoSize        =   -1  'True
            Caption         =   "�̵�ɱ����ϼƣ�"
            Height          =   180
            Left            =   4680
            TabIndex        =   15
            Top             =   120
            Width           =   1620
         End
         Begin VB.Label lbl�ɱ����� 
            AutoSize        =   -1  'True
            Caption         =   "�ɱ�����ϼƣ�"
            Height          =   180
            Left            =   6480
            TabIndex        =   14
            Top             =   120
            Width           =   1440
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1455
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   6255
         _cx             =   11033
         _cy             =   2566
         Appearance      =   1
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
         Height          =   1155
         Left            =   0
         TabIndex        =   46
         Top             =   4200
         Width           =   6255
         _cx             =   11033
         _cy             =   2037
         Appearance      =   1
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
         BackColorSel    =   16053482
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   15724527
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7800
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17304
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tbcDetail 
      Height          =   975
      Left            =   0
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
      _Version        =   589884
      _ExtentX        =   2566
      _ExtentY        =   1720
      _StockProps     =   64
      Enabled         =   -1  'True
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   1920
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmNewCheckMain.frx":06EA
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmNewCheckMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl
Private mcbrMenuBar As CommandBarPopup
Private mcbrToolBar As CommandBar

'�̵������
Private Const mconTab_CheckCourseCard = 0                 '�̵��¼���б�
Private Const mconTab_CheckCard = 1                       '�̵��б�

Private mblnLoad As Boolean
Private mbln�� As Boolean      '�ж��Ƿ�����
Private mintLastIndex As Integer '������һ�ε����Tab
Private mstrSelectTag As String     '��ǰѡ����������˻��������

Private mlngMode As Long
Private mstrFind As String
Private mblnBootUp As Boolean
Private mlastRow As Long                '�ϴε������
Private mstrTitle As String             '����ı���
Private mblnViewCost As Boolean         '�鿴�ɱ���
'Private Const mstrTitle As String = "ҩƷ�̵����"

Public mstrPrivs As String              'Ȩ��

'��������
Private mstrStart As Date
Private mstrEnd As Date
Private mstrVerifyStart As Date
Private mstrVerifyEnd As Date

Private mlng�ⷿID As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private Const mcstComment As String = "��-��ƽ;��-��ӯ;��-�̿�;����-ͣ��ҩƷ"

'�Ӳ�������ȡҩƷ�۸����������С��λ������ʾ���ȣ�
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private mintMaxMoneyBit As Integer          'ҩƷ�����н��С��λ��
Private mstrMaxMoneyFormat As String

Private mbln����ģʽ As Boolean

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    lngҩƷ As Long
    lng����ⷿ As Long
    str������ As String
    str����� As String
    lngҩƷ���� As Long
    str���� As String
End Type

Private SQLCondition As Type_SQLCondition



Private Function InitComandBars() As Boolean
    '----------------------------------------------------------------------------------------
    '����:��ʼ���˵���������
    '����:���óɹ�,����true,���򷵻�False
    '----------------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl

    Dim panThis As Pane
    Err = 0: On Error GoTo ErrHand:
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    
'    Set cbsThis.Icons = zlCommFun.GetPubIcons
    Set cbsThis.Icons = imgPublic.Icons
    
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_FilePopup, "�ļ�(&F)", -1, False)
    mcbrMenuBar.id = mconMenu_FilePopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_PrintSet, "��ӡ����(&S)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Preview, "��ӡԤ��(&V)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ(&P)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_BillPrint, "���ݴ�ӡ(&B)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_BillPreview, "����Ԥ��(&L)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Excel, "�����Excel"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Parameter, "��������(&R)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�(&X)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_EditPopup, "�༭(&E)", -1, False)
    mcbrMenuBar.id = mconMenu_EditPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddBill, "���Ӽ�¼��(&B)")
        
        Set mcbrControl = .Add(xtpControlPopup, mconMenu_Edit_AddTable, "�����̵��(&T)")
        mcbrControl.id = mconMenu_Edit_AddTable
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableAuto, "�Զ������̵��(&A)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableTotal, "���ܼ�¼�������̵��(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableZero, "ȫ����Ϊ��(&Z)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableHouseAll, "�ⷿȫ��ҩƷ�̵�(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableSpecial, "����ҩƷ�̵�(&S)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddModify, "�޸�(&M)")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddDel, "ɾ��(&D)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddVerify, "���(&C)"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddStrike, "����(&K)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddAffirmant, "�¶�ȷ��(&O)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddDisplay, "�鿴����(&W)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_CheckTable, "�̵�����ܼ��(&T)"): mcbrControl.BeginGroup = True
    End With

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_ViewPopup, "�鿴(&V)")
    mcbrMenuBar.id = mconMenu_ViewPopup
    With mcbrMenuBar.CommandBar.Controls

        Set mcbrControl = .Add(xtpControlButton, mconMenu_View_StatusBar, "״̬��(&S)")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_View_Refresh, "ˢ��(&R)"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_View_ColSet, "������(&C)"): mcbrControl.BeginGroup = True
        
    End With
    

    Set mcbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_HelpPopup, "����(&H)", -1, False)
    mcbrMenuBar.id = mconMenu_HelpPopup
    With mcbrMenuBar.CommandBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Help_Help, "��������(&H)")
        
        Set mcbrControl = .Add(xtpControlPopup, mconMenu_Help_Web, "&WEB�ϵ�����")
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Home, "������ҳ(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Forum, "������̳(&F)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Help_Web_Mail, "���ͷ���(&M)��", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Help_About, "����(&A)��"): mcbrControl.BeginGroup = True
        
    End With

    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("P"), mconMenu_File_Print 'Ctrl+P
        .Add FCONTROL, Asc("B"), mconMenu_File_BillPrint
        .Add FCONTROL, Asc("A"), mconMenu_Edit_AddBill
        .Add 0, VK_DELETE, mconMenu_Edit_AddDel
        .Add 0, VK_F12, mconMenu_File_Parameter
        .Add 0, VK_F5, mconMenu_View_Refresh
        .Add 0, VK_F1, mconMenu_Help_Help
        .Add 0, VK_ESCAPE, mconMenu_File_Exit
    End With

    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand mconMenu_File_PrintSet
        .AddHiddenCommand mconMenu_File_Excel
    End With

    '-----------------------------------------------------
    '����������
    Set mcbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched
    With mcbrToolBar.Controls
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Preview, "Ԥ��")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Print, "��ӡ")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddBill, "��¼��"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlPopup, mconMenu_Edit_AddTable, "�̵��")
        mcbrControl.id = mconMenu_Edit_AddTable
        mcbrControl.IconId = mconMenu_Edit_AddBill
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableAuto, "�Զ������̵��(&A)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableTotal, "���ܼ�¼�������̵��(&T)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableZero, "ȫ����Ϊ��(&Z)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableHouseAll, "�ⷿȫ��ҩƷ�̵�(&H)", -1, False
        mcbrControl.CommandBar.Controls.Add xtpControlButton, mconMenu_Edit_AddTableSpecial, "����ҩƷ�̵�(&S)", -1, False
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddModify, "�޸�")
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddDel, "ɾ��")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddVerify, "���"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddStrike, "����")
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_AddAffirmant, "�¶�ȷ��"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Edit_CheckTable, "�̵�����ܼ��"): mcbrControl.BeginGroup = True
        
        Set mcbrControl = .Add(xtpControlButton, mconMenu_Help_Help, "����"): mcbrControl.BeginGroup = True
        Set mcbrControl = .Add(xtpControlButton, mconMenu_File_Exit, "�˳�")
    End With
    For Each mcbrControl In mcbrToolBar.Controls
        mcbrControl.Style = xtpButtonIconAndCaption
    Next
    InitComandBars = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume

End Function
Private Sub cboStock_Click()
    If mlng�ⷿID <> Me.cboStock.ItemData(Me.cboStock.ListIndex) Then
        mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
        Call GetDrugDigit(mlng�ⷿID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '������֯��ʽ����
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
        mstrMaxMoneyFormat = "'999999999990." & String(mintMaxMoneyBit, "0") & "'"
    End If
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    
    str�������� = "H,I,J,K,L,M,N"

    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), str��������, IIf(zlStr.IsHavePrivs(mstrPrivs, "���пⷿ"), False, True)) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '�������뵥����
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    If cboStock.ListCount > 0 Then
        If cboStock.ListIndex = -1 Then
            MsgBox "��ѡ��һ��ҩ�����ҩ����", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub cboδ���_Click()
    Dim dateCurrentDate As Date
    
    If cboδ���.ListIndex = cboδ���.ListCount - 1 Then 'ѡ���Զ�������ѡ���ſ���
        dtp��ʼʱ��(0).Enabled = True
        dtp����ʱ��(0).Enabled = True
    Else
        dtp��ʼʱ��(0).Enabled = False
        dtp����ʱ��(0).Enabled = False
    End If
    
    '����ѡ��ı�ʱ��
    dateCurrentDate = Sys.Currentdate
    Select Case cboδ���.ListIndex
        Case 1
            dtp��ʼʱ��(0).Value = Format(DateAdd("d", 0, dateCurrentDate), "yyyy-MM-dd")
            dtp����ʱ��(0).Value = dateCurrentDate
        Case 2
            dtp��ʼʱ��(0).Value = Format(DateAdd("d", -6, dateCurrentDate), "yyyy-MM-dd")
            dtp����ʱ��(0).Value = dateCurrentDate
        Case 3
            dtp��ʼʱ��(0).Value = Format(DateAdd("m", 0, dateCurrentDate), "yyyy-MM")
            dtp����ʱ��(0).Value = dateCurrentDate
    End Select
    
End Sub

Private Sub cbo�����_Click()
    Dim dateCurrentDate As Date

    If cbo�����.ListIndex = cbo�����.ListCount - 1 And cbo�����.Enabled Then    'ѡ���Զ�������ѡ���ſ���
        dtp��ʼʱ��(1).Enabled = True
        dtp����ʱ��(1).Enabled = True
    Else
        dtp��ʼʱ��(1).Enabled = False
        dtp����ʱ��(1).Enabled = False
    End If
    chkStrike.Enabled = cbo�����.ListIndex <> 0 And cbo�����.Enabled
    
    '����ѡ��ı�ʱ��
    dateCurrentDate = Sys.Currentdate
    Select Case cbo�����.ListIndex
        Case 1
            dtp��ʼʱ��(1).Value = Format(DateAdd("d", 0, dateCurrentDate), "yyyy-MM-dd")
            dtp����ʱ��(1).Value = dateCurrentDate
        Case 2
            dtp��ʼʱ��(1).Value = Format(DateAdd("d", -6, dateCurrentDate), "yyyy-MM-dd")
            dtp����ʱ��(1).Value = dateCurrentDate
        Case 3
            dtp��ʼʱ��(1).Value = Format(DateAdd("m", 0, dateCurrentDate), "yyyy-MM")
            dtp����ʱ��(1).Value = dateCurrentDate
    End Select
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        '�ļ�
        Case mconMenu_File_PrintSet
            cbsFilePrintSet '��ӡ����
        Case mconMenu_File_Preview
            cbsFilePreView '��ӡԤ��
        Case mconMenu_File_Print
            cbsFilePrint '��ӡ
        Case mconMenu_File_BillPrint
            cbsFileBillPrint '���ݴ�ӡ
        Case mconMenu_File_BillPreview
            cbsFileBillPreview '����Ԥ��
        Case mconMenu_File_Excel
            cbsFileExcel '�����&Excel
        Case mconMenu_File_Parameter
            cbsFileParameter '��������
        Case mconMenu_File_Exit
            cbsfileExit '�˳�
        '�༭
        Case mconMenu_Edit_AddBill
            cbsEditaddBill '���Ӽ�¼��
        Case mconMenu_Edit_AddTableAuto
            cbsAddTableAuto '�Զ������̵��
        Case mconMenu_Edit_AddTableTotal
            cbsAddTableTotal '���ܼ�¼�������̵��
        Case mconMenu_Edit_AddTableZero
            cbsAddTableZero 'ȫ����Ϊ��
        Case mconMenu_Edit_AddTableHouseAll
            cbsAddTableHouseAll '�ⷿȫ��ҩƷ�̵�
        Case mconMenu_Edit_AddTableSpecial
            cbsAddTableSpecial '����ҩƷ�̵�
        Case mconMenu_Edit_AddModify
            cbsEditModify '�޸�
        Case mconMenu_Edit_AddDel
            cbsEditDel 'ɾ��
        Case mconMenu_Edit_AddVerify
            cbsVerify '���
        Case mconMenu_Edit_AddStrike
            cbsEditStrike '����
        Case mconMenu_Edit_AddAffirmant
            cbsAffirmant 'ȷ��
        Case mconMenu_Edit_AddDisplay
            cbsDisplay '�鿴����
        Case mconMenu_Edit_CheckTable
            cbsCheckTable '�̵�����ܼ��
            
        '�鿴
        Case mconMenu_View_StatusBar
            cbsViewStatus '״̬��
        Case mconMenu_View_Refresh
            cbsViewRefresh 'ˢ��
        Case mconMenu_View_ColSet
            cbsViewColSet '������
        
        '����
        Case mconMenu_Help_Help
            cbsHelpTitle '��������
        Case mconMenu_Help_Web_Home
            cbsHelpWebHome '������ҳ
        Case mconMenu_Help_Web_Forum
            cbsHelpWebForum '������̳
        Case mconMenu_Help_Web_Mail
            cbsHelpWebMail '���ͷ���
        Case mconMenu_Help_About
            cbsHelpAbout '����
        Case Else
            If Control.id > 401 And Control.id < 499 Then
                'ִ���Զ��屨��
                Call BillPrint_Custom(Control)
            End If
    End Select
    
End Sub

Private Sub cbsEditaddBill()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmNewCheckCourseCard.ShowCard Me, strNo, 1, , blnSuccess
    
    If blnSuccess Then Call cbsViewRefresh
End Sub

Private Sub cbsViewStatus()
    Dim cbrMenuPop As CommandBarControl
    
    Set cbrMenuPop = Me.cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_StatusBar, , True)
    
    With cbrMenuPop
        .Checked = Not .Checked  ' Xor True
        stbThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub BillPrint_Custom(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '��ӡ�Զ��屨��
    'Ĭ�ϲ�����ҩƷ=ҩƷid���ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬�̵㵥=�̵㵥NO���̵��=�̵��NO
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim strNo As String
    Dim strReportName As String

    strReportName = Split(Control.Parameter, ",")(1)

    Select Case strReportName
        Case "ZL1_INSIDE_1307"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307", Me, "�ⷿ=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)))
        Case "ZL1_INSIDE_1307_1"
            Call ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1307_1", Me, "�ⷿ=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)), "��λ=" & Choose(mintUnit, "�ۼ۵�λ", "���ﵥλ", "סԺ��λ", "ҩ�ⵥλ") & "|" & Choose(mintUnit, 1, 3, 4, 2))
        Case Else
            If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
                strNo = vsfList.TextMatrix(vsfList.Row, 0)
            End If

            str��ʼʱ�� = IIf(Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd"))
            str����ʱ�� = IIf(Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd"))

            Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), strReportName, Me, _
                "ҩƷ=" & IIf(SQLCondition.lngҩƷ = 0, "", SQLCondition.lngҩƷ), _
                "�ⷿ=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
                "��ʼʱ��=" & str��ʼʱ��, _
                "����ʱ��=" & str����ʱ��, _
                "�̵㵥=" & strNo, _
                "�̵��=" & strNo)
    End Select
End Sub



Private Sub cbsViewRefresh()
    'ˢ��
    GetList mstrFind
End Sub


Private Sub cbsViewColSet()
    Dim strColsName As String '�������ε���
    Dim strDefaultColsName As String 'Ĭ�ϵ���
    Dim i As Integer
    Dim strColName As String
    '������
    strDefaultColsName = ":ҩƷ��Դ,0:����ҩ��,0:�ⷿ��λ,0:��׼�ĺ�,0:����,0:��۲�,0:�̵�ɱ�����,0:�������,0:�ɱ�����,0:��ǰ���,1:" '���п������ص���
    
    
    strColsName = zlDataBase.GetPara("������", glngSys, mlngMode, "") '��ȡ���ݿ�ı�����Ϣ
    
    '���ݴ���
    If strColsName = "" Then 'δ��ȡ����������Ϣ
        strColsName = strDefaultColsName
    Else
        '�ж���ȡ������Ĭ���и�������һ����ȡĬ�ϵ�
        If UBound(Split(strColsName, ":")) <> UBound(Split(strDefaultColsName, ":")) Then strColsName = strDefaultColsName
        
        '�ж���ȡ�������Ƿ���Ĭ�ϵ�һ�£���һ��ȡĬ�ϵ�
        For i = LBound(Split(strColsName, ":")) + 1 To UBound(Split(strColsName, ":")) - 1
            strColName = Split(Split(strColsName, ":")(i), ",")(0) '��ȡ��������
            
            If InStr(1, strDefaultColsName, ":" & strColName) = 0 Then '������������Ĭ��������
                strColsName = strDefaultColsName
                Exit For
            End If
        Next
        
    End If
    
    strColsName = frm����������.ShowME(Me, strColsName)
    
    If strColsName <> "" Then
        zlDataBase.SetPara "������", strColsName, glngSys, mlngMode
    End If
    
    cbsViewRefresh

End Sub

Private Sub cbsHelpAbout()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub cbsHelpTitle()
    '��������
    Dim StrWinName As String
    With vsfList
        StrWinName = "frmMainList8"
    End With
    Call ShowHelp(App.ProductName, Me.hWnd, StrWinName)
End Sub

Private Sub cbsHelpWebHome()
    '������ҳ
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub cbsHelpWebForum()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub cbsHelpWebMail()
    '���ͷ���
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub cbsFilePreView()
    '��ӡԤ��
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsFilePrint()
    '��ӡ
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub cbsFilePrintSet()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub cbsFileParameter()
    '��������
    Dim int��ѯ����  As Integer
    
    frm��������.���ò��� Me, mstrPrivs, mstrTitle
    
    '������Ҫ�䶯
    int��ѯ���� = Val(zlDataBase.GetPara("��ѯ����", glngSys, mlngMode, 7))
    int��ѯ���� = IIf(int��ѯ���� <> 1 And int��ѯ���� <> 7, 7, int��ѯ����)
    
    cboδ���.ListIndex = IIf(int��ѯ���� = 7, 2, 1)
    
    cmdȷ��_Click 'ȷ��ˢ�½���
    
End Sub

Private Sub cbsfileExit()
    '�˳�
    Unload Me
End Sub

Private Sub cbsFileExcel()
    '�����Excel
    
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfDetail Then
        vsfDetail.Redraw = flexRDNone
        subExcel 3
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
End Sub

Private Sub cbsFileBillPrint()
    Dim int��λϵ�� As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                int��λϵ�� = 4
            Case mconint���ﵥλ
                int��λϵ�� = 2
            Case mconintסԺ��λ
                int��λϵ�� = 1
            Case mconintҩ�ⵥλ
                int��λϵ�� = 3
        End Select
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 2
    End With
End Sub

Private Sub cbsFileBillPreview()
    Dim int��λϵ�� As Integer
    
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                int��λϵ�� = 4
            Case mconint���ﵥλ
                int��λϵ�� = 2
            Case mconintסԺ��λ
                int��λϵ�� = 1
            Case mconintҩ�ⵥλ
                int��λϵ�� = 3
        End Select
        ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), "��¼״̬=" & .TextMatrix(.Row, .Cols - 2), "��λϵ��=" & int��λϵ��, 1
    End With
End Sub


Private Sub cbsEditModify()
    '�޸�
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        If tbcDetail.Selected.Index = 0 Then
            frmNewCheckCourseCard.ShowCard Me, strNo, 2, 1, blnSuccess
        Else
            frmNewCheckCard.ShowCard Me, strNo, 2, vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2), blnSuccess
        End If
        
        If blnSuccess Then Call cbsViewRefresh
    End With
End Sub

Private Sub cbsEditStrike()
    Dim blnPurchase As Boolean, blnRefresh As Boolean
    
    '������⹺(blnPurchaseΪ��)����ֱ�ӽ������
    'ѯ���Ƿ����(blnPurchaseΪ��ʾ�򷵻�ֵ)������������
    blnPurchase = (InStr(1, "1300,1302,1304,1305,1306", mlngMode) <> 0)
    With vsfList
        If Not blnPurchase Then
            blnPurchase = (MsgBox("��ȷʵҪȫ���������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���ĵ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
        End If
        If blnPurchase Then
            blnRefresh = StrikeSave
            If blnRefresh Then cbsViewRefresh
        End If
    End With
End Sub


Private Function StrikeSave() As Boolean
    Dim blnSuccess As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim int����� As Integer
    Dim strMsg As String
    Dim n As Integer
    
    StrikeSave = False
    
    int����� = MediWork_GetCheckStockRule(mlng�ⷿID)
    
    On Error GoTo ErrHandle
    If int����� <> 0 Then
        gstrSQL = "Select A.ҩƷ��Ϣ " & _
            " From (Select Distinct '(' || I.���� || ')' || Nvl(N.����, I.����) As ҩƷ��Ϣ, A.ʵ������, Nvl(K.ʵ������, 0) As ������� " & _
            " From ҩƷ�շ���¼ A, (Select ҩƷid, �ⷿid, ʵ������, Nvl(����, 0) ���� From ҩƷ��� Where ���� = 1) K, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N " & _
            " Where A.ҩƷid = K.ҩƷid(+) And A.�ⷿid = K.�ⷿid(+) And Nvl(A.����, 0) = K.����(+) And A.ҩƷid = B.ҩƷid And " & _
            " A.ҩƷid = I.ID And A.ҩƷid = N.�շ�ϸĿid(+) And N.����(+) = 3 And A.���� = 12 And A.���ϵ�� = 1 And A.NO = [1]) A " & _
            " Where A.ʵ������ > A.������� "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����", vsfList.TextMatrix(vsfList.Row, 0))
        
        With rsTemp
            If .RecordCount > 0 Then
                For n = 1 To .RecordCount
                    If n > 5 Then
                        strMsg = strMsg & vbCrLf & "��������" & .RecordCount - 5 & "��ҩƷ......"
                        Exit For
                    End If
                    strMsg = IIf(strMsg = "", "", strMsg & "," & vbCrLf) & !ҩƷ��Ϣ
                    .MoveNext
                Next
                
                If int����� = 1 Then
                    If MsgBox("ע�⣬����ҩƷ��治�㣺" & vbCrLf & strMsg & vbCrLf & Space(4) & "�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                ElseIf int����� = 2 Then
                    MsgBox "�Բ�������ҩƷ��治�㣬���ܳ�����" & vbCrLf & strMsg, vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End With
    End If
    
    With vsfList
        gstrSQL = "zl_ҩƷ�̵�_Strike('" & .TextMatrix(.Row, 0) & "','" & UserInfo.�û����� & "')"
        
        Call zlDataBase.ExecuteProcedure(gstrSQL, mstrTitle)
        
        '��ʾͣ��ҩƷ
        Call CheckStopMedi(���ݺ�.�̵�� & "|" & .TextMatrix(.Row, 0))
    End With
    StrikeSave = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    
    'MsgBox "����ʧ�ܣ�", vbInformation, gstrSysName
    Call SaveErrLog

End Function

Private Sub cbsEditDel()
    'ɾ��
    Dim strBillNo As String
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        strTitle = IIf(tbcDetail.Selected.Index = 0, "�̵��¼��", "�̵��")
        
        On Error GoTo ErrHandle
        intRow = .Row
        strBillNo = .TextMatrix(intRow, 0)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & strBillNo & "����" & strTitle & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            If tbcDetail.Selected.Index = 1 Then
                gstrSQL = "zl_ҩƷ�̵�_Delete('" & strBillNo & "')"
            Else
                gstrSQL = "zl_ҩƷ�̵��¼��_Delete('" & strBillNo & "')"
            End If
            Call zlDataBase.ExecuteProcedure(gstrSQL, mstrTitle)
            
            intRecord = intRecord - 1
            mlastRow = 0
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                With vsfDetail
                    .rows = 1
                    .rows = 2
                    .FixedRows = 1
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                End With
                
            End If
                
            '.RowHeight(intRow) = 0
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
            vsfList_EnterCell
        End If
    End With
    stbThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    cbsViewRefresh
    Exit Sub

ErrHandle:
    If ErrCenter() = 1 Then Resume 'Resume����������õ���
    Call SaveErrLog
End Sub

Private Sub cbsAffirmant()
    Dim str������� As String       'ȱʡ��Ϊȷ�ϼ�¼�Ľ�������
    '��д�¶�ȷ�ϼ�¼
    If tbcDetail.Selected.Index = 1 Then
        str������� = vsfList.TextMatrix(vsfList.Row, 5)
    End If
    With frm�¶�ȷ��
        Call .ShowEditor(Me.cboStock.ItemData(Me.cboStock.ListIndex), str�������)
    End With
End Sub

Private Sub cbsDisplay()
    '�鿴����
    
    Dim strNo As String
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        If tbcDetail.Selected.Index = 0 Then
            frmNewCheckCourseCard.ShowCard Me, strNo, 4
        Else
            frmNewCheckCard.ShowCard Me, strNo, 4, .TextMatrix(.Row, .Cols - 2)
        End If
    End With
End Sub

Private Sub cbsCheckTable()
    Dim blnSuccess As Boolean
    '���ܼ��
    frmSmartCheck.ShowME cboStock.ItemData(cboStock.ListIndex), Me, blnSuccess
    
    If blnSuccess Then cbsViewRefresh
End Sub

Private Sub cbsAddTableAuto()
    Dim strNo As String
    Dim blnSuccess As Boolean

    frmNewCheckCard.ShowCard Me, strNo, 1, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableTotal()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmNewCheckCard.ShowCard Me, strNo, 5, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableZero()
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    frmNewCheckCard.ShowCard Me, strNo, 6, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableHouseAll()
    Dim strNo As String
    Dim blnSuccess As Boolean

    frmNewCheckCard.ShowCard Me, strNo, 7, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsAddTableSpecial()
    Dim strNo As String
    Dim blnSuccess As Boolean

    frmNewCheckCard.ShowCard Me, strNo, 8, , blnSuccess
    
    If blnSuccess Then
        Call cbsViewRefresh
    End If
End Sub

Private Sub cbsVerify()
    '����
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmNewCheckCard.ShowCard Me, strNo, 3, .TextMatrix(.Row, .Cols - 2), blnSuccess
    End With
    
    If blnSuccess Then Call cbsViewRefresh
End Sub



Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Not mblnLoad Then Exit Sub
    If Not mblnBootUp Then Exit Sub

    '���ÿؼ����������
    Call Ȩ�޿���(Control)
    
    '���ò˵��͹��߰�ť�Ŀ�������

    Dim strVerify As String, blnVisible As Boolean
    
    blnVisible = (tbcDetail.Selected.Index = 1)
    If tbcDetail.Selected.Index = 1 Then
        strVerify = vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 8)
    Else
        strVerify = ""
    End If
    
    With vsfList
        .ToolTipText = ""
    
        Select Case Control.id
            Case mconMenu_File_Preview    'Ԥ��
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Case mconMenu_File_Print   '��ӡ
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Case mconMenu_File_BillPreview    '����Ԥ��
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    Control.Enabled = tbcDetail.Selected.Index = 1
                End If
            Case mconMenu_File_BillPrint    '���ݴ�ӡ
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    Control.Enabled = tbcDetail.Selected.Index = 1
                End If
            Case mconMenu_File_Excel    '�����Excel
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    Control.Enabled = True
                End If
            Case mconMenu_Edit_AddModify    '�޸�
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    'δ��˵�
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '��˵�
                        Control.Enabled = False
                    Else '2,3 ������
                        Control.Enabled = False
                    End If
                End If
            Case mconMenu_Edit_AddDel    'ɾ��
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    'δ��˵�
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '��˵�
                        Control.Enabled = False
                    Else '2,3 ������
                        Control.Enabled = False
                    End If
                End If
            Case mconMenu_Edit_AddVerify   '���
                Control.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "���")
                
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    'δ��˵�
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '��˵�
                        Control.Enabled = False
                    Else '2,3 ������
                        Control.Enabled = False
                    End If
                End If
            Case mconMenu_Edit_AddStrike   '����
                Control.Visible = blnVisible And zlStr.IsHavePrivs(mstrPrivs, "����")
                If Not zlStr.IsHavePrivs(mstrPrivs, "���") And Control.Visible Then Control.BeginGroup = True
                
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    'δ��˵�
                        Control.Enabled = False
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '��˵�
                        Control.Enabled = True
                    Else '2,3 ������
                        If .TextMatrix(.Row, .Cols - 2) Mod 3 = 0 Then
                            .ToolTipText = "�������ݵ�ԭ����"
                            Control.Enabled = True
                        ElseIf .TextMatrix(.Row, .Cols - 2) Mod 3 = 2 Then
                            .ToolTipText = "��������"
                            Control.Enabled = False
                        End If
                    End If
                End If
            Case mconMenu_Edit_AddDisplay    '�鿴����
                If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
                    Control.Enabled = False
                Else
                    If strVerify = "" Then    'δ��˵�
                        Control.Enabled = True
                    ElseIf .TextMatrix(.Row, .Cols - 2) = 1 Then    '��˵�
                        Control.Enabled = True
                    Else '2,3 ������
                        Control.Enabled = True
                    End If
                End If
                
        End Select
    End With
    
    
End Sub

Private Sub ChkҩƷ_Click()
    TxtҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
    CmdҩƷ.Enabled = IIf(ChkҩƷ.Value = 1, True, False)
End Sub

Private Sub Cmd����_Click()
    Call cbsDisplay
End Sub

Private Sub cmdȷ��_Click()
    Dim strFind As String
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    
    If cbo�����.Enabled Then
        If cboδ���.ListIndex = 0 And cbo�����.ListIndex = 0 Then
            MsgBox "�Բ��𣬱���ѡ��һ�ֵ�����ʾ��Ĭ����ʾ����δ��˵��ݣ�!", vbInformation, gstrSysName
            cboδ���.ListIndex = 1
            cboδ���.SetFocus
            Exit Sub
        ElseIf cboδ���.ListIndex <> 0 And cbo�����.ListIndex = 0 Then 'ֻ��δ��˵���
            strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
        ElseIf cboδ���.ListIndex = 0 And cbo�����.ListIndex <> 0 Then 'ֻ������˵���
            If chkStrike.Value = 1 Then '��������
                strFind = " AND  A.������� is not Null And A.������� Between [5] And [6] "
            Else
                strFind = " AND A.��¼״̬ = 1 And A.������� is not Null And A.������� Between [5] And [6] "
            End If
        Else '������˺�δ��˵���
            If chkStrike.Value = 1 Then  '��������
                strFind = " AND (( A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4]) or ( A.������� is not Null And A.������� Between [5] And [6])) "
            Else
                strFind = " AND (( A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4]) or (A.��¼״̬ = 1 And A.������� is not Null And A.������� Between [5] And [6])) "
            End If
        End If
    Else
        If cboδ���.ListIndex = 0 Then
            MsgBox "�Բ��𣬱���ѡ��һ�ֵ�����ʾ��Ĭ����ʾ����δ��˵��ݣ�!", vbInformation, gstrSysName
            cboδ���.ListIndex = 1
            cboδ���.SetFocus
            Exit Sub
        ElseIf cboδ���.ListIndex <> 0 Then 'ֻ��δ��˵���
            strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
        End If
    End If
    
    If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
        txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, intNO, lng�ⷿID)
    End If
    If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
        txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿID)
    End If

    If Me.txt��ʼNo <> "" And Me.txt����NO <> "" Then strFind = strFind & " And A.No >= [1] And A.No <=[2] "
    If Me.txt��ʼNo <> "" And Me.txt����NO = "" Then strFind = strFind & " And A.No >= [1] "
    If Me.txt��ʼNo = "" And Me.txt����NO <> "" Then strFind = strFind & " And A.No <= [2] "

    SQLCondition.strNO��ʼ = Me.txt��ʼNo
    SQLCondition.strNO���� = Me.txt����NO
    
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(dtp����ʱ��(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date���ʱ�俪ʼ = CDate(Format(dtp��ʼʱ��(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date���ʱ����� = CDate(Format(dtp����ʱ��(1), "yyyy-mm-dd") & " 23:59:59")
    
    If ChkҩƷ.Value = 1 Then
        strFind = strFind & " And A.ҩƷID + 0 =[7] "
    End If
    
    SQLCondition.lngҩƷ = Val(TxtҩƷ.Tag)
    
    If Me.Txt����� <> "" And Txt�����.Enabled Then strFind = strFind & " And A.����� like [10] "
    If Me.Txt������ <> "" Then strFind = strFind & " And A.������ like [9] "
    
    SQLCondition.str����� = Me.Txt����� & "%"
    SQLCondition.str������ = Me.Txt������ & "%"
    
    mstrFind = strFind
    
    GetList (mstrFind)  '�г�����ͷ
End Sub

Private Sub CmdҩƷ_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "ҩƷ�ƿ����", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , True)

    Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
    
End Sub

Private Sub cmd����_Click()
    cboStock.ListIndex = 0
    cboδ���.ListIndex = 1
    cbo�����.ListIndex = 0
    txt��ʼNo.Text = ""
    txt����NO.Text = ""
    Txt������.Text = ""
    Txt�����.Text = ""
    ChkҩƷ.Value = 0
    TxtҩƷ.Text = ""
    chkStrike.Value = 0
End Sub

Private Sub Form_Activate()
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
        vsfDetail.Row = 1
    End If
End Sub

Private Sub Form_Load()
    
    mblnLoad = False
    
    mintMaxMoneyBit = gtype_UserDrugDigits.Digit_���
    mbln����ģʽ = gtype_UserSysParms.P275_���۹���ģʽ <> 0
    
    InitComandBars
    InitTabControl
    loadCbo
    
    Me.dtp����ʱ��(1) = Sys.Currentdate
    Me.dtp��ʼʱ��(1) = DateAdd("d", -7, Me.dtp����ʱ��(1))
    
    Me.Caption = mstrTitle
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    Call zlDataBase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    stbThis.Panels(2).Picture = picColor
    
    Dim cbrMenuPop As CommandBarControl
    
    
    mblnLoad = True
End Sub

Private Sub loadCbo()
    '��ʼ��������
    Dim int��ѯ����  As Integer
    
    int��ѯ���� = Val(zlDataBase.GetPara("��ѯ����", glngSys, mlngMode, 7))
    int��ѯ���� = IIf(int��ѯ���� <> 1 And int��ѯ���� <> 7, 7, int��ѯ����)
    
    cboδ���.AddItem "0-����ʾ"
    cboδ���.AddItem "1-��ʾ����"
    cboδ���.AddItem "2-��ʾ7��֮��"
    cboδ���.AddItem "3-��ʾ����"
    cboδ���.AddItem "4-�Զ���"
    cboδ���.ListIndex = IIf(int��ѯ���� = 7, 2, 1)
    
    cbo�����.AddItem "0-����ʾ"
    cbo�����.AddItem "1-��ʾ����"
    cbo�����.AddItem "2-��ʾ7��֮��"
    cbo�����.AddItem "3-��ʾ����"
    cbo�����.AddItem "4-�Զ���"
    cbo�����.ListIndex = 0
    
End Sub

Private Sub InitTabControl()
    '��ʼ����ҳ�ؼ�
    
    With Me.tbcDetail
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        .InsertItem(mconTab_CheckCourseCard, "�̵��¼���嵥(&1)", Me.picMain.hWnd, 0).Tag = "�̵��¼���嵥(&1)_"
        .InsertItem(mconTab_CheckCard, "�̵���嵥(&2)", Me.picMain.hWnd, 0).Tag = "�̵���嵥(&2)_"
        
        .Item(.ItemCount - 1).Selected = True
        .Item(0).Selected = True
    End With
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 12900 Then Me.Width = 12900
    If Me.Height < 8000 Then Me.Height = 8000
    If picMain.Height < 4000 - picMain.Top Then picMain.Height = 4000 - picMain.Top
    
    
    fraCondition.Move 0, 900, Me.ScaleWidth, 1300
    cmdȷ��.Left = fraCondition.Width - cmdȷ��.Width - 100
    cmdȷ��.Top = dtp����ʱ��(1).Top - (cmdȷ��.Height - dtp����ʱ��(1).Height)
    cmd����.Left = cmdȷ��.Left - cmd����.Width - 50
    cmd����.Top = cmdȷ��.Top
    
    With tbcDetail
        .Top = fraCondition.Top + fraCondition.Height + 20
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - fraCondition.Top - fraCondition.Height - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    '״̬���Ƿ�ѡ
    Me.cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_View_StatusBar, , True).Checked = stbThis.Visible
    
    picMain.Move 0, 360, tbcDetail.Width, tbcDetail.Height - stbThis.Height
    
'    vsfList.Move 0, 0, picMain.Width, (picMain.Height - picSeparate_s.Height) / 2
    vsfList.Move 0, 0, picMain.Width
    
    With picSeparate_s
        .Left = 0
        .Top = vsfList.Top + vsfList.Height
        .Width = picMain.Width
    End With
    
    
    
'    vsfDetail.Move 0, picSeparate_s.Top + picSeparate_s.Height + 100, picMain.Width, (picMain.Height - picSeparate_s.Height) / 2 - 110
    If picSeparate_s.Top > picMain.Height - 2000 Then
        vsfList.Move 0, 0, picMain.Width, picMain.Height - (2100 + picSeparate_s.Height)
        picSeparate_s.Top = vsfList.Top + vsfList.Height
    End If
    
    With Cmd����
        .Left = picMain.Width - .Width - 100
        .Top = vsfList.Top + vsfList.Height + 30
    End With
    
    With vsfDetail
        .Left = 0
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Width = picMain.Width
        .Height = picMain.Height - .Top
    End With
    
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
    
End Sub


'******************************************************************************************************************
'���ܣ�
'������
'���أ�
'******************************************************************************************************************
Public Function DockPannelInit(ByRef dkpMain As DockingPane) As Boolean
On Error GoTo errH
    dkpMain.Options.ThemedFloatingFrames = True
    dkpMain.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMain.Options.AlphaDockingContext = True
    dkpMain.Options.CloseGroupOnButtonClick = True
    dkpMain.Options.HideClient = True
    DockPannelInit = True
    Exit Function
errH:
    MsgBox Err.Description, vbCritical, "ZLSOFT"
    Err.Clear
End Function



Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, mstrTitle
    Call SaveFlexState(vsfList, tbcDetail.Selected.Caption)
    Call SaveFlexState(vsfDetail, tbcDetail.Selected.Caption)
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        mshSelect.Visible = False
        Exit Sub
    End If
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Booker"
                    Txt������ = .TextMatrix(.Row, 2)
                    If tbcDetail.Selected.Index = mconTab_CheckCard Then
                        Txt�����.SetFocus
                    Else
                        cboδ���.SetFocus
                    End If
                Case "Verify"
                    Txt����� = .TextMatrix(.Row, 2)
                    cboδ���.SetFocus
            End Select
            .Visible = False
            Exit Sub
        End If
    End With
    
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub



Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    If vsfList.Height + y <= 1500 Then Exit Sub
    If vsfDetail.Height - y <= 1500 Then Exit Sub

    picSeparate_s.Move 0, picSeparate_s.Top + y
    Cmd����.Move Me.ScaleWidth - Cmd����.Width - 500, picSeparate_s.Top + 50
    vsfList.Move 0, 0, Me.ScaleWidth, vsfList.Height + y
    vsfDetail.Move 0, vsfList.Height + Cmd����.Height + 100, Me.ScaleWidth, vsfDetail.Height - y

    
'    With picSeparate_s
'        If .Top + picMain.Top + y < 2000 Then Exit Sub
'        If .Top + y > picMain.Height - 2000 Then Exit Sub
'        .Move .Left, .Top + y
'    End With
'
'    With vsfList
'        .Height = picSeparate_s.Top - .Top
'    End With
'
'    With Cmd����
'        .Top = vsfList.Top + vsfList.Height + 30
'    End With
'
'    With vsfDetail
'        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
'        .Height = picMain.Height - .Top
'    End With
End Sub


Public Sub ShowList(ByVal lngMode As Long, ByVal strTitle As String, ByVal FrmMain As Variant)
    Dim strFind As String
    Dim dateCurrentDate As Date
    Dim strTemp As String
    Dim int��ѯ���� As Integer
    
    mblnBootUp = False
    mlngMode = lngMode
    mstrTitle = strTitle
    mstrPrivs = gstrprivs
    Me.Caption = strTitle
    
    If Not CheckDepend Then Exit Sub            '���������Բ���
    
    mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    Call GetDrugDigit(mlng�ⷿID, mstrTitle, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        
    mstrVerifyStart = "1901-01-01"
    mstrVerifyEnd = "1901-01-01"
    
    dateCurrentDate = Sys.Currentdate
    mstrStart = Format(DateAdd("d", -6, dateCurrentDate), "yyyy-MM-dd") 'Ĭ����ȡ7������
    mstrEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    strFind = " AND A.��¼״̬ = 1 And A.������� is Null And A.�������� Between [3] And [4] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(mstrStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(mstrEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = strFind
    
    GetList (mstrFind)  '�г�����ͷ
    
    RestoreWinState Me, App.ProductName, mstrTitle
        
    mblnBootUp = True
    
    If IsObject(FrmMain) Then
        Me.Show , FrmMain
    Else
        OS.ShowChildWindow Me.hWnd, FrmMain
    End If
    Me.ZOrder 0
End Sub

'�������������
Private Function CheckDepend() As Boolean
    
    Dim rsDepend As New Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    CheckDepend = False
    
    strStock = "HIJKLMN"
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
             & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� is Null) And c.�������� = b.���� " _
              & "AND Instr([1],b.����,1) > 0 " _
             & " AND a.id = c.����id " _
              & "AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"

    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, strStock) '�鿴�Ƿ���ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ���
    
    If rsDepend.EOF Then
        MsgBox "����Ӧ������һ������ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
             & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = '" & gstrNodeNo & "' Or a.վ�� is Null) And c.�������� = b.���� " _
              & "AND Instr([1],b.����,1) > 0 " _
             & " AND a.id = c.����id " _
              & "AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
              & IIf(zlStr.IsHavePrivs(mstrPrivs, "���пⷿ"), "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[2])")

    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, strStock, UserInfo.�û�ID) '�鿴�û���������ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ���
            
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 And .ListCount > 0 Then .ListIndex = 0  'ȱʡ���Ų���ҩ�����ʣ�ҩ�����ʣ������Ƽ������ʵĲ�����Ĭ��ѡ���һ���������ʵĲ���
        
        If .ListIndex = -1 Then
            If Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ") Then
                MsgBox "�㲻�ǿⷿ������Ա�򲻾������пⷿ��Ȩ�ޣ����ܽ��룡", vbInformation, gstrSysName
                Unload Me
                Exit Function
            End If
            .ListIndex = 0
        End If
    End With

    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub Ȩ�޿���(ByVal Control As XtremeCommandBars.ICommandBarControl)
    'Ȩ�޿���
  
    Select Case Control.id
        Case mconMenu_Edit_AddBill, mconMenu_Edit_AddTable '��¼������¼��
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�Ǽ�")
            If mconMenu_Edit_AddTable = Control.id Then
                Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�̵��") And zlStr.IsHavePrivs(mstrPrivs, "�Ǽ�")
                tbcDetail.Item(mconTab_CheckCard).Visible = zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�̵��")
            End If
        Case mconMenu_Edit_AddModify '�޸�
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�޸�")
            '�ж��Ƿ�ӷֽ���
            If Not zlStr.IsHavePrivs(mstrPrivs, "�Ǽ�") And zlStr.IsHavePrivs(mstrPrivs, "�޸�") Then Control.BeginGroup = True
        Case mconMenu_Edit_AddDel  'ɾ��
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ɾ��")
            '�ж��Ƿ�ӷֽ���
            If Not zlStr.IsHavePrivs(mstrPrivs, "�Ǽ�") And Not zlStr.IsHavePrivs(mstrPrivs, "�޸�") And zlStr.IsHavePrivs(mstrPrivs, "ɾ��") Then Control.BeginGroup = True
        Case mconMenu_Edit_AddVerify  '���
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "���")
        Case mconMenu_Edit_AddStrike   '����
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "����")
            '�ж��Ƿ�ӷֽ���
            If Not zlStr.IsHavePrivs(mstrPrivs, "���") And zlStr.IsHavePrivs(mstrPrivs, "����") Then Control.BeginGroup = True
        Case mconMenu_Edit_AddAffirmant  'ȷ��
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "�¶�ȷ��")
        Case mconMenu_Edit_AddTableZero   'ȫ����Ϊ��
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "ȫ����Ϊ��")
        Case mconMenu_File_BillPrint    '���ݴ�ӡ
            Control.Visible = zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ")
    End Select
End Sub

Private Sub GetList(ByVal strFind As String)
    Dim rsList As New Recordset
    Dim strUserPart As String
    Dim str��װϵ�� As String
    Dim strSqlForm As String
    Dim n As Integer
    
    '����ͳ�ƺϼƽ��
    Dim dbl1 As Double
    Dim dbl2 As Double
    Dim dbl3 As Double
    Dim dbl�̵�ɱ���� As Double
    Dim dbl�̵���� As Double

    mlastRow = 0
    On Error GoTo ErrHandle

    Call FS.ShowFlash("��������ҩƷ��¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    strUserPart = " And A.�ⷿID+0=[11] "
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            str��װϵ�� = "1"
        Case mconint���ﵥλ
            str��װϵ�� = "B.�����װ"
        Case mconintסԺ��λ
            str��װϵ�� = "B.סԺ��װ"
        Case mconintҩ�ⵥλ
            str��װϵ�� = "B.ҩ���װ"
    End Select
    
    vsfList.Redraw = flexRDNone
    'Ƶ���ֶα���� �̵�ʱ��
    If tbcDetail.Selected.Index = 1 Then 'ѡ������̵���嵥
        If SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� = 0 Then
            strSqlForm = " , ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " And b.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� = "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ������ĿĿ¼ G"
            strFind = strFind & " And b.ҩ��id = g.Id And g.����id + 0=[12] and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " And b.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7') and g.����id + 0=[12]"
        End If
        
        gstrSQL = "Select NO, �̵�ʱ��, ������, ��������, �޸���, �޸�����, �����, �������, " & _
                "   to_char(Sum(�̵���), " & mstrMoneyFormat & ") �̵���, to_char(Sum(����), " & mstrMoneyFormat & ") ����,to_char(Sum(�������), " & mstrMoneyFormat & ") �������,to_char(Sum(�̵�ɱ����)," & mstrMoneyFormat & ") �̵�ɱ����, to_char(Sum(�ɱ�����)," & mstrMoneyFormat & ") �ɱ�����, ��¼״̬, ժҪ" & _
                " from ( SELECT a.no,a.���, Ƶ�� AS �̵�ʱ��," _
                & "a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������,a.�޸���,TO_CHAR (min(a.�޸�����), 'yyyy-mm-dd HH24:Mi:SS') AS �޸�����, a.�����," _
                & "TO_CHAR (min(a.�������), 'yyyy-mm-dd HH24:Mi:SS') AS �������, " _
                & "     LTrim(To_Char(to_char(A.���� /" & str��װϵ�� & "," & mstrNumberFormat & ") * TO_CHAR (a.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & ") , " & mstrMoneyFormat & ")) As �̵���," _
                & "ltrim(to_char(���۽��*a.���ϵ��,decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) ����," _
                & "ltrim(to_char(to_char((A.����-A.��д����) /" & str��װϵ�� & "," & mstrNumberFormat & ") * TO_CHAR (a.���ۼ�* Decode(��¼״̬, 1, 1, Decode(Mod(��¼״̬, 3), 0, 1, -1))*" & str��װϵ�� & ", " & mstrPriceFormat & "),decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," _
                & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS �������," _
                & "ltrim(to_char((a.�ɱ���+to_char(a.���۽��*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1)),decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat _
                & ") ", mstrMoneyFormat) & ")))-(a.�ɱ����+to_char(a.���*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1)),decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & ")))," & mstrMoneyFormat & ")) as �̵�ɱ����," _
                & "ltrim(to_char(a.���۽��*a.���ϵ��-a.���*a.���ϵ��,decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) as �ɱ�����," _
                & " a.��¼״̬, a.ժҪ " _
                & " FROM ҩƷ�շ���¼ a,ҩƷ��� B " & strSqlForm _
                & " Where A.ҩƷID=B.ҩƷID And A.���� = 12  " & strUserPart & strFind _
                & " Group By a.No,a.���, Ƶ��, a.������, a.�޸���, a.�����, a.�ɱ���, a.���ϵ��, a.�ɱ���,a.�ɱ����," & str��װϵ�� & ", a.���۽��, a.��¼״̬, a.����, a.��д����, a.���ۼ�,a.����, a.����, a.���, a.ժҪ,b.�Ƿ����۹���) " _
                & " Group By NO, �̵�ʱ��, ������, ��������, �޸���, �޸�����, �����, �������, ��¼״̬, ժҪ ORDER BY no DESC,�������� ASC"
    Else 'ѡ������̵��¼���嵥
        If SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� = 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in (select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� = "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.����id + 0=[12] and (g.���='5' or g.���='6' or g.���='7')"
        ElseIf SQLCondition.str���� <> "" And SQLCondition.lngҩƷ���� <> 0 Then
            strSqlForm = " , ҩƷ��� F, ������ĿĿ¼ G, ҩƷ���� H"
            strFind = strFind & " and a.ҩƷid = f.ҩƷid And f.ҩ��id = g.Id And g.Id = h.ҩ��id(+) and h.ҩƷ���� in(select * from Table(Cast(f_Str2list([13]) As zlTools.t_Strlist))) and (g.���='5' or g.���='6' or g.���='7') and g.����id + 0=[12]"
        End If
        gstrSQL = " SELECT a.no, Ƶ�� AS �̵�ʱ��," _
                    & "a.������,TO_CHAR (min(a.��������), 'yyyy-mm-dd HH24:Mi:SS') AS ��������,a.�޸���,TO_CHAR (min(a.�޸�����), 'yyyy-mm-dd HH24:Mi:SS') AS �޸�����,a.ժҪ " _
                    & " FROM ҩƷ�շ���¼ a " & strSqlForm _
                    & " Where a.���� = 14  " & strUserPart & strFind _
                    & " Group by a.no,Ƶ��,a.������,a.�޸���,a.ժҪ " _
                    & " ORDER BY no DESC,�������� ASC "
    End If
    
    Set rsList = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, _
        SQLCondition.strNO��ʼ, _
        SQLCondition.strNO����, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.date���ʱ�俪ʼ, _
        SQLCondition.date���ʱ�����, _
        SQLCondition.lngҩƷ, _
        SQLCondition.lng����ⷿ, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        cboStock.ItemData(cboStock.ListIndex), _
        SQLCondition.lngҩƷ����, _
        SQLCondition.str����)
    
    mbln�� = False
    Set vsfList.DataSource = rsList
    mbln�� = True
    
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .Redraw = flexRDDirect
            
            .TopRow = 1
            .rows = .rows - 99
        End If
    
        
        .ColAlignment(.ColIndex("�̵�ɱ����")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�̵���")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("����")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�������")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�ɱ�����")) = flexAlignRightCenter
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        If tbcDetail.Selected.Index = 1 Then 'ѡ������̵���嵥
            .colHidden(.Cols - 2) = True 'ʼ������"��¼״̬"��һ��
            .colHidden(.ColIndex("����")) = True 'Ĭ�ϲ���ʾ
            .colHidden(.ColIndex("�������")) = True 'Ĭ�ϲ���ʾ
            .colHidden(.ColIndex("�ɱ�����")) = True 'Ĭ�ϲ���ʾ
            
            vsfHidden vsfList
            
            lbl2.Visible = Not .colHidden(.ColIndex("����")) '�����ʾ�������ϼƲ���ʾ
            lbl3.Visible = Not .colHidden(.ColIndex("�������")) '��������ʾ�����������ϼƲ���ʾ
            lbl�ɱ�����.Visible = Not .colHidden(.ColIndex("�ɱ�����")) '�ɱ������ʾ����ɱ�����ϼƲ���ʾ
        End If
        
    
        
        For n = 0 To .Cols - 1
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    'ͳ�ƺϼƽ��
    lbl1.Caption = "�̵���ϼƣ�"
    lbl2.Caption = "����ϼƣ�"
    lbl3.Caption = "�������ϼƣ�"
    
    If tbcDetail.Selected.Index = 1 Then 'ѡ������̵���嵥
        lbl1.Visible = True
        If mblnViewCost = False Then
            lblSum�ɱ����.Visible = False
            lbl�ɱ�����.Visible = False
        Else
            lblSum�ɱ����.Visible = True
            lbl�ɱ�����.Visible = Not vsfList.colHidden(vsfList.ColIndex("�ɱ�����")) '�ɱ������ʾ����ɱ�����ϼƲ���ʾ
        End If
        If (Not rsList.EOF) And (Not rsList.BOF) Then
            rsList.MoveFirst
            Do While Not rsList.EOF
                dbl1 = dbl1 + IIf(IsNull(rsList!�̵���), 0, rsList!�̵���)
                dbl2 = dbl2 + IIf(IsNull(rsList!����), 0, rsList!����)
                dbl3 = dbl3 + IIf(IsNull(rsList!�������), 0, rsList!�������)
                dbl�̵�ɱ���� = dbl�̵�ɱ���� + IIf(IsNull(rsList!�̵�ɱ����), 0, rsList!�̵�ɱ����)
                dbl�̵���� = dbl�̵���� + IIf(IsNull(rsList!�ɱ�����), 0, rsList!�ɱ�����)
                rsList.MoveNext
            Loop
            rsList.MoveFirst
            
            lbl1.Caption = "�̵���ϼƣ�" & Format(dbl1, "0." & String(mintShowMoneyDigit, "0"))
            lbl2.Caption = "����ϼƣ�" & Format(dbl2, "0." & String(mintShowMoneyDigit, "0"))
            lbl3.Caption = "�������ϼƣ�" & Format(dbl3, "0." & String(mintShowMoneyDigit, "0"))
            lblSum�ɱ����.Caption = "�̵�ɱ����ϼƣ�" & Format(dbl�̵�ɱ����, "0." & String(mintShowMoneyDigit, "0"))
            lbl�ɱ�����.Caption = "�ɱ����" & Format(dbl�̵����, "0." & String(mintShowMoneyDigit, "0"))
        End If
    Else
        lblSum�ɱ����.Visible = False
        lbl�ɱ�����.Visible = False
        lbl1.Visible = False
        lbl2.Visible = False
        lbl3.Visible = False
    End If
    
    lbl2.Left = lbl1.Width + lbl1.Left + 200
    lbl3.Left = IIf(lbl2.Visible, lbl2.Width + lbl2.Left + 200, lbl2.Left)
    lblSum�ɱ����.Left = IIf(lbl3.Visible, lbl3.Width + lbl3.Left + 200, lbl3.Left)
    lbl�ɱ�����.Left = lblSum�ɱ����.Width + lblSum�ɱ����.Left + 200
    
    vsfList_EnterCell    '�г�������
    
    SetStrikeColor
    With vsfList
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
    vsfList.Redraw = flexRDDirect
    Call FS.StopFlash
    Screen.MousePointer = vbDefault
    stbThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    If vsfList.Visible = True Then
        vsfList.SetFocus
        vsfList.Row = 1
    End If
    
    Cmd����.Enabled = Not (vsfList.TextMatrix(vsfList.Row, 0) = "" Or vsfList.Row = 0)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        If tbcDetail.Selected.Index = 1 Then 'ѡ������̵���嵥
            If mblnBootUp = False Then
                For intCol = 1 To .Cols - 1
                    If intCol = 1 Then
                        .ColWidth(intCol) = 2000
                    ElseIf intCol = .Cols - 2 Then
                        .ColWidth(intCol) = 0
                    Else
                        .ColWidth(intCol) = 1000
                    End If
                Next
            End If
        Else
            If mblnBootUp = False Then
                .ColWidth(1) = 2000
                .ColWidth(4) = 3000
            End If
        End If
        .ColWidth(.ColIndex("�̵�ɱ����")) = 1500
    End With
    
    Call RestoreFlexState(vsfList, tbcDetail.Selected.Caption)
    If tbcDetail.Selected.Index = 1 And mblnViewCost = False Then
        vsfList.colHidden(vsfList.ColIndex("�̵�ɱ����")) = True
        vsfList.colHidden(vsfList.ColIndex("�ɱ�����")) = True
    End If
End Sub



Private Sub tbcDetail_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '��¼��û�����
    cbo�����.Enabled = tbcDetail.Selected.Index = mconTab_CheckCard
    cbo�����_Click
    
    Txt�����.Enabled = tbcDetail.Selected.Index = mconTab_CheckCard
    If Txt�����.Enabled Then
        Txt�����.BackColor = &H80000005
    Else
        Txt�����.BackColor = &H8000000F
    End If
    
    If Not mblnLoad Then Exit Sub

    Call SaveFlexState(vsfList, tbcDetail.Item(mintLastIndex).Caption)
    Call SaveFlexState(vsfDetail, tbcDetail.Item(mintLastIndex).Caption)
    
    mblnBootUp = False
    If Item.Index = 1 Then
        vsfDetail.ToolTipText = mcstComment
    Else
        vsfDetail.ToolTipText = ""
    End If
    
    If cbo�����.Enabled Then
        If cboδ���.ListIndex = 0 And cbo�����.ListIndex = 0 Then
            MsgBox "�Բ��𣬱���ѡ��һ�ֵ�����ʾ��Ĭ����ʾ����δ��˵��ݣ�!", vbInformation, gstrSysName
            cboδ���.ListIndex = 1
        End If
    Else
        If cboδ���.ListIndex = 0 Then
            MsgBox "�Բ��𣬱���ѡ��һ�ֵ�����ʾ��Ĭ����ʾ����δ��˵��ݣ�!", vbInformation, gstrSysName
            cboδ���.ListIndex = 1
        End If
    End If
    
    cmdȷ��_Click   '�г�����ͷ
    
    mintLastIndex = Item.Index
    
    mblnBootUp = True
End Sub


Private Sub txt����NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If tbcDetail.Selected.Index = 1 Then
            '�̵��
            intNO = 29
        Else
            '�̵��¼��
            intNO = 62
        End If
    End If
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt����NO) < 8 And Len(txt����NO) > 0 Then
            txt����NO.Text = zlCommFun.GetFullNO(txt����NO.Text, intNO, lng�ⷿID)
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub txt����NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt��ʼNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    
    '��ʼ׼��
    intNO = Switch(mlngMode = 1303, 25, mlngMode = 1304, 26, mlngMode = 1305, 27, mlngMode = 1306, 28, mlngMode = 1307, 29)
    If mlngMode = 1307 Then
        If tbcDetail.Selected.Index = 1 Then
            '�̵��
            intNO = 29
        Else
            '�̵��¼��
            intNO = 62
        End If
    End If
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    
    If KeyCode = vbKeyReturn Then
        If Len(txt��ʼNo) < 8 And Len(txt��ʼNo) > 0 Then
            txt��ʼNo.Text = zlCommFun.GetFullNO(txt��ʼNo.Text, intNO, lng�ⷿID)
        End If
        Me.txt����NO.SetFocus
    End If
End Sub

Private Sub txt��ʼNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt�����_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then cmdȷ��.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt�����.Text) = "" Then
            SendKeys vbTab
            Exit Sub
        End If
        Txt�����.Text = UCase(Txt�����.Text)

        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ�����]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt����� & "%", _
                        Me.Txt����� & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt�����.SelStart = 0
                Txt�����.SelLength = Len(Txt�����.Text)
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = Txt������.Top + fraCondition.Top + Txt������.Height
                    .Left = Txt������.Left + fraCondition.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt����� = IIf(IsNull(!����), "", !����)
                SendKeys vbTab
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt�����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt������_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt�����.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt������.Text) = "" Then
            If tbcDetail.Selected.Index = mconTab_CheckCard Then
                Txt�����.SetFocus
            Else
                SendKeys vbTab
            End If
            
            Exit Sub
        End If
        Txt������.Text = UCase(Txt������.Text)

        gstrSQL = "Select ���,����,���� From ��Ա�� " & _
                  "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(���) like [1] or Upper(����) like [2]) " & _
                  "  And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[ȡ������]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt������ & "%", _
                        Me.Txt������ & "%", gstrNodeNo)

        With rsTemp
            If .EOF Then
                MsgBox "����ֵ��Ч��", vbInformation, gstrSysName
                KeyCode = 0
                Txt������.SelStart = 0
                Txt������.SelLength = Len(Txt������.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = Txt������.Top + fraCondition.Top + Txt������.Height
                    .Left = Txt������.Left + fraCondition.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt������ = IIf(IsNull(!����), "", !����)
                If tbcDetail.Selected.Index = mconTab_CheckCard Then
                    Me.Txt�����.SetFocus
                Else
                    SendKeys vbTab
                End If
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub TxtҩƷ_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(TxtҩƷ.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fraCondition.Left + TxtҩƷ.Left
    sngTop = Me.Top + fraCondition.Top + TxtҩƷ.Top + TxtҩƷ.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - TxtҩƷ.Height - 3630
    End If
    
    strkey = Trim(TxtҩƷ.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "ҩƷ�ƿ����", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , True)

'    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gintҩƷ������ʾ = 1 Then
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & IIf(IsNull(RecReturn!��Ʒ��), RecReturn!ͨ����, RecReturn!��Ʒ��)
    Else
        TxtҩƷ.Text = "[" & RecReturn!ҩƷ���� & "]" & RecReturn!ͨ����
    End If
    TxtҩƷ.Tag = RecReturn!ҩƷid
 
    
End Sub

Private Sub TxtҩƷ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub vsfDetail_EnterCell()
    With vsfDetail
        If .Row = 0 Then Exit Sub
        
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub


Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub


Private Sub vsfList_DblClick()
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Visible Then Exit Sub
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Enabled Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    cbsEditModify
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Visible Then Exit Sub
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_AddModify, , True).Enabled Then Exit Sub
        cbsEditModify
    End If
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfList_EnterCell()
    Dim rsDetail As New Recordset
    Dim intBill As Integer                      '��������  �磺1���⹺��⣻2��
    Dim str��װϵ�� As String
    Dim str��λ�ֶ� As String
    Dim strOrder As String
    Dim strCompare As String
    Dim strSqlЧ�� As String
    Dim lngColor As Long
    Dim n As Long
    Dim i As Integer
    Dim intCol As Integer
    Dim strSqlҩ�� As String
    Dim strSqlOrder As String
    
    If Not mbln�� Then Exit Sub
    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo ErrHandle
    With vsfList
        .Redraw = flexRDNone
        .ForeColorSel = .Cell(flexcpForeColor, mlastRow, 1)
        .Redraw = flexRDDirect
    End With
    
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.ҩƷ�̵�)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        '���������
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        '����������
        strSqlOrder = "ҩƷ��Ϣ"
    ElseIf strCompare = "2" Then
        '����������
        strSqlOrder = "Substr(ҩƷ��Ϣ, Instr(ҩƷ��Ϣ, ']') + 1)"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",ҩƷ��Ϣ,���"
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
        
        vsfDetail.Redraw = flexRDNone
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                str��װϵ�� = "1"
                str��λ�ֶ� = "I.���㵥λ"
            Case mconint���ﵥλ
                str��װϵ�� = "B.�����װ"
                str��λ�ֶ� = "B.���ﵥλ"
            Case mconintסԺ��λ
                str��װϵ�� = "B.סԺ��װ"
                str��λ�ֶ� = "B.סԺ��λ"
            Case mconintҩ�ⵥλ
                str��װϵ�� = "B.ҩ���װ"
                str��λ�ֶ� = "B.ҩ�ⵥλ"
        End Select
        
        strSqlЧ�� = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "TO_CHAR(A.Ч��-1,'YYYY-MM-DD') AS ��Ч����", "TO_CHAR(A.Ч��,'YYYY-MM-DD') AS ʧЧ��")
        
        If gintҩƷ������ʾ = 0 Then
            strSqlҩ�� = ",('['||I.����||']'||I.����) AS ҩƷ��Ϣ"
        ElseIf gintҩƷ������ʾ = 1 Then
            strSqlҩ�� = ",('['||I.����||']'||NVL(N.����,I.����)) AS ҩƷ��Ϣ"
        Else
            strSqlҩ�� = ",('['||I.����||']'||I.����) AS ҩƷ��Ϣ,N.���� As ��Ʒ��"
        End If
        
        intBill = IIf(tbcDetail.Selected.Index = 1, 12, 14)
        If tbcDetail.Selected.Index = 1 Then 'ѡ������̵���嵥
            gstrSQL = "Select DISTINCT a.���" & strSqlҩ�� & "," _
                    & "     B.ҩƷ��Դ,B.����ҩ��,I.���,a.���� as ������,a.ԭ����," & str��λ�ֶ� & " as ��λ,a.����," & strSqlЧ�� & ",a.��׼�ĺ�," _
                    & "     LTRIM(to_char(A.��д���� /" & str��װϵ�� & ",decode(a.����,0,'999999999990.00000'," & mstrNumberFormat & "))) AS ������," _
                    & "     LTRIM(to_char(A.���� /" & str��װϵ�� & "," & mstrNumberFormat & ")) AS ʵ����," _
                    & "     Decode(Sign(A.����-A.��д����),-1,'��',1,'ӯ','ƽ') as ��־," _
                    & "     LTRIM(to_char(A.ʵ������ /" & str��װϵ�� & ",decode(a.����,0,'999999999990.00000'," & mstrNumberFormat & "))) AS ������," _
                    & "     LTRIM(TO_CHAR (a.����*" & str��װϵ�� & ", " & mstrCostFormat & ")) AS �ɱ���," _
                    & "     LTRIM(TO_CHAR (a.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & ")) AS �ۼ�," _
                    & "     LTRIM(TO_CHAR (a.���۽��*a.���ϵ��,decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS ����," _
                    & "     LTRIM(TO_CHAR (to_char((A.����-A.��д����) /" & str��װϵ�� & "," & mstrNumberFormat & ") * TO_CHAR (a.���ۼ�* Decode(��¼״̬, 1, 1, Decode(Mod(��¼״̬, 3), 0, 1, -1))*" & str��װϵ�� & ", " & mstrPriceFormat & "),decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," _
                    & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS �������," _
                    & "     LTRIM(TO_CHAR (a.���*a.���ϵ��, decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) AS ��۲�, " _
                    & "     LTrim(To_Char(to_char(A.���� /" & str��װϵ�� & "," & mstrNumberFormat & ")*TO_CHAR (a.���ۼ�*" & str��װϵ�� & ", " & mstrPriceFormat & "), " & mstrMoneyFormat & ")) As �̵���," _
                    & "     LTrim(To_Char(((a.�ɱ���+to_char(a.���۽��*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1)),decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) _
                    & "     )))-(a.�ɱ����+to_char(a.���*a.���ϵ��*Decode(a.��¼״̬, 1, 1, Decode(Mod(a.��¼״̬, 3), 0, 1, -1)),decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))))," & mstrMoneyFormat & ")) as �̵�ɱ����, " _
                    & "     ltrim(To_Char((a.���۽��*a.���ϵ�� - a.���*a.���ϵ�� ), decode(nvl(a.����,0),0," & mstrMaxMoneyFormat & "," & IIf(mbln����ģʽ, " decode(nvl(b.�Ƿ����۹���,0),1,decode(a.����-a.���ۼ�,0," & mstrMaxMoneyFormat & "," & mstrMoneyFormat & ")," & mstrMoneyFormat & ") ", mstrMoneyFormat) & "))) As �ɱ�����," _
                    & " Nvl(I.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) As ����ʱ��,e.�ⷿ��λ " _
                    & " From (Select a.���ϵ��,a.��¼״̬,a.���,a.ҩƷid,a.����,a.ԭ����,a.����,a.Ч��,A.��д����,A.����,A.ʵ������,a.�ɱ���,a.�ɱ����,a.���ۼ�,a.���۽��,a.���,a.����,a.��׼�ĺ�,a.�ⷿid" _
                    & "         From ҩƷ�շ���¼ a" _
                    & "        Where a.��¼״̬= [2] And a.����= 12 And a.No=[1]) a," _
                    & "        ҩƷ��� b,�շ���ĿĿ¼ I ,�շ���Ŀ���� n,ҩƷ�����޶� e" _
                    & " Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id" _
                    & "        And a.ҩƷid=n.�շ�ϸĿid(+) And n.����(+)=3 " _
                    & "        And a.ҩƷid = e.ҩƷid(+) and a.�ⷿid = e.�ⷿid(+) " _
                    & " ORDER BY " & strSqlOrder
        Else
            gstrSQL = "Select DISTINCT a.���" & strSqlҩ�� & "," _
                    & "     B.ҩƷ��Դ,B.����ҩ��,I.���,a.���� as ������,a.ԭ����," & str��λ�ֶ� & " as ��λ,a.����," & strSqlЧ�� & ",a.��׼�ĺ�," _
                    & "     to_char(A.���� /" & str��װϵ�� & "," & mstrNumberFormat & ") AS ʵ����" _
                    & " From (Select a.���,a.ҩƷid,a.����,a.ԭ����,a.����,a.Ч��,A.��д����,A.����,A.ʵ������,a.���ۼ�,a.���۽��,a.���,a.��׼�ĺ�" _
                    & "         From ҩƷ�շ���¼ a" _
                    & "        Where a.��¼״̬= 1 And a.����= 14 And a.No=[1]) a," _
                    & "        ҩƷ��� b,�շ���ĿĿ¼ I ,�շ���Ŀ���� n" _
                    & " Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id" _
                    & "        And a.ҩƷid=n.�շ�ϸĿid(+) And n.����(+)=3 " _
                    & " ORDER BY " & strSqlOrder
        End If
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, mstrTitle, vsfList.TextMatrix(vsfList.Row, 0), vsfList.TextMatrix(vsfList.Row, vsfList.Cols - 2))
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close
        
        With vsfDetail
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
        End With

        
        vsfDetail.Redraw = flexRDDirect
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Cols = IIf(tbcDetail.Selected.Index = 1, 25, 12)
            If gintҩƷ������ʾ = 2 Then .Cols = .Cols + 1
            .rows = 2
            .Clear
            
            intCol = 0
            
            .TextMatrix(0, intCol) = "���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ҩƷ��Ϣ": intCol = intCol + 1
            
            If gintҩƷ������ʾ = 2 Then
                .TextMatrix(0, intCol) = "��Ʒ��": intCol = intCol + 1
            End If
            
            .TextMatrix(0, intCol) = "ҩƷ��Դ": intCol = intCol + 1
            .TextMatrix(0, intCol) = "����ҩ��": intCol = intCol + 1
            .TextMatrix(0, intCol) = "���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ԭ����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
            .TextMatrix(0, intCol) = "����": intCol = intCol + 1
            .TextMatrix(0, intCol) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��"): intCol = intCol + 1
            .TextMatrix(0, intCol) = "��׼�ĺ�": intCol = intCol + 1
            If tbcDetail.Selected.Index = 0 Then
                .TextMatrix(0, intCol) = "ʵ����": intCol = intCol + 1
            Else
                .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                .TextMatrix(0, intCol) = "ʵ����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "��־": intCol = intCol + 1
                .TextMatrix(0, intCol) = "������": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
                .TextMatrix(0, intCol) = "����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "��۲�": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�̵�ɱ����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�̵���": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ɱ�����": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�������": intCol = intCol + 1
                .TextMatrix(0, intCol) = "����ʱ��": intCol = intCol + 1
                .TextMatrix(0, intCol) = "�ⷿ��λ": intCol = intCol + 1
            End If
            
            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
        End With
    End If
    
    With vsfDetail
        .colHidden(.ColIndex("ҩƷ��Դ")) = True  'Ĭ�ϲ���ʾ
        .colHidden(.ColIndex("����ҩ��")) = True  'Ĭ�ϲ���ʾ
        .colHidden(.ColIndex("��׼�ĺ�")) = True  'Ĭ�ϲ���ʾ
        If tbcDetail.Selected.Index = 1 Then 'ֻ�����̵�����ϸ��Ĭ�ϲ���ʾ
            .colHidden(.ColIndex("����")) = True  'Ĭ�ϲ���ʾ
            .colHidden(.ColIndex("��۲�")) = True  'Ĭ�ϲ���ʾ
            .colHidden(.ColIndex("�ɱ�����")) = True  'Ĭ�ϲ���ʾ
            .colHidden(.ColIndex("�������")) = True  'Ĭ�ϲ���ʾ
            .colHidden(.ColIndex("�ⷿ��λ")) = True  'Ĭ�ϲ���ʾ
        End If
    End With
    
    vsfHidden vsfDetail
    SetDetailColWidth
    
    '��ɫ
    If tbcDetail.Selected.Index = 1 Then
        With vsfDetail
            .Redraw = flexRDNone
            For n = 1 To .rows - 1
                If .TextMatrix(n, 0) <> "" Then
                    If .TextMatrix(n, .ColIndex("��־")) = "ӯ" Then
                        lngColor = vbRed
                    ElseIf .TextMatrix(n, .ColIndex("��־")) = "��" Then
                        lngColor = vbBlue
                    Else
                        lngColor = vbBlack
                    End If
                    
                    '�̿���ӯ������ɫ���֣�
                    If lngColor <> vbBlack Then
                        .Cell(flexcpForeColor, n, 0, n, .Cols - 1) = lngColor
                    End If
                    
                    '�����ͣ��ҩƷ�����д�����ʾ
                    If Format(.TextMatrix(n, .ColIndex("����ʱ��")), "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, n, 0, n, .Cols - 1) = True
                    End If
                End If
            Next
            .Redraw = flexRDDirect
        End With
    End If
    
    vsfDetail.Row = 1
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDetailColWidth()
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo ErrHandle
    
    With vsfDetail
        .ColAlignment(.ColIndex("��λ")) = flexAlignCenterCenter    '��λ
        .ColAlignment(.ColIndex("ʵ����")) = flexAlignRightCenter 'ʵ����
        If tbcDetail.Selected.Index = 1 Then
            .ColAlignment(.ColIndex("������")) = flexAlignRightCenter     '������
            .ColAlignment(.ColIndex("��־")) = flexAlignCenterCenter    '��־
            .ColAlignment(.ColIndex("������")) = flexAlignRightCenter     '������
            .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter    '�ɱ���
            .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter    '�ۼ�
            .ColAlignment(.ColIndex("����")) = flexAlignRightCenter    '����
            .ColAlignment(.ColIndex("��۲�")) = flexAlignRightCenter    '��۲�
            .ColAlignment(.ColIndex("�̵���")) = flexAlignRightCenter    '�̵���
            .ColAlignment(.ColIndex("�������")) = flexAlignRightCenter    '�������
            .ColAlignment(.ColIndex("�̵�ɱ����")) = flexAlignRightCenter    '�̵�ɱ����
            .ColAlignment(.ColIndex("�ɱ�����")) = flexAlignRightCenter    '�ɱ�����
            
        End If
        
        If tbcDetail.Selected.Index = 1 Then
            If mblnBootUp = False Then
                .ColWidth(0) = 500
                .ColWidth(.ColIndex("ҩƷ��Ϣ")) = 2500
                For intCol = 2 To .Cols - 1
                    .ColWidth(intCol) = 1000
                Next
                .ColWidth(.ColIndex("����ʱ��")) = 0
                .ColWidth(.ColIndex("�̵�ɱ����")) = 1500
            End If
        Else
            .ColWidth(0) = 500
            .ColWidth(.ColIndex("ҩƷ��Ϣ")) = 2500
            For intCol = 2 To .Cols - 1
                .ColWidth(intCol) = 1000
            Next
        End If
        
        Call RestoreFlexState(vsfDetail, tbcDetail.Selected.Caption)
        If tbcDetail.Selected.Index = 1 And mblnViewCost = False Then
            .colHidden(.ColIndex("�ɱ���")) = True
            .colHidden(.ColIndex("��۲�")) = True
            .colHidden(.ColIndex("�̵�ɱ����")) = True
            .colHidden(.ColIndex("�ɱ�����")) = True
        End If
        
        str�ⷿ���� = ""
        gstrSQL = "Select a.�������� From ��������˵�� A Where a.����id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�ж��ǿⷿ����", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
            rsDetail.MoveNext
        Loop
        If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
        If bln��ҩ�ⷿ Then
            .colHidden(.ColIndex("ԭ����")) = False
        Else
            .colHidden(.ColIndex("ԭ����")) = True
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetStrikeColor()
    Dim intStatus As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    
    With vsfList
        If .rows <= 2 Then Exit Sub
        For intRow = 1 To .rows - 1
            intStatus = IIf(tbcDetail.Selected.Index = 0, 1, Val(.TextMatrix(intRow, .Cols - 2)))
            If intStatus Mod 3 = 0 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &H80000001
            End If
            If intStatus Mod 3 = 2 Then
                .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = &HFF
            End If
        Next
    End With
End Sub


Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(mstrStart, "yyyy-mm-dd") = "1901-01-01" And Format(mstrVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(mstrVerifyStart, "yyyy��MM��dd��") & "��" & Format(mstrVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(mstrVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(mstrStart, "yyyy��MM��dd��") & "��" & Format(mstrEnd, "yyyy��MM��dd��") & "  ������� " & Format(mstrVerifyStart, "yyyy��MM��dd��") & "��" & Format(mstrVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(mstrStart, "yyyy��MM��dd��") & "��" & Format(mstrEnd, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    objRow.Add "ʱ�䣺" & strRange
    objRow.Add "���ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
        
    objRow.Add "��ӡ��:" & UserInfo.�û�����
    objRow.Add "��ӡ����:" & Format(Sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfList
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub subExcel(bytMode As Byte)
'����:���������EXCEL
'����:bytMode3 �����EXCEL

    Dim objPrint As Object
    Dim objRow As zlTabAppRow
    Dim strRange As String
    
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = mstrTitle
        
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objRow.Add "NO." & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "NO")))
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "�̵�ⷿ��" & Trim(cboStock.Text)
    objRow.Add "�̵�ʱ�䣺" & Trim(vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�̵�ʱ��")))
    objPrint.UnderAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "ժҪ:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "ժҪ"))
    objPrint.BelowAppRows.Add objRow
        
    Set objRow = New zlTabAppRow
    objRow.Add "������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "������")) & "  ��������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "��������"))
    
    If tbcDetail.Selected.Index = 1 Then
        objRow.Add "�����:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�����")) & "  �������:" & vsfList.TextMatrix(vsfList.Row, GetCol(vsfList, "�������"))
        objPrint.BelowAppRows.Add objRow
    End If
    
    Set objPrint.Body = vsfDetail
    zlPrintOrView1Grd objPrint, bytMode
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    
    If Button <> 2 Then Exit Sub
    
    If Not cbsThis.ActiveMenuBar.Controls.Find(xtpControlPopup, mconMenu_EditPopup, , True).Visible Then Exit Sub '�༭���ɼ��˳�
    
    Set objPopup = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_EditPopup)
    If Not objPopup Is Nothing Then
        objPopup.CommandBar.ShowPopup
    End If
    
End Sub

'���ܣ���vsf�����ڵ��в�������������δ��ѡ���ص��н�������
Private Sub vsfHidden(ByRef objVSF As Object)
    Dim strColsName As String
    Dim strColName() As String
    Dim i As Integer
    Dim n As Integer
    Dim strDefaultColsName As String 'Ĭ�ϵ���
    Dim strTempColName As String
    
    strDefaultColsName = ":ҩƷ��Դ,0:����ҩ��,0:�ⷿ��λ,0:��׼�ĺ�,0:����,0:��۲�,0:�̵�ɱ�����,0:�������,0:�ɱ�����,0:��ǰ���,1:" '���п������ص���
    
    With objVSF
        strColsName = zlDataBase.GetPara("������", glngSys, mlngMode, "")
        
        '���ݴ���
        If strColsName = "" Then 'δ��ȡ����������Ϣ
            strColsName = strDefaultColsName
        Else
            '�ж���ȡ������Ĭ���и�������һ����ȡĬ�ϵ�
            If UBound(Split(strColsName, ":")) <> UBound(Split(strDefaultColsName, ":")) Then strColsName = strDefaultColsName
            
            '�ж���ȡ�������Ƿ���Ĭ�ϵ�һ�£���һ��ȡĬ�ϵ�
            For i = LBound(Split(strColsName, ":")) + 1 To UBound(Split(strColsName, ":")) - 1
                strTempColName = Split(Split(strColsName, ":")(i), ",")(0) '��ȡ��������
                
                If InStr(1, strDefaultColsName, ":" & strTempColName) = 0 Then '������������Ĭ��������
                    strColsName = strDefaultColsName
                    Exit For
                End If
            Next
            
        End If
        
        strColName = Split(strColsName, ":") '��ʽ:C,1
        
        For i = 0 To .Cols - 1
            '�жϽ����Ӧ���Ƿ��ǿ�������
            If InStr(1, strColsName, ":" & .TextMatrix(0, i)) > 0 Then
                For n = LBound(strColName) + 1 To UBound(strColName) - 1
                    If Split(strColName(n), ",")(0) = .TextMatrix(0, i) Then .colHidden(i) = Split(strColName(n), ",")(1) <> 1
                Next
            End If
             
        Next
    End With
End Sub
