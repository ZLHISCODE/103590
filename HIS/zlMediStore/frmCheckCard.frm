VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckCard 
   Caption         =   "ҩƷ�̵��"
   ClientHeight    =   6765
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11670
   Icon            =   "frmCheckCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11670
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdZero 
      Caption         =   "ʵ��������(&N)"
      Height          =   350
      Left            =   9840
      TabIndex        =   32
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "�̵㵽�������(&S)"
      Height          =   350
      Left            =   7800
      TabIndex        =   31
      Top             =   5760
      Width           =   1695
   End
   Begin VB.CommandButton cmdBatch 
      Caption         =   "�������ȡ(&G)"
      Height          =   350
      Left            =   4680
      TabIndex        =   29
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdReSet 
      Caption         =   "����(&R)"
      Height          =   350
      Left            =   6480
      TabIndex        =   28
      Top             =   5760
      Width           =   1100
   End
   Begin VB.CommandButton cmd�̶��� 
      Caption         =   "�̶���(&L)"
      Height          =   350
      Left            =   6090
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   2400
      TabIndex        =   6
      Top             =   5100
      Width           =   1815
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7440
      TabIndex        =   3
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8730
      TabIndex        =   4
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   7
      Top             =   0
      Width           =   11715
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   27
         Top             =   950
         Width           =   11235
         _cx             =   19817
         _cy             =   4948
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764622
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   315
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCheckCard.frx":014A
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
         ExplorerBar     =   5
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
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   2
         Top             =   4080
         Width           =   10410
      End
      Begin VB.Label lblCostPrice 
         AutoSize        =   -1  'True
         Caption         =   "�̵�ɱ����ϼƣ�"
         Height          =   180
         Left            =   4080
         TabIndex        =   30
         Top             =   3840
         Width           =   1620
      End
      Begin VB.Label lblCheckSum 
         AutoSize        =   -1  'True
         Caption         =   "�̵���ϼƣ�"
         Height          =   180
         Left            =   1920
         TabIndex        =   25
         Top             =   3840
         Width           =   1260
      End
      Begin VB.Label lblCheckDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ʱ��"
         Height          =   180
         Left            =   8640
         TabIndex        =   23
         Top             =   660
         Width           =   720
      End
      Begin VB.Label txtCheckDate 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9600
         TabIndex        =   22
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label txtStock 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1080
         TabIndex        =   21
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "����ϼƣ�"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   3840
         Width           =   1080
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   18
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   17
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   16
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   15
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   14
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   13
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   1
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ�̵��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   12
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�̵�ⷿ"
         Height          =   180
         Left            =   270
         TabIndex        =   0
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   11
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   10
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   9
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   8
         Top             =   4500
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":01BF
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":03D9
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":05F3
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":080D
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0A27
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0C41
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":0E5B
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1075
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":128F
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":14A9
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":16C3
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":18DD
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1AF7
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1D11
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":1F2B
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckCard.frx":2145
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   6405
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCheckCard.frx":235F
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14235
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCard.frx":2BF3
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmCheckCard.frx":30F5
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VB.Label lblCode 
      Caption         =   "����ҩƷ"
      Height          =   180
      Left            =   1530
      TabIndex        =   19
      Top             =   5145
      Width           =   720
   End
   Begin VB.Menu PopMenu 
      Caption         =   "�̶���"
      Visible         =   0   'False
      Begin VB.Menu mnuFirst 
         Caption         =   "��ҩ������λ��(&1)"
      End
      Begin VB.Menu mnuSecond 
         Caption         =   "��ҩ����Ч����(&2)"
      End
      Begin VB.Menu mnuSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDefault 
         Caption         =   "�ָ�(&D)"
      End
   End
   Begin VB.Menu mnuCol 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(���������)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmCheckCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSelectStock As Integer           '�Ƿ��ѡ�ⷿ
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5�������̵��¼��,�����̵��;6��ȫ����Ϊ��
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean                '��һ����ʾ
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Dim mstrPrivs As String                     'Ȩ��
Private mblnNoStock As Boolean              '���ز������Ƿ������̵�û�����ô洢�ⷿ��ҩƷ
Private mblnLoadData As Boolean             '���ڼ���Ƿ���װ�����ݣ������Ѵ��ڵ��ݣ�
Private mstr����ID As String
Private mbln��ͣ��ҩƷ As Boolean
Private mbln�����̵�ʱ�� As Boolean         'Ϊ��ʱʼ���Ե�ǰ�����Ϊ��������
Private mbln���Է������ As Boolean         'Ϊ��ʱ����ҩƷ�ķ������
Private mbln����ҩƷ�̵����� As Boolean     'Ϊ��ʱ����ҩƷ���̵�����
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private Const MStrCaption As String = "ҩƷ�̵��"
Private mbln���䶯 As Boolean         '������Ƿ�䶯   true-�Ѿ���飬false-δ��飬ֻ�����ҵ�������

Private mstr�̵㵥�� As String              '�̵㵥��(��¼���������̵����̵㵥��)
Private mblnɾ���̵㵥 As Boolean           '���������̵����Ƿ�ɾ����Ӧ���̵㵥

Private mlngFindFirst As Long
Private mlngFind As Long                             '���ڲ���
Private mrsFindName As ADODB.Recordset              '���ڲ���

Private mblnNotTrigger As Boolean
Private mblnKeyPressReturn As Boolean

Private Const mlngColor_��ӯ As Long = vbRed
Private Const mlngColor_�̿� As Long = vbBlue
Private Const mlngColor_��ƽ As Long = vbBlack
Private mlngCurrColor As Long
Private mlngNextColor As Long
'Private blnColorRefresh As Boolean
Private mstrMsg As String
Private mlongCurrRow As Long                '��ǰѡ����
Private mlngFindCurrRow As Long             '��ѯ���ĵ�ǰ��
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mlng�ⷿ As Long

Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ�⣨˵��������0ʱ�д�С��װ���֣�����0ʱΪĬ�ϰ�װ��
Private mint��λ As Integer
Private mintС��λ As Integer

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
Private mbln���������� As Boolean         '�̿�ʱ������������0������飻1�����

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��

Private mintMoneyDigit As Integer           '���С��λ��

Private mintCostDigit0 As Integer            'С��λ�ɱ���С��λ��
Private mintPriceDigit0 As Integer           'С��λ�ۼ�С��λ��
Private mintNumberDigit0 As Integer          'С��λ����С��λ��

Private mintCostDigit1 As Integer            '��λ�ɱ���С��λ��
Private mintPriceDigit1 As Integer           '��λ�ۼ�С��λ��
Private mintNumberDigit1 As Integer          '��λ����С��λ��


Private mintMaxMoneyBit As Integer          'ҩƷ�����н��С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

Private Const mcstComment As String = "��-��ƽ;��-��ӯ;��-�̿�;����-ͣ��ҩƷ"

Private mlngSum As Long '��¼��治��ҩƷ����

'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntColҩ�� As Integer = 2
Private Const mconIntCol��Ʒ�� As Integer = 3
Private Const mconIntCol��Դ As Integer = 4
Private Const mconIntCol����ҩ�� As Integer = 5
Private Const mconIntCol��� As Integer = 6
Private Const mconIntCol��� As Integer = 7
Private Const mconIntCol���� As Integer = 8
Private Const mconIntCol�������� As Integer = 9
Private Const mconIntCol����ϵ�� As Integer = 10
Private Const mconIntCol����ϵ���� As Integer = 11
Private Const mconIntCol����ϵ��С As Integer = 12
Private Const mconIntcol�ӳ��� As Integer = 13
Private Const mconIntColʵ�ʲ�� As Integer = 14
Private Const mconIntColʵ�ʽ�� As Integer = 15
Private Const mconIntCol���� As Integer = 16
Private Const mconIntCol�ⷿ��λ As Integer = 17
Private Const mconIntCol��λ As Integer = 18

Private Const mconIntCol���� As Integer = 19
Private Const mconIntColЧ�� As Integer = 20
Private Const mconIntCol��׼�ĺ� As Integer = 21

Private Const mconintCol�������� As Integer = 22

Private Const mconintCol���װ�������� As Integer = 23
Private Const mconIntCol����������λ�� As Integer = 24

Private Const mconintColС��װ�������� As Integer = 25
Private Const mconIntCol����������λС As Integer = 26

Private Const mconintColʵ������ As Integer = 27

Private Const mconintCol���װʵ������ As Integer = 28
Private Const mconIntColʵ��������λ�� As Integer = 29

Private Const mconintColС��װʵ������ As Integer = 30
Private Const mconIntColʵ��������λС As Integer = 31

Private Const mconintCol�ϼ� As Integer = 32
Private Const mconintCol��־ As Integer = 33
Private Const mconintCol������ As Integer = 34
Private Const mconintCol�ɱ��� As Integer = 35
Private Const mconIntCol�ۼ� As Integer = 36
Private Const mconintCol���� As Integer = 37
Private Const mconintCol��۲� As Integer = 38
Private Const mconintCol�̵��� As Integer = 39
Private Const mconintCol�̵�ɱ���� As Integer = 40
Private Const mconintCol�̵�ɱ����� As Integer = 41
Private Const mconintCol������� As Integer = 42      'ȡ���ԭʼ����
Private Const mconIntColҩƷ��������� As Integer = 43
Private Const mconIntColҩƷ���� As Integer = 44
Private Const mconIntColҩƷ���� As Integer = 45
Private Const mconIntCol������ As Integer = 46
Private Const mconIntCol������� As Integer = 47
Private Const mconIntCol�������� As Integer = 48
Private Const mconIntColS  As Integer = 49              '������
'=========================================================================================

Private Function CheckUnVerify(ByVal lng�ⷿID As Long) As Boolean
    '���δ��˵��ݣ��������ʾͨ�����
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = " Select id From ҩƷ�շ���¼" & _
            " Where ����� Is NULL And �ⷿID=[1] AND Rownum<2 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���δ��˵���", lng�ⷿID)
    If rsData.EOF Then
        CheckUnVerify = True
    Else
        CheckUnVerify = False
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Get��С��λ()
    Dim intUnit As Integer, strUnit As String, strDefault As String
    Dim strCompare As String
    Dim str��С��λ As String
    Dim int���� As Integer
    
    Const conInt���㾫�� As Integer = 0
    
    Const conIntҩƷ As Integer = 1
    
    Const conint�ۼ۵�λ As Integer = 1
    Const conint���ﵥλ As Integer = 2
    Const conintסԺ��λ As Integer = 3
    Const conintҩ�ⵥλ As Integer = 4
    
    Const conInt�ɱ��� As Integer = 1
    Const conInt�ۼ� As Integer = 2
    Const conInt���� As Integer = 3
    Const conInt��� As Integer = 4
    
    int���� = conInt���㾫��
        
    strCompare = "ҩ�ⵥλ;���ﵥλ;סԺ��λ;�ۼ۵�λ"
    
    'ȡ�ô��װ��λ
    strDefault = GetDrugUnit(Val(txtStock.Tag), "ҩƷ�̵����")
    
    'ȡ��С��װ��λ
    intUnit = Val(zlDatabase.GetPara("С��װ��λ", glngSys, ģ���.ҩƷ�̵�))
    
    If intUnit = 0 Then
        strUnit = strDefault
    Else
        strUnit = Split(strCompare, ";")(intUnit - 1)
    End If

    '��ָ����λ��ȱʡ��λ����λ��С��λ��˳������
    mintUnit = 0
    If strUnit <> strDefault Then
        If InStr(1, strCompare, strUnit) < InStr(1, strCompare, strDefault) Then
            str��С��λ = strUnit & "|" & strDefault
        Else
            mintUnit = 0
            str��С��λ = strDefault & "|" & strUnit
        End If
        
        mintMoneyDigit = GetDigit(int����, conIntҩƷ, conInt���)
        
        Call GetDrugDigit(mlng�ⷿ, "ҩƷ�̵����", 0, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    Else
        Call GetDrugDigit(mlng�ⷿ, "ҩƷ�̵����", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End If
    
    If str��С��λ = "" Then Exit Sub
    
    'ȡ��λ�ľ��ȣ��ۼۡ���������
    Select Case Split(str��С��λ, "|")(0)
        Case "�ۼ۵�λ"
            mint��λ = conint�ۼ۵�λ
        Case "���ﵥλ"
            mint��λ = conint���ﵥλ
        Case "סԺ��λ"
            mint��λ = conintסԺ��λ
        Case "ҩ�ⵥλ"
            mint��λ = conintҩ�ⵥλ
    End Select
    
    mintCostDigit1 = GetDigit(int����, conIntҩƷ, conInt�ɱ���, mint��λ)
    mintPriceDigit1 = GetDigit(int����, conIntҩƷ, conInt�ۼ�, mint��λ)
    mintNumberDigit1 = GetDigit(int����, conIntҩƷ, conInt����, mint��λ)

    'ȡС��λ�ľ��ȣ�������
    Select Case Split(str��С��λ, "|")(1)
        Case "�ۼ۵�λ"
            mintС��λ = conint�ۼ۵�λ
        Case "���ﵥλ"
            mintС��λ = conint���ﵥλ
        Case "סԺ��λ"
            mintС��λ = conintסԺ��λ
        Case "ҩ�ⵥλ"
            mintС��λ = conintҩ�ⵥλ
    End Select
    
    mintCostDigit0 = GetDigit(int����, conIntҩƷ, conInt�ɱ���, mintС��λ)
    mintPriceDigit0 = GetDigit(int����, conIntҩƷ, conInt�ۼ�, mintС��λ)
    mintNumberDigit0 = GetDigit(int����, conIntҩƷ, conInt����, mintС��λ)
    
'    '����С������󾫶�ȡֵ����������̲��ɾ�
'    mintNumberDigit = gtype_UserDrugDigits.Digit_����
'    mintNumberDigit0 = gtype_UserDrugDigits.Digit_����
End Sub
Private Sub RefreshListSN()
    '���������������
    Dim lngRow As Long
    
    With vsfBill
        .Redraw = flexRDNone
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                .TextMatrix(lngRow, mconIntCol�к�) = lngRow
            End If
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If vsfBill.rows < 2 Then Exit Sub
    If vsfBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To vsfBill.rows - 1
            If vsfBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(vsfBill.TextMatrix(n, mconIntCol���)) = 0, n, Val(vsfBill.TextMatrix(n, mconIntCol���)))
                !ҩƷid = Val(vsfBill.TextMatrix(n, 0))
                !���� = Val(vsfBill.TextMatrix(n, mconIntCol����))
                
                .Update
            End If
        Next
        
    End With
End Sub
'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    GetDepend = False
    strSQL = "SELECT B.Id " _
           & "FROM ҩƷ�������� A, ҩƷ������ B " _
           & "Where A.���id = B.ID AND A.���� = 12  and b.ϵ��=1 "
    Set rsDepend = zlDatabase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ�̵�������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    strSQL = "SELECT B.Id " _
           & "FROM ҩƷ�������� A, ҩƷ������ B " _
           & "Where A.���id = B.ID AND A.���� = 12  and b.ϵ��=-1 "
    Set rsDepend = zlDatabase.OpenSQLRecord(strSQL, MStrCaption)
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ�̵��ĳ����������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    GetDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetStocktakingColor(ByVal vsfObj As VSFlexGrid, ByVal Row As Long)
    '�̿���ӯ������ɫ���֣���ɫ����-��ӯ����ɫ����-�̿�����ɫ����-��ƽ
    With vsfObj
        .Row = Row
        mlngCurrColor = .CellForeColor
        If .TextMatrix(Row, mconintCol��־) = "ӯ" Then
            mlngNextColor = mlngColor_��ӯ
        ElseIf .TextMatrix(Row, mconintCol��־) = "��" Then
            mlngNextColor = mlngColor_�̿�
        Else
            mlngNextColor = mlngColor_��ƽ
        End If
        
        If mlngNextColor <> mlngCurrColor Then
            .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = mlngNextColor
        End If
    End With
End Sub

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mblnFirst = True
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1307)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If mint�༭״̬ = 4 Then
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub vsfBill_MoveNextCell(ByVal Row As Long, ByVal Col As Long)
    With vsfBill
        Select Case Col
            Case mconIntColҩ��
                If Val(.TextMatrix(Row, 0)) = 0 Then Exit Sub
                .Col = IIf(mintUnit = 0, mconintCol���װʵ������, mconintColʵ������)
            Case mconIntCol����
                If Val(.TextMatrix(Row, mconIntCol����)) = -1 And .TextMatrix(Row, mconIntColЧ��) = "" Then
                    .Col = mconIntColЧ��
                Else
                    .Col = IIf(mintUnit = 0, mconintCol���װʵ������, mconintColʵ������)
                End If
            Case mconIntColЧ��
                .Col = IIf(mintUnit = 0, mconintCol���װʵ������, mconintColʵ������)
            Case mconintColʵ������
                If Row < .rows - 1 Then
                    .Row = Row + 1
                    If .TextMatrix(.Row, mconIntColҩ��) = "" Then
                        .Col = mconIntColҩ��
                    Else
                        .Col = mconintColʵ������
                    End If
                Else
                    If Val(.TextMatrix(Row, 0)) <> 0 Then
                        .rows = .rows + 1
                        .Row = .rows - 1
                        .Col = mconIntColҩ��
                    End If
                End If
            Case mconintCol���װʵ������, mconintColС��װʵ������
                If Col = mconintCol���װʵ������ Then
                    If .ColWidth(mconintColС��װʵ������) > 0 Then
                        .Col = mconintColС��װʵ������
                    Else
                        '�����һ��Ϊ�ջ���ҩ����Ϊ���򷵻ص�ҩ���У����򷵻ص�ʵ��������
                        If Row < .rows - 1 Then
                            .Row = Row + 1
                            If .TextMatrix(Row, mconIntColҩ��) <> "" Then
                                .Col = mconintCol���װʵ������
                            Else
                                .Col = mconIntColҩ��
                            End If
                        Else
                            If Val(.TextMatrix(Row, 0)) <> 0 Then
                                .rows = .rows + 1
                                .Row = .rows - 1
                                .Col = mconIntColҩ��
                            End If
                        End If
                    End If
                Else
                    If Row < .rows - 1 Then
                        .Row = Row + 1
                        If .TextMatrix(Row, mconIntColҩ��) <> "" Then
                            .Col = mconintCol���װʵ������
                        Else
                            .Col = mconIntColҩ��
                        End If
                    Else
                        If Val(.TextMatrix(Row, 0)) <> 0 Then
                            .rows = .rows + 1
                            .Row = .rows - 1
                            .Col = mconIntColҩ��
                        End If
                    End If
                End If
        End Select
        
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub cmdBatch_Click()
    '��֤������еļ�¼����ȡ����
    Dim rsPhysic As ADODB.Recordset 'ҩƷ����¼��
    Dim rsDetail As ADODB.Recordset
    Dim str�̵����� As String
    Dim dbl�ɱ��� As Double, dbl���ۼ� As Double, dbl�ӳ��� As Double
    Dim bln�ⷿ As Boolean
    Dim intMoneyBit As Integer
    Dim intOld As Integer
    Dim intCol As Integer
    Dim rsʱ�۷��� As ADODB.Recordset
    Dim strҩ�� As String
    Dim strOrder As String, strCompare As String
    Dim str�̵�ʱ�� As String
    
    str�̵�ʱ�� = txtCheckDate.Caption
    
    If MsgBox("�����������������������ݽ�������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    Else
        With vsfBill
            .rows = 2
            For intCol = 0 To .Cols - 1
                .TextMatrix(1, intCol) = ""
            Next
        End With
    End If
    
    strOrder = zlDatabase.GetPara("����", glngSys, ģ���.ҩƷ�̵�)
    strCompare = Mid(strOrder, 1, 1)
    
    gstrSQL = "Select  Distinct a.ҩƷid, b.����, b.����, c.�ⷿ��λ " & _
        " From ҩƷ��� A, �շ���ĿĿ¼ B, ҩƷ�����޶� C " & _
        " where��a.���� = 1 And a.ҩƷid = b.Id And a.�ⷿid = c.�ⷿid(+) And a.ҩƷid = c.ҩƷid(+) And a.�ⷿid = [1]" & _
        " And (Nvl(A.ʵ������,0)<>0 Or Nvl(A.ʵ�ʽ��,0)<>0 Or Nvl(A.ʵ�ʲ��,0)<>0 )"

    
    If mbln���Է������ = False Then
        gstrSQL = gstrSQL & _
            " and (Decode(B.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(1,3)) " & _
                " or Decode(B.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(2,3)) " & _
                " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[1]))"
    End If
    
    gstrSQL = gstrSQL & " Order by " & _
          IIf(strCompare = "0", "B.����", IIf(strCompare = "1", "B.����", IIf(strCompare = "2", "B.����", "C.�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc") & ",B.����"
    
    Set rsPhysic = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���ҩƷ", Val(txtStock.Tag))
    With vsfBill
        Do While Not rsPhysic.EOF
            'ȡ��ҩƷ����ϸ��Ϣ�����ֶܷ�����Σ�
            Set rsDetail = GetPhysicDetail(Val(txtStock.Tag), rsPhysic!ҩƷid, False, False, False)
            Do While Not rsDetail.EOF
                If rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1 Then .rows = .rows + 1
                'ʱ��ҩƷ�����ۼ�
                dbl�ɱ��� = Nvl(rsDetail!ƽ���ɱ���, 0)
                dbl���ۼ� = Nvl(rsDetail!�ۼ�, 0)
                If rsDetail!�Ƿ��� = 1 Then
                    dbl���ۼ� = Get�̵�ʱ�����ۼ�(CLng(rsPhysic!ҩƷid), Val(txtStock.Tag), CLng(rsDetail!����), 1, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                '������������и�ʽ��
                .TextMatrix(.rows - 1, 0) = rsPhysic!ҩƷid
                
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rsDetail!ͨ����
                Else
                    strҩ�� = IIf(IsNull(rsDetail!��Ʒ��), rsDetail!ͨ����, rsDetail!��Ʒ��)
                End If
                
                .TextMatrix(.rows - 1, mconIntColҩƷ���������) = rsDetail!ҩƷ���� & strҩ��
                .TextMatrix(.rows - 1, mconIntColҩƷ����) = rsDetail!ҩƷ����
                .TextMatrix(.rows - 1, mconIntColҩƷ����) = strҩ��
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ����)
                Else
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ���������)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol��Ʒ��) = IIf(IsNull(rsDetail!��Ʒ��), "", rsDetail!��Ʒ��)
                
                .TextMatrix(.rows - 1, mconIntCol��Դ) = zlStr.Nvl(rsDetail!ҩƷ��Դ)
                .TextMatrix(.rows - 1, mconIntCol����ҩ��) = zlStr.Nvl(rsDetail!����ҩ��)
                .TextMatrix(.rows - 1, mconIntCol���) = IIf(IsNull(rsDetail!���), "", rsDetail!���)
                .TextMatrix(.rows - 1, mconIntCol����) = zlStr.Nvl(rsDetail!����, zlStr.Nvl(rsDetail!ȱʡ����))
                .TextMatrix(.rows - 1, mconIntCol�ⷿ��λ) = IIf(IsNull(rsDetail!�ⷿ��λ), "", rsDetail!�ⷿ��λ)
                .TextMatrix(.rows - 1, mconIntCol����) = IIf(IsNull(rsDetail!����), "", rsDetail!����)
                .TextMatrix(.rows - 1, mconIntColЧ��) = IIf(IsNull(rsDetail!Ч��), "", Format(rsDetail!Ч��, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(.rows - 1, mconIntColЧ��) <> "" Then
                    '����Ϊ��Ч��
                    .TextMatrix(.rows - 1, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntColЧ��)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol��׼�ĺ�) = IIf(IsNull(rsDetail!��׼�ĺ�), "", rsDetail!��׼�ĺ�)
                .TextMatrix(.rows - 1, mconIntColʵ�ʽ��) = zlStr.Nvl(rsDetail!ʵ�ʽ��, 0)
                .TextMatrix(.rows - 1, mconIntColʵ�ʲ��) = zlStr.Nvl(rsDetail!ʵ�ʲ��, 0)
                .TextMatrix(.rows - 1, mconIntcol�ӳ���) = rsDetail!�ӳ��� / 100 & "||" & rsDetail!�Ƿ��� & "||" & rsDetail!ҩ����������
                .TextMatrix(.rows - 1, mconintCol��־) = "ƽ"
                .TextMatrix(.rows - 1, mconintCol������) = "0"
                .TextMatrix(.rows - 1, mconintCol�������) = zlStr.Nvl(rsDetail!��������, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol��λ) = IIf(IsNull(rsDetail!��λ), "", rsDetail!��λ)
                    .TextMatrix(.rows - 1, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol����ϵ��) = zlStr.Nvl(rsDetail!����ϵ��, 0)
                    .TextMatrix(.rows - 1, mconintColʵ������) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol��������), mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintColʵ������)) * Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol��������) = zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��
                    .TextMatrix(.rows - 1, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�ɱ���, 0) * rsDetail!����ϵ��, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��С, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol����ϵ����) = zlStr.Nvl(rsDetail!����ϵ����, 0)
                    .TextMatrix(.rows - 1, mconIntCol����ϵ��С) = zlStr.Nvl(rsDetail!����ϵ��С, 0)
                    .TextMatrix(.rows - 1, mconIntCol����������λ��) = rsDetail!���װ��λ
                    .TextMatrix(.rows - 1, mconIntCol����������λС) = rsDetail!С��װ��λ
                    .TextMatrix(.rows - 1, mconIntColʵ��������λ��) = rsDetail!���װ��λ
                    .TextMatrix(.rows - 1, mconIntColʵ��������λС) = rsDetail!С��װ��λ
                    .TextMatrix(.rows - 1, mconintCol���װ��������) = Int(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ����)
                    .TextMatrix(.rows - 1, mconintCol���װʵ������) = .TextMatrix(.rows - 1, mconintCol���װ��������)
                    .TextMatrix(.rows - 1, mconintColС��װ��������) = zlStr.FormatEx((Val(rsDetail!��������) - Val(.TextMatrix(.rows - 1, mconintCol���װ��������)) * Val(rsDetail!����ϵ����)) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintColС��װʵ������) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintColС��װ��������), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintColʵ������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol�ϼ�) = .TextMatrix(.rows - 1, mconintColʵ������) & .TextMatrix(.rows - 1, mconIntColʵ��������λС)
                    .TextMatrix(.rows - 1, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintColʵ������)) * Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol��������) = zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С
                    .TextMatrix(.rows - 1, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�ɱ���, 0) * rsDetail!����ϵ��С, mintCostDigit0, , True)
                End If
                
                
                '����Ƿ���ҩƷ�������θ���Ϊ-1����ʾ��������
                .TextMatrix(.rows - 1, mconIntCol����) = zlStr.Nvl(rsDetail!����, 0)
                If CheckPhysicBatch(bln�ⷿ, rsDetail!��������, rsDetail!ҩ����������) And Val(.TextMatrix(.rows - 1, mconIntCol����)) = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol����) = -1
'                    '�����ã��Զ�Ϊ������������������Ч��
'                    .TextMatrix(.Rows - 1, mconIntCol����) = "20040601"
'                    .TextMatrix(.Rows - 1, mconIntColЧ��) = "2006-06-01"
                End If
                 
                If Val(.TextMatrix(.rows - 1, mconintColʵ������)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) = Val(.TextMatrix(.rows - 1, mconintCol�ɱ���))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                .TextMatrix(.rows - 1, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) * Val(.TextMatrix(.rows - 1, mconintColʵ������)) - Val(.TextMatrix(.rows - 1, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) - Val(.TextMatrix(.rows - 1, mconintCol�ɱ���))) * Val(.TextMatrix(.rows - 1, mconintColʵ������)) - Val(.TextMatrix(.rows - 1, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                
                If mbln��ͣ��ҩƷ = True Then
                    '�����ͣ��ҩƷ�����д�����ʾ
                    If Format(rsDetail!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, .rows - 1, 0, .rows - 1, .Cols - 1) = True
                    End If
                End If
                '.TextMatrix(.rows - 1, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol�ɱ���)) * Val(.TextMatrix(.rows - 1, mconintColʵ������)), mintMoneyDigit)
                '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                .TextMatrix(.rows - 1, mconintCol�̵�ɱ����) = zlStr.FormatEx((zlStr.Nvl(rsDetail!ʵ�ʽ��, 0) + Val(.TextMatrix(.rows - 1, mconintCol����))) - (zlStr.Nvl(rsDetail!ʵ�ʲ��, 0) + Val(.TextMatrix(.rows - 1, mconintCol��۲�))), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol����)) - Val(.TextMatrix(.rows - 1, mconintCol��۲�)), mintMoneyDigit, , True)
                
                
                '���÷�������
                Call GetҩƷ��������(.rows - 1)
                
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmCheckCard)
            rsPhysic.MoveNext
        Loop
        
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintColʵ������, .rows - 1, mconintColʵ������) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol���װʵ������, .rows - 1, mconintCol���װʵ������) = True
            .Cell(flexcpFontBold, 1, mconintColС��װʵ������, .rows - 1, mconintColС��װʵ������) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) <> "" Then
        vsfBill.Col = IIf(mintUnit = 0, mconintCol���װʵ������, mconintColʵ������)
    Else
        vsfBill.Col = mconIntColҩ��
    End If
    If Me.Visible = True Then
        vsfBill.SetFocus
'        vsfBill.EditCell
    End If
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdReset_Click()
    Dim str��;ID As String, str�ⷿ��λ As String, str���ͱ��� As String, strALL���ͱ��� As String
    Dim str���ʷ��� As String, lng�ⷿID As Long, int�̵㷽ʽ As Integer, str�̵�ʱ�� As String
    Dim int���޿��ҩƷ As Integer, bln�̵㵥 As Boolean   '�Ƿ�ֻ����̵㵥�е�ҩƷ�����̵㣬FALSE-��ʾ������ҩƷ�����̵㣬�̵㵥�в����ڵ�ҩƷ�Զ���Ϊ��
    Dim bln���޿���н��ҩƷ As Boolean
    Dim intCol As Integer
    
'    If mblnFirst = False Then Exit Sub
    
    With vsfBill
        If MsgBox("�����������������������ݽ�������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End With
    
    mblnLoadData = False
    If mintParallelRecord <> 1 Then mblnChange = False
    vsfBill.ToolTipText = mcstComment
    
    '��ʼ������
    str��;ID = "": str���ͱ��� = ""
    
    If mint�༭״̬ = 1 Then
        '�Զ��������ֹ������̵��
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        If frmCheckCondition.GetCondition(mfrmMain, str���ͱ���, lng�ⷿID, int�̵㷽ʽ, str�̵�ʱ��, int���޿��ҩƷ, str�ⷿ��λ, bln���޿���н��ҩƷ, mstr����ID, mbln�����̵�ʱ��) = True Then
            If mlng�ⷿ = 0 Then
                mlng�ⷿ = lng�ⷿID
            End If
            vsfBill.rows = 2
            For intCol = 0 To vsfBill.Cols - 1
                vsfBill.TextMatrix(1, intCol) = ""
            Next
'            Call Get��С��λ
            Call SearchData(str���ͱ���, lng�ⷿID, int�̵㷽ʽ, str�̵�ʱ��, (int���޿��ҩƷ = 1), str�ⷿ��λ, bln���޿���н��ҩƷ)
        Else
            vsfBill.rows = 2
            For intCol = 0 To vsfBill.Cols - 1
                vsfBill.TextMatrix(1, intCol) = ""
            Next
            Exit Sub
        End If
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
    ElseIf mint�༭״̬ = 5 Then
        '�����̵������ָ��ʱ�̵��̵��¼����ָ��ʱ�̵Ŀ�棩
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        If FrmCheckCourseCondition.GetCondition(mfrmMain, lng�ⷿID, mstr�̵㵥��, bln�̵㵥, mblnɾ���̵㵥) = True Then
            If mlng�ⷿ = 0 Then
                mlng�ⷿ = lng�ⷿID
            End If
            vsfBill.rows = 2
            Call Get��С��λ
            Call SearchTableData(lng�ⷿID, bln�̵㵥)
        Else
            Exit Sub
        End If
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
    End If
    
    mblnLoadData = True
End Sub

Private Sub cmdSet_Click()
    Dim lngRow As Long, n As Long
    Dim rsDetail As ADODB.Recordset
    Dim lngҩƷID As Long, lng���� As Long, dblʵ������ As Double
    Dim dlbSum As Double
    Dim intMoneyBit As Integer
    Dim dbl���� As Double, dbl��۲� As Double
    
    On Error GoTo ErrHand
    
    If MsgBox("�ò�����ҩƷ��ʵ���������ܵ���������ϣ��Ƿ���иò�����", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '���Ǳ�����򣬿�����ͬҩƷ���������ģ��Ȱѽ�������װ�����ݼ�����
    Set rsDetail = New ADODB.Recordset
    With rsDetail
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To vsfBill.rows - 1
            If vsfBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !ҩƷid = Val(vsfBill.TextMatrix(n, 0))
                !���� = Val(vsfBill.TextMatrix(n, mconIntCol����))
                !ʵ������ = Val(vsfBill.TextMatrix(n, mconintColʵ������))
                
                .Update
            End If
        Next
        
        .Sort = "ҩƷid,����"
        
        Do While Not .EOF
            If lngҩƷID <> !ҩƷid Then
                dlbSum = !ʵ������
                lngҩƷID = !ҩƷid
            Else
                dlbSum = dlbSum + !ʵ������
            End If
            
            !ʵ������ = 0
            .Update
            
            .MoveNext
            
            '��������Ѿ�û�������˻��ߺ��治��ͬһ��ҩƷʱ����ʵ���������ܵ����һ��������
            If .EOF Then
                .MovePrevious
                !ʵ������ = dlbSum
                .Update
                
                .MoveNext
            Else
                If lngҩƷID <> !ҩƷid Then
                    .MovePrevious
                    !ʵ������ = dlbSum
                    .Update
                    
                    .MoveNext
                End If
            End If
        Loop
    End With
    
    
    
    With vsfBill
        .Redraw = flexRDNone
        
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                lngҩƷID = Val(vsfBill.TextMatrix(lngRow, 0))
                lng���� = Val(vsfBill.TextMatrix(lngRow, mconIntCol����))
                
                rsDetail.Filter = "ҩƷid=" & lngҩƷID & " And ����=" & lng����
                If Not rsDetail.EOF Then
                    '�����ݼ���ʵ�����������̵�����
                    dblʵ������ = rsDetail!ʵ������
                    
                    '����ɴ�С��װ��λ
                    If mintUnit = 0 Then
                        .TextMatrix(lngRow, mconintCol���װʵ������) = zlStr.FormatEx(Int(dblʵ������ / Val(.TextMatrix(lngRow, mconIntCol����ϵ����))), mintNumberDigit0, , True)
                        .TextMatrix(lngRow, mconintColС��װʵ������) = zlStr.FormatEx((dblʵ������ - Val(.TextMatrix(lngRow, mconintCol���װʵ������)) * Val(.TextMatrix(lngRow, mconIntCol����ϵ����))) / Val(.TextMatrix(lngRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                        .TextMatrix(lngRow, mconintCol�ϼ�) = zlStr.FormatEx(dblʵ������, mintNumberDigit, , True) & .TextMatrix(lngRow, mconIntCol����������λС)
                    End If
                    
                    .TextMatrix(lngRow, mconintColʵ������) = zlStr.FormatEx(dblʵ������, mintNumberDigit, , True)
                    .TextMatrix(lngRow, mconintCol������) = zlStr.FormatEx(Abs(dblʵ������ - Val(.TextMatrix(lngRow, mconintCol��������))), mintNumberDigit, , True)
                    If dblʵ������ > Val(.TextMatrix(lngRow, mconintCol��������)) Then
                        .TextMatrix(lngRow, mconintCol��־) = "ӯ"
                    ElseIf dblʵ������ < Val(.TextMatrix(lngRow, mconintCol��������)) Then
                        .TextMatrix(lngRow, mconintCol��־) = "��"
                    Else
                        .TextMatrix(lngRow, mconintCol��־) = "ƽ"
                    End If
                
                    '���ҩƷ���������Ϊ0�������۲�Ϊ0��ҩƷ�޷�ͨ���̵��������¼������
                    '��������µ�ͨ��ҩƷ�������۵�ʵ��λ������ϵͳ���������õĽ��λ��
                    '����취�����ʵ������Ϊ0�������Ͳ�۲�С��λ�����ֺ�ҩƷ�����н��Ͳ��λ��һ��
                    If Val(.TextMatrix(lngRow, mconIntCol������)) = 1 Then
                        intMoneyBit = mintMoneyDigit
                    ElseIf dblʵ������ = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(lngRow, 0))) = True And Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) = Val(.TextMatrix(lngRow, mconintCol�ɱ���))) Then
                        '��0��������ҩƷ�̵�ʱ
                        intMoneyBit = mintMaxMoneyBit
                    Else
                        intMoneyBit = mintMoneyDigit
                    End If
                
                    '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                    '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                    .TextMatrix(lngRow, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) * dblʵ������ - Val(.TextMatrix(lngRow, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                    .TextMatrix(lngRow, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) - Val(.TextMatrix(lngRow, mconintCol�ɱ���))) * dblʵ������ - Val(.TextMatrix(lngRow, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                    dbl���� = Val(.TextMatrix(lngRow, mconintCol����))
                    dbl��۲� = Val(.TextMatrix(lngRow, mconintCol��۲�))
                    If .TextMatrix(lngRow, mconintCol��־) = "��" Then
                        .TextMatrix(lngRow, mconintCol����) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol����)), intMoneyBit, , True)
                        .TextMatrix(lngRow, mconintCol��۲�) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol��۲�)), intMoneyBit, , True)
                    End If
                
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .TextMatrix(lngRow, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) * dblʵ������, mintMoneyDigit, , True)
                
                    '.TextMatrix(lngRow, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol�ɱ���)) * Val(.TextMatrix(lngRow, mconintColʵ������)), mintMoneyDigit)
                    '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                    .TextMatrix(lngRow, mconintCol�̵�ɱ����) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntColʵ�ʽ��)) + dbl����) - (Val(.TextMatrix(lngRow, mconIntColʵ�ʲ��)) + dbl��۲�), mintMoneyDigit, , True)
                    .TextMatrix(lngRow, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol����)) - Val(.TextMatrix(lngRow, mconintCol��۲�)), mintMoneyDigit, , True)
                
                    '�̿���ӯ������ɫ����
                    Call SetStocktakingColor(vsfBill, lngRow)
                End If
            End If
        Next
        
        .Redraw = flexRDDirect
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdZero_Click()
    Dim lngRow As Integer
    Dim dblʵ������ As Double
    Dim dbl���� As Double, dbl��۲� As Double
    Dim intMoneyBit As Integer
    
    If MsgBox("�Ƿ��ʵ�������㣿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    dblʵ������ = 0
    
    With vsfBill
        .Redraw = flexRDNone
        
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
    
                '����ɴ�С��װ��λ
                If mintUnit = 0 Then
                      .TextMatrix(lngRow, mconintCol���װʵ������) = zlStr.FormatEx(dblʵ������, mintNumberDigit0, , True)
                      .TextMatrix(lngRow, mconintColС��װʵ������) = zlStr.FormatEx(dblʵ������, mintNumberDigit0, , True)
                      .TextMatrix(lngRow, mconintCol�ϼ�) = zlStr.FormatEx(dblʵ������, mintNumberDigit, , True) & .TextMatrix(lngRow, mconIntCol����������λС)
                End If
              
                .TextMatrix(lngRow, mconintColʵ������) = zlStr.FormatEx(dblʵ������, mintNumberDigit, , True)
                .TextMatrix(lngRow, mconintCol������) = zlStr.FormatEx(Abs(dblʵ������ - Val(.TextMatrix(lngRow, mconintCol��������))), mintNumberDigit, , True)
                If dblʵ������ > Val(.TextMatrix(lngRow, mconintCol��������)) Then
                    .TextMatrix(lngRow, mconintCol��־) = "ӯ"
                ElseIf dblʵ������ < Val(.TextMatrix(lngRow, mconintCol��������)) Then
                    .TextMatrix(lngRow, mconintCol��־) = "��"
                Else
                    .TextMatrix(lngRow, mconintCol��־) = "ƽ"
                End If
                
                  intMoneyBit = mintMaxMoneyBit
        
                  '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                  '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                  .TextMatrix(lngRow, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) * dblʵ������ - Val(.TextMatrix(lngRow, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                  .TextMatrix(lngRow, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) - Val(.TextMatrix(lngRow, mconintCol�ɱ���))) * dblʵ������ - Val(.TextMatrix(lngRow, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                  dbl���� = Val(.TextMatrix(lngRow, mconintCol����))
                  dbl��۲� = Val(.TextMatrix(lngRow, mconintCol��۲�))
                  If .TextMatrix(lngRow, mconintCol��־) = "��" Then
                      .TextMatrix(lngRow, mconintCol����) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol����)), intMoneyBit, , True)
                      .TextMatrix(lngRow, mconintCol��۲�) = zlStr.FormatEx(-1 * Val(.TextMatrix(lngRow, mconintCol��۲�)), intMoneyBit, , True)
                  End If
          
                  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                  .TextMatrix(lngRow, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) * dblʵ������, mintMoneyDigit, , True)
        
                  '.TextMatrix(lngRow, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol�ɱ���)) * Val(.TextMatrix(lngRow, mconintColʵ������)), mintMoneyDigit)
                  '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                  .TextMatrix(lngRow, mconintCol�̵�ɱ����) = zlStr.FormatEx((Val(.TextMatrix(lngRow, mconIntColʵ�ʽ��)) + dbl����) - (Val(.TextMatrix(lngRow, mconIntColʵ�ʲ��)) + dbl��۲�), mintMoneyDigit, , True)
                  .TextMatrix(lngRow, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconintCol����)) - Val(.TextMatrix(lngRow, mconintCol��۲�)), mintMoneyDigit, , True)
              
                '�̿���ӯ������ɫ����
                Call SetStocktakingColor(vsfBill, lngRow)
            End If
        Next
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub cmd�̶���_Click()
    Call PopupMenu(PopMenu, 2)
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            txtCode.SetFocus
        End If
    ElseIf KeyCode = vbKeyF3 Then
        If Trim(txtCode.Text) = "" Then
            txtCode.SetFocus
        Else
            Call FindGridRow(txtCode.Text)
        End If
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub CheckDataUpdate()
    '��������Ƿ����仯������仯����ʾ�û����Զ����½�������
    'ֻ�����ʱ�ŵ��ô˹���
    Dim intRow As Integer
    Dim lngҩƷID As Long
    Dim lng�ⷿID As Long
    Dim lng���� As Long
    Dim dat�̵�ʱ�� As Date
    Dim dblԭ�������� As Double
    Dim dbl���������� As Double
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    Dim intMoneyBit As Integer
    Dim rsTemp As ADODB.Recordset
    Dim bln�䶯 As Boolean
    
    On Error GoTo ErrHand
    
    If mint�༭״̬ = 3 Then
        With vsfBill
            If .rows > 1 Then
                Call FS.ShowFlash("����ҩƷ�䶯,���Ժ� ...", Me)
                
                lng�ⷿID = txtStock.Tag
                .Redraw = flexRDNone
                For intRow = 1 To .rows - 1
                    If Val(.TextMatrix(intRow, 0)) <> 0 Then
                        lngҩƷID = Val(.TextMatrix(intRow, 0))
                        lng���� = Val(.TextMatrix(intRow, mconIntCol����))
                        dat�̵�ʱ�� = CDate(txtCheckDate.Caption)
                        dblԭ�������� = Val(.TextMatrix(intRow, mconintCol�������))
                        
                        gstrSQL = "Select �ⷿid, ҩƷid, ����, Nvl(Sum(ʵ������), 0) As ��������, Nvl(Sum(�̵�����), 0) As �̵�����, Nvl(Sum(ʵ�ʽ��), 0) As ʵ�ʽ��," & vbNewLine & _
                                    "       Nvl(Sum(ʵ�ʲ��), 0) As ʵ�ʲ��, Nvl(Sum(��������), 0) As ��������" & vbNewLine & _
                                    "From (Select a.�ⷿid, a.ҩƷid, Nvl(����, 0) As ����, Nvl(a.ʵ������, 0) ʵ������, 0 �̵�����, Nvl(a.ʵ�ʽ��, 0) ʵ�ʽ��, Nvl(a.ʵ�ʲ��, 0) ʵ�ʲ��," & vbNewLine & _
                                    "              Nvl(a.��������, 0) ��������" & vbNewLine & _
                                    "       From ҩƷ��� A" & vbNewLine & _
                                    "       Where a.���� = 1 And a.�ⷿid = [1] And a.ҩƷid = [2] And Nvl(a.����, 0) = [3]" & vbNewLine & _
                                    "       Union All" & vbNewLine & _
                                    "       Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, Sum(-1 * a.���ϵ�� * a.ʵ������ * a.����) As ʵ������, 0 �̵�����," & vbNewLine & _
                                    "              Sum(-1 * a.���ϵ�� * a.���۽��) As ʵ�ʽ��, Sum(-1 * a.���ϵ�� * a.���) As ʵ�ʲ��, 0 As ��������" & vbNewLine & _
                                    "       From ҩƷ�շ���¼ A" & vbNewLine & _
                                    "       Where a.�ⷿid + 0 = [1] And a.ҩƷid + 0 = [2] And Nvl(a.����, 0) = [3] And a.������� > [4]" & vbNewLine & _
                                    "       Group By a.�ⷿid, a.ҩƷid, a.����)" & vbNewLine & _
                                    "Group By �ⷿid, ҩƷid, ����"

                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���䶯���", lng�ⷿID, lngҩƷID, lng����, dat�̵�ʱ��)
                        
                        If rsTemp.RecordCount > 0 Then
                            dbl���������� = rsTemp!��������
                            If dblԭ�������� <> dbl���������� Then
                                bln�䶯 = True
                                                                
                                .TextMatrix(intRow, mconintCol�������) = Nvl(rsTemp!��������, 0)
                                .TextMatrix(intRow, mconIntColʵ�ʽ��) = zlStr.Nvl(rsTemp!ʵ�ʽ��, 0)
                                .TextMatrix(intRow, mconIntColʵ�ʲ��) = zlStr.Nvl(rsTemp!ʵ�ʲ��, 0)
                                If mintUnit > 0 Then
                                    .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsTemp!��������, 0) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintNumberDigit, , True)
                                Else
                                    .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsTemp!��������, 0) / Val(.TextMatrix(intRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                                    
                                    .TextMatrix(intRow, mconintCol���װ��������) = zlStr.FormatEx(Int(zlStr.Nvl(rsTemp!��������, 0) / Val(.TextMatrix(intRow, mconIntCol����ϵ����))), mintNumberDigit0, , True)
                                    .TextMatrix(intRow, mconintColС��װ��������) = zlStr.FormatEx((Val(rsTemp!��������) - Val(.TextMatrix(intRow, mconintCol���װ��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ����))) / Val(.TextMatrix(intRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                                     
                                    .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(zlStr.Nvl(rsTemp!��������, 0) / Val(.TextMatrix(intRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                                End If

                                If Val(.TextMatrix(intRow, mconintColʵ������)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True And Val(.TextMatrix(intRow, mconIntCol�ۼ�)) = Val(.TextMatrix(intRow, mconintCol�ɱ���))) Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) = Val(.TextMatrix(.rows - 1, mconintCol�ɱ���))) Then
                                    intMoneyBit = mintMaxMoneyBit
                                Else
                                    intMoneyBit = mintMoneyDigit
                                End If
                                
                                .TextMatrix(intRow, mconintCol������) = zlStr.FormatEx(Abs(Val(.TextMatrix(intRow, mconintColʵ������)) - Val(.TextMatrix(intRow, mconintCol��������))), mintNumberDigit, , True)
                                If Val(.TextMatrix(intRow, mconintColʵ������)) > Val(.TextMatrix(intRow, mconintCol��������)) Then
                                    .TextMatrix(intRow, mconintCol��־) = "ӯ"
                                ElseIf Val(.TextMatrix(intRow, mconintColʵ������)) < Val(.TextMatrix(intRow, mconintCol��������)) Then
                                    .TextMatrix(intRow, mconintCol��־) = "��"
                                Else
                                    .TextMatrix(intRow, mconintCol��־) = "ƽ"
                                End If

                                '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                                '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                                .TextMatrix(intRow, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) * Val(.TextMatrix(intRow, mconintColʵ������)) - Val(.TextMatrix(intRow, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                                .TextMatrix(intRow, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol�ۼ�)) - Val(.TextMatrix(intRow, mconintCol�ɱ���))) * Val(.TextMatrix(intRow, mconintColʵ������)) - Val(.TextMatrix(intRow, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                                dbl���� = Val(.TextMatrix(intRow, mconintCol����))
                                dbl��۲� = Val(.TextMatrix(intRow, mconintCol��۲�))
                                If .TextMatrix(intRow, mconintCol��־) = "��" Then
                                    .TextMatrix(intRow, mconintCol����) = zlStr.FormatEx(-1 * Val(.TextMatrix(intRow, mconintCol����)), intMoneyBit, , True)
                                    .TextMatrix(intRow, mconintCol��۲�) = zlStr.FormatEx(-1 * Val(.TextMatrix(intRow, mconintCol��۲�)), intMoneyBit, , True)
                                End If
                            
                                '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                                .TextMatrix(intRow, mconintCol�̵�ɱ����) = zlStr.FormatEx((zlStr.Nvl(rsTemp!ʵ�ʽ��, 0) + dbl����) - (zlStr.Nvl(rsTemp!ʵ�ʲ��, 0) + dbl��۲�), mintMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol����)) - Val(.TextMatrix(intRow, mconintCol��۲�)), mintMoneyDigit, , True)

                            End If
                        End If
                    End If
                Next
                .Redraw = flexRDDirect
                If bln�䶯 = True Then
                    MsgBox "��淢���仯�����Զ����½������ݣ����飡", vbInformation, gstrSysName
                    mbln���䶯 = True
                End If
            End If
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    Dim intLop As Integer
    Dim strҩƷ As String '��¼������������ʱ��ҩƷ��������Ϊ��
    
    '�����������ݼ�
    Call SetSortRecord
    
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        
        '�Զ�������鲢ִ�е���
        Call AutoAdjustPrice_ByNO(12, mstr���ݺ�)
    
        mstrTime_End = GetBillInfo(12, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub
        
        '�����˱䶯�ֽ�ԭʼ�̵㵥ɾ��Ȼ���ٲ���NO��ͬ���µ��̵㵥
        If mbln���䶯 = True Then
            BlnSuccess = SaveCard
        End If
        If mbln���䶯 = False Then
            '������Ƿ����仯
            Call CheckDataUpdate
            If mbln���䶯 = True Then
                Exit Sub
            End If
        End If
        
        '���۹�������Ƿ���ڲ��������۵�ҩƷ
        For intLop = 1 To vsfBill.rows - 1
            If Val(vsfBill.TextMatrix(intLop, mconIntCol������)) = 0 Then
                '������������ʱ
                If vsfBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                    If IsPriceAdjustMod(Val(vsfBill.TextMatrix(intLop, 0))) = True Then
                        If CheckPriceAdjust(Val(vsfBill.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(vsfBill.TextMatrix(intLop, mconIntCol����))) = False Then
                            MsgBox "��" & intLop & "��ҩƷ���������۹���������¼���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                            vsfBill.SetFocus
                            vsfBill.Row = intLop
                            vsfBill.TopRow = intLop
                            Exit Sub
                        End If
                    End If
                End If
            Else
                '����ʱ
                If vsfBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                    If IsPriceAdjustMod(Val(vsfBill.TextMatrix(intLop, 0))) = True Then
                        '��������۹����������ۼۺͳɱ��۹�ϵ
                        If Val(vsfBill.TextMatrix(intLop, mconintCol�ɱ���)) <> Val(vsfBill.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                            MsgBox "��" & intLop & "��ҩƷ���������۹������̵������ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                            vsfBill.SetFocus
                            vsfBill.Row = intLop
                            vsfBill.TopRow = intLop
                            Exit Sub
                        End If
                    End If
                End If
            End If
            
            
            If vsfBill.TextMatrix(intLop, mconintCol��־) = "��" Then '�̿�ʱ���⣬������Ƿ��㹻
                If mintUnit > 0 Then '��С��װ��ͬ
                    If Not ���ʵ���������(Val(vsfBill.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(vsfBill.TextMatrix(intLop, mconIntCol����)), Val(vsfBill.TextMatrix(intLop, mconintCol������)), Val(vsfBill.TextMatrix(intLop, mconIntCol����ϵ��)), mintNumberDigit) Then
                        mlngSum = mlngSum + 1
                        If mlngSum <= 3 Then 'ƴ��ʾ��Ϣ��
                            mstrMsg = mstrMsg & "��" & vsfBill.TextMatrix(intLop, mconIntColҩ��) & "(" & vsfBill.TextMatrix(intLop, mconIntCol����) & "��" & "����" & Chr(10)
                        End If
                    End If
                Else
                    If Not ���ʵ���������(Val(vsfBill.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(vsfBill.TextMatrix(intLop, mconIntCol����)), Val(vsfBill.TextMatrix(intLop, mconintCol������)), Val(vsfBill.TextMatrix(intLop, mconIntCol����ϵ��С)), mintNumberDigit0) Then
                        mlngSum = mlngSum + 1
                        If mlngSum <= 3 Then 'ƴ��ʾ��Ϣ��
                            mstrMsg = mstrMsg & "��" & vsfBill.TextMatrix(intLop, mconIntColҩ��) & "(" & vsfBill.TextMatrix(intLop, mconIntCol����) & "��" & "����" & Chr(10)
                        End If
                    End If
                End If
            End If
        Next
        
        '��治����ʾ��Ϣ
        If mlngSum > 0 Then
            If mint����� = 1 Then '��������
                If MsgBox(mstrMsg & IIf(mlngSum <= 3, mlngSum & "��ҩƷ��治�㣬�Ƿ������", "��" & mlngSum & "��ҩƷ��治�㣬�Ƿ������"), vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then
                    mlngSum = 0
                    mstrMsg = ""
                    Exit Sub
                End If
            ElseIf mint����� = 2 Then '�����ֹ
                MsgBox mstrMsg & IIf(mlngSum <= 3, mlngSum & "��ҩƷ��治�㣬������ˣ�", "��" & mlngSum & "��ҩƷ��治�㣬������ˣ�"), vbInformation, gstrSysName
                mlngSum = 0
                mstrMsg = ""
                Exit Sub
            End If
        End If
        mlngSum = 0
        mstrMsg = ""
        
        If SaveCheck = True Then
            If Val(zlDatabase.GetPara("��˴�ӡ", glngSys, ģ���.ҩƷ�̵�)) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
            
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDatabase.GetPara("���̴�ӡ", glngSys, ģ���.ҩƷ�̵�)) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                printbill
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    vsfBill.rows = 2
    vsfBill.Cell(flexcpText, 1, 0, 1, vsfBill.Cols - 1) = ""

    Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
    txtժҪ.Text = ""
    mblnChange = False
    
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Activate()
    Dim str��;ID As String, str�ⷿ��λ As String, str���ͱ��� As String, strALL���ͱ��� As String
    Dim str���ʷ��� As String, lng�ⷿID As Long, int�̵㷽ʽ As Integer, str�̵�ʱ�� As String
    Dim int���޿��ҩƷ As Integer, bln�̵㵥 As Boolean   '�Ƿ�ֻ����̵㵥�е�ҩƷ�����̵㣬FALSE-��ʾ������ҩƷ�����̵㣬�̵㵥�в����ڵ�ҩƷ�Զ���Ϊ��
    Dim bln���޿���н��ҩƷ As Boolean
    
    If mblnFirst = False Then Exit Sub
    
    mstr����ID = ""
    mblnLoadData = False
    mintBatchNoLen = GetBatchNoLen()
    If mintParallelRecord <> 1 Then mblnChange = False
    vsfBill.ToolTipText = mcstComment
    
    mbln��ͣ��ҩƷ = IIf(Val(zlDatabase.GetPara("����ͣ�õ�ҩƷ", glngSys, 1307, 0)) = 0, False, True)
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            '�����ѱ�ɾ��
            MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 5
            MsgBox "������δ��˵�ҩƷ���ݣ���ȫ����˺����ԣ�", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
     
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    mblnFirst = False
    '��ʼ������
    str��;ID = "": str���ͱ��� = ""
    
    If mint�༭״̬ = 1 Then
        '�Զ��������ֹ������̵��
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        If frmCheckCondition.GetCondition(mfrmMain, str���ͱ���, lng�ⷿID, int�̵㷽ʽ, str�̵�ʱ��, int���޿��ҩƷ, str�ⷿ��λ, bln���޿���н��ҩƷ, mstr����ID, mbln�����̵�ʱ��) = True Then
            If mlng�ⷿ = 0 Then
                mlng�ⷿ = lng�ⷿID
            End If
            Call Get��С��λ
            Call SearchData(str���ͱ���, lng�ⷿID, int�̵㷽ʽ, str�̵�ʱ��, (int���޿��ҩƷ = 1), str�ⷿ��λ, bln���޿���н��ҩƷ)
        Else
            Unload Me
            Exit Sub
        End If
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
        
    ElseIf mint�༭״̬ = 5 Then
        '�����̵������ָ��ʱ�̵��̵��¼����ָ��ʱ�̵Ŀ�棩
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        If FrmCheckCourseCondition.GetCondition(mfrmMain, lng�ⷿID, mstr�̵㵥��, bln�̵㵥, mblnɾ���̵㵥) = True Then
            If mlng�ⷿ = 0 Then
                mlng�ⷿ = lng�ⷿID
            End If
            Call Get��С��λ
            Call SearchTableData(lng�ⷿID, bln�̵㵥)
        Else
            Unload Me
            Exit Sub
        End If
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
    ElseIf mint�༭״̬ = 6 Then
        'ȫ����Ϊ��
        str�̵�ʱ�� = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        txtCheckDate = str�̵�ʱ��
        txtStock.Caption = mfrmMain.cboStock.Text
        lng�ⷿID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        txtStock.Tag = lng�ⷿID
        mint����� = MediWork_GetCheckStockRule(Val(txtStock.Tag))
        
        If mlng�ⷿ = 0 Then
            mlng�ⷿ = lng�ⷿID
        End If
        Call Get��С��λ
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        Call SearchTableData(lng�ⷿID)
        If CmdCancel.Enabled = False Then
            CmdCancel.Enabled = True
        End If
        If CmdSave.Enabled = False Then
            CmdSave.Enabled = True
        End If
        If vsfBill.Visible = True Then
            vsfBill.SetFocus
        End If
    End If
    
    mblnLoadData = True
End Sub

Private Sub SetSortCode()
    '����ҩƷ���뷵�ظ�ʽ�����������
    '�����п��ܺ���"-"���ţ��������б�����"-"ǰ��༸λ��"-"����༸λ�����б��붼�����λ�����и�ʽ������
    Dim str���� As String
    Dim lngRow As Long
    Dim intǰ׺ As Integer
    Dim int��׺ As Integer
    Dim str����ǰ׺ As String
    Dim str�����׺ As String
    Dim blnLine As Boolean
    
    With vsfBill
       For lngRow = 1 To vsfBill.rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                str���� = Replace(.TextMatrix(lngRow, mconIntColҩƷ����), "[", "")
                str���� = Replace(str����, "]", "")
                
                If InStr(1, str����, "-") > 0 Then
                    blnLine = True
                    If Len(Mid(str����, 1, InStr(str����, "-") - 1)) > intǰ׺ Then
                        intǰ׺ = Len(Mid(str����, 1, InStr(str����, "-") - 1))
                    End If
                    
                    If Len(Mid(str����, InStr(str����, "-") + 1)) > int��׺ Then
                        int��׺ = Len(Mid(str����, InStr(str����, "-") + 1))
                    End If
                Else
                    If Len(str����) > intǰ׺ Then
                        intǰ׺ = Len(str����)
                    End If
                End If
            End If
        Next
        
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, 0) <> "" Then
                str���� = Replace(.TextMatrix(lngRow, mconIntColҩƷ����), "[", "")
                str���� = Replace(str����, "]", "")
                
                If blnLine = False Then
                    .TextMatrix(lngRow, mconIntCol�������) = Format(str����, String(intǰ׺, "0"))
                Else
                    If InStr(str����, "-") > 0 Then
                        str����ǰ׺ = Mid(str����, 1, InStr(str����, "-") - 1)
                        str�����׺ = Mid(str����, InStr(str����, "-") + 1)
                        
                        str����ǰ׺ = Format(str����ǰ׺, String(intǰ׺, "0"))
                        str�����׺ = Format(str�����׺, String(int��׺, "0"))
                    Else
                        str����ǰ׺ = Format(str����, String(intǰ׺, "0"))
                        str�����׺ = String(int��׺, "0")
                    End If
                    
                    .TextMatrix(lngRow, mconIntCol�������) = str����ǰ׺ & "-" & str�����׺
                End If
            End If
        Next
    End With
End Sub
Private Sub SearchData(ByVal str���ͱ��� As String, ByVal lng�ⷿID As Long, _
    ByVal int�̵㷽ʽ As Integer, ByVal str�̵�ʱ�� As String, ByVal bln���޿��ҩƷ As Boolean, ByVal str�ⷿ��λ As String, ByVal bln���޿���н��ҩƷ As Boolean)
    
    Dim rsPhysic As ADODB.Recordset 'ҩƷ����¼��
    Dim rsDetail As ADODB.Recordset
    Dim str�̵����� As String
    Dim dbl�ɱ��� As Double, dbl���ۼ� As Double, dbl�ӳ��� As Double
    Dim bln�ⷿ As Boolean
    Dim intMoneyBit As Integer
    Dim intOld As Integer
    Dim n As Integer
    Dim rsʱ�۷��� As ADODB.Recordset
    Dim strҩ�� As String
    Dim rsTemp As ADODB.Recordset
    Dim strArry As Variant
    Dim x As Long
    Dim strTemp As String
    Dim j As Long
    Dim str��λid As String
    Dim str��λ As String
    
'    On Error Resume Next
    On Error GoTo ErrHandle
    
    '��ʼ�����ݼ�
    Set rsPhysic = New ADODB.Recordset
    With rsPhysic
        If .State = 1 Then .Close
        .Fields.Append "ҩƷid", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 80, adFldIsNullable
        .Fields.Append "�ⷿ��λ", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '���ý�����ʾ����
    Select Case int�̵㷽ʽ
        Case 1
            staThis.Panels(2).Text = "���ڶ�" & txtStock & "��ҩƷ�������̵�"
        Case 2
            staThis.Panels(2).Text = "���ڶ�" & txtStock & "��ҩƷ�������̵�"
        Case 3
            staThis.Panels(2).Text = "���ڶ�" & txtStock & "��ҩƷ�������̵�"
        Case 4
            staThis.Panels(2).Text = "���ڶ�" & txtStock & "��ҩƷ���м����̵�"
        Case 5
            staThis.Panels(2).Text = "���ڶ����е�ҩƷ���м����̵�"
    End Select
    str�̵����� = " And Substr(A.�̵�����," & int�̵㷽ʽ & ",1)='1'"
    If int�̵㷽ʽ = 5 Then str�̵����� = "����"
    Call FS.ShowFlash("���ڼ���ҩƷ�������,���Ժ� ...", Me)
    DoEvents
    
    x = 1
    strArry = Array()
    str��λid = ""
    For j = 0 To UBound(Split(str�ⷿ��λ, ",")) - 1
        str��λ = Mid(str�ⷿ��λ, x, InStr(x, str�ⷿ��λ, ",") - x)
        x = InStr(x, str�ⷿ��λ, ",") + 1
        If Len(IIf(str��λid = "", "", str��λid & ",") & str��λ) > 4000 Then
            ReDim Preserve strArry(UBound(strArry) + 1)
            strArry(UBound(strArry)) = str��λid
            str��λid = str��λ
        Else
            str��λid = IIf(str��λid = "", "", str��λid & ",") & str��λ
        End If
    Next
    
    If str��λid <> "" Then
        ReDim Preserve strArry(UBound(strArry) + 1)
        strArry(UBound(strArry)) = str��λid
    End If
    
    If str�ⷿ��λ = "" Then
        Set rsPhysic = GetPhysic(lng�ⷿID, str�̵�����, str���ͱ���, str�ⷿ��λ, bln���޿��ҩƷ, False, False, bln���޿���н��ҩƷ)
    Else
        For j = 0 To UBound(strArry)
            Set rsTemp = GetPhysic(lng�ⷿID, str�̵�����, str���ͱ���, CStr(strArry(j)), bln���޿��ҩƷ, False, False, bln���޿���н��ҩƷ)
            If Not rsTemp.EOF Then
                Do While Not rsTemp.EOF
                    With rsPhysic
                        .AddNew
                        !ҩƷid = rsTemp!ҩƷid
                        !���� = rsTemp!����
                        !���� = rsTemp!����
                        !�ⷿ��λ = rsTemp!�ⷿ��λ
                        
                        .Update
                    End With
                    rsTemp.MoveNext
                Loop
            End If
        Next
    End If
    
    Call FS.StopFlash
    
    If rsPhysic.RecordCount = 0 Then
        If mint�༭״̬ = 6 Then
            MsgBox "δ����ȷ��ȡҩƷ�������,�����ԣ�", vbInformation, gstrSysName: Exit Sub
        Else
            MsgBox "δ����ȷ��ȡҩƷ�������,�����Ի��ֹ�����ҩƷ��", vbInformation, gstrSysName
            vsfBill.Row = 1
            vsfBill.Col = mconIntColҩ��
            Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("����װ��ҩƷ����,���Ժ� ...", Me)
    DoEvents
    vsfBill.Redraw = flexRDNone
    
    bln�ⷿ = CheckPartProp(lng�ⷿID)
    With vsfBill
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            'ȡ��ҩƷ����ϸ��Ϣ�����ֶܷ�����Σ�
            Set rsDetail = GetPhysicDetail(lng�ⷿID, rsPhysic!ҩƷid, bln���޿��ҩƷ, False, bln���޿���н��ҩƷ)
            Do While Not rsDetail.EOF
                If rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1 Then .rows = .rows + 1
                'ʱ��ҩƷ�����ۼ�
                dbl�ɱ��� = zlStr.Nvl(rsDetail!ƽ���ɱ���, 0)
                dbl���ۼ� = zlStr.Nvl(rsDetail!�ۼ�, 0)
                If rsDetail!�Ƿ��� = 1 Then
                    dbl���ۼ� = Get�̵�ʱ�����ۼ�(CLng(rsPhysic!ҩƷid), lng�ⷿID, CLng(rsDetail!����), 1, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                '������������и�ʽ��
                .TextMatrix(.rows - 1, 0) = rsPhysic!ҩƷid
                
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rsDetail!ͨ����
                Else
                    strҩ�� = IIf(IsNull(rsDetail!��Ʒ��), rsDetail!ͨ����, rsDetail!��Ʒ��)
                End If
                
                .TextMatrix(.rows - 1, mconIntColҩƷ���������) = rsDetail!ҩƷ���� & strҩ��
                .TextMatrix(.rows - 1, mconIntColҩƷ����) = rsDetail!ҩƷ����
                .TextMatrix(.rows - 1, mconIntColҩƷ����) = strҩ��
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ����)
                Else
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ���������)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol��Ʒ��) = IIf(IsNull(rsDetail!��Ʒ��), "", rsDetail!��Ʒ��)
                
                .TextMatrix(.rows - 1, mconIntCol��Դ) = zlStr.Nvl(rsDetail!ҩƷ��Դ)
                .TextMatrix(.rows - 1, mconIntCol����ҩ��) = zlStr.Nvl(rsDetail!����ҩ��)
                .TextMatrix(.rows - 1, mconIntCol���) = IIf(IsNull(rsDetail!���), "", rsDetail!���)
                .TextMatrix(.rows - 1, mconIntCol����) = zlStr.Nvl(rsDetail!����, zlStr.Nvl(rsDetail!ȱʡ����))
                .TextMatrix(.rows - 1, mconIntCol�ⷿ��λ) = IIf(IsNull(rsDetail!�ⷿ��λ), "", rsDetail!�ⷿ��λ)
                .TextMatrix(.rows - 1, mconIntCol����) = IIf(IsNull(rsDetail!����), "", rsDetail!����)
                .TextMatrix(.rows - 1, mconIntColЧ��) = IIf(IsNull(rsDetail!Ч��), "", Format(rsDetail!Ч��, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(.rows - 1, mconIntColЧ��) <> "" Then
                    '����Ϊ��Ч��
                    .TextMatrix(.rows - 1, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntColЧ��)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol��׼�ĺ�) = IIf(IsNull(rsDetail!��׼�ĺ�), "", rsDetail!��׼�ĺ�)
                .TextMatrix(.rows - 1, mconIntColʵ�ʽ��) = zlStr.Nvl(rsDetail!ʵ�ʽ��, 0)
                .TextMatrix(.rows - 1, mconIntColʵ�ʲ��) = zlStr.Nvl(rsDetail!ʵ�ʲ��, 0)
                .TextMatrix(.rows - 1, mconIntcol�ӳ���) = rsDetail!�ӳ��� / 100 & "||" & rsDetail!�Ƿ��� & "||" & rsDetail!ҩ����������
                .TextMatrix(.rows - 1, mconintCol��־) = "ƽ"
                .TextMatrix(.rows - 1, mconintCol������) = "0"
                .TextMatrix(.rows - 1, mconintCol�������) = zlStr.Nvl(rsDetail!��������, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol��λ) = IIf(IsNull(rsDetail!��λ), "", rsDetail!��λ)
                    .TextMatrix(.rows - 1, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol����ϵ��) = zlStr.Nvl(rsDetail!����ϵ��, 0)
                    .TextMatrix(.rows - 1, mconintColʵ������) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintCol��������), mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintColʵ������)) * Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�ɱ���, 0) * rsDetail!����ϵ��, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��С, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol����ϵ����) = zlStr.Nvl(rsDetail!����ϵ����, 0)
                    .TextMatrix(.rows - 1, mconIntCol����ϵ��С) = zlStr.Nvl(rsDetail!����ϵ��С, 0)
                    .TextMatrix(.rows - 1, mconIntCol����������λ��) = rsDetail!���װ��λ
                    .TextMatrix(.rows - 1, mconIntCol����������λС) = rsDetail!С��װ��λ
                    .TextMatrix(.rows - 1, mconIntColʵ��������λ��) = rsDetail!���װ��λ
                    .TextMatrix(.rows - 1, mconIntColʵ��������λС) = rsDetail!С��װ��λ
                    .TextMatrix(.rows - 1, mconintCol���װ��������) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ����), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol���װʵ������) = .TextMatrix(.rows - 1, mconintCol���װ��������)
                    .TextMatrix(.rows - 1, mconintColС��װ��������) = zlStr.FormatEx((Val(rsDetail!��������) - Val(.TextMatrix(.rows - 1, mconintCol���װ��������)) * Val(rsDetail!����ϵ����)) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintColС��װʵ������) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintColС��װ��������), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintColʵ������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol�ϼ�) = .TextMatrix(.rows - 1, mconintColʵ������) & .TextMatrix(.rows - 1, mconIntColʵ��������λС)
                    .TextMatrix(.rows - 1, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintColʵ������)) * Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�ɱ���, 0) * rsDetail!����ϵ��С, mintCostDigit0, , True)
                End If
                
                
                '����Ƿ���ҩƷ�������θ���Ϊ-1����ʾ��������
                .TextMatrix(.rows - 1, mconIntCol����) = zlStr.Nvl(rsDetail!����, 0)
                If CheckPhysicBatch(bln�ⷿ, rsDetail!��������, rsDetail!ҩ����������) And Val(.TextMatrix(.rows - 1, mconIntCol����)) = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol����) = -1
'                    '�����ã��Զ�Ϊ������������������Ч��
'                    .TextMatrix(.Rows - 1, mconIntCol����) = "20040601"
'                    .TextMatrix(.Rows - 1, mconIntColЧ��) = "2006-06-01"
                End If
                 
                If Val(.TextMatrix(.rows - 1, mconintColʵ������)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) = Val(.TextMatrix(.rows - 1, mconintCol�ɱ���))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                .TextMatrix(.rows - 1, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) * Val(.TextMatrix(.rows - 1, mconintColʵ������)) - Val(.TextMatrix(.rows - 1, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) - Val(.TextMatrix(.rows - 1, mconintCol�ɱ���))) * Val(.TextMatrix(.rows - 1, mconintColʵ������)) - Val(.TextMatrix(.rows - 1, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                
                If mbln��ͣ��ҩƷ = True Then
                    '�����ͣ��ҩƷ�����д�����ʾ
                    If Format(rsDetail!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, .rows - 1, 0, .rows - 1, .Cols - 1) = True
                    End If
                End If
                '.TextMatrix(.rows - 1, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol�ɱ���)) * Val(.TextMatrix(.rows - 1, mconintColʵ������)), mintMoneyDigit)
                '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                .TextMatrix(.rows - 1, mconintCol�̵�ɱ����) = zlStr.FormatEx((zlStr.Nvl(rsDetail!ʵ�ʽ��, 0) + Val(.TextMatrix(.rows - 1, mconintCol����))) - (zlStr.Nvl(rsDetail!ʵ�ʲ��, 0) + Val(.TextMatrix(.rows - 1, mconintCol��۲�))), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol����)) - Val(.TextMatrix(.rows - 1, mconintCol��۲�)), mintMoneyDigit, , True)
                
                '���÷�������
                Call GetҩƷ��������(.rows - 1)
                
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmCheckCard)
            rsPhysic.MoveNext
        Loop
        
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintColʵ������, .rows - 1, mconintColʵ������) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol���װʵ������, .rows - 1, mconintCol���װʵ������) = True
            .Cell(flexcpFontBold, 1, mconintColС��װʵ������, .rows - 1, mconintColС��װʵ������) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1
    If vsfBill.TextMatrix(1, 0) <> "" Then
        vsfBill.Col = IIf(mintUnit = 0, mconintCol���װʵ������, mconintColʵ������)
    Else
        vsfBill.Col = mconIntColҩ��
    End If
    If Me.Visible = True Then
        vsfBill.SetFocus
        vsfBill.EditCell
    End If
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SearchTableData(ByVal lng�ⷿID As Long, Optional ByVal bln�̵㵥 As Boolean = False)
    Dim strPhysic As String
    Dim dbl�ɱ��� As Double, dbl���ۼ� As Double, dbl�ӳ��� As Double
    Dim lngPhysic As Long
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim rsPhysic As New ADODB.Recordset 'ҩƷ����¼��
    Dim rsDetail As New ADODB.Recordset
    Dim n As Integer
    Dim intOld As Integer
    Dim rsʱ�۷��� As ADODB.Recordset
    Dim strҩ�� As String
    Dim lngDrugID As Long
    Dim rsDingPrice As ADODB.Recordset
    Dim intMoneyBit As Integer
    Dim dbl����, dbl��۲� As Double
    Dim str�̵�ʱ�� As String
    
'    On Error Resume Next
    On Error GoTo ErrHandle
    
    str�̵�ʱ�� = txtCheckDate.Caption
    
    Call FS.ShowFlash("���ڼ���ҩƷ�������,���Ժ� ...", Me)
    DoEvents
    Set rsPhysic = GetPhysic(lng�ⷿID, "����", "����", "����", False, IIf(mint�༭״̬ = 5, True, False), bln�̵㵥)
    Call FS.StopFlash
    
    If rsPhysic.RecordCount = 0 Then
        If mint�༭״̬ = 6 Then
            MsgBox "δ����ȷ��ȡҩƷ�������,�����ԣ�", vbInformation, gstrSysName: Exit Sub
        Else
            MsgBox "δ����ȷ��ȡҩƷ�������,�����Ի��ֹ�����ҩƷ��", vbInformation, gstrSysName: Exit Sub
        End If
    End If
    
    Call FS.ShowFlash("����װ��ҩƷ����,���Ժ� ...", Me)
    DoEvents
    
    With vsfBill
        .Redraw = flexRDNone
        Do While Not rsPhysic.EOF
            Set rsDetail = GetPhysicDetail(lng�ⷿID, rsPhysic!ҩƷid, False, IIf(mint�༭״̬ = 5, True, False))
            Do While Not rsDetail.EOF
                If rsPhysic.AbsolutePosition > 1 Or rsDetail.AbsolutePosition > 1 Then .rows = .rows + 1
                dbl�ɱ��� = zlStr.Nvl(rsDetail!�ɱ���, 0)
                dbl���ۼ� = IIf(IsNull(rsDetail!�ۼ�), 0, rsDetail!�ۼ�)
                '�������̵���������˵�ҩƷ
                If rsDetail!�Ƿ��� = 0 And IsNull(rsDetail!�ۼ�) Then
                    gstrSQL = "select �ּ� from �շѼ�Ŀ where �շ�ϸĿid=[1] and sysdate between ִ������ and ��ֹ����" & _
                            GetPriceClassString("")
                    
                    lngDrugID = rsPhysic!ҩƷid
                    
                    Set rsDingPrice = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ۸�", lngDrugID)
                    If rsDingPrice.EOF = False Then
                        dbl���ۼ� = rsDingPrice!�ּ�
                    End If
                End If
                
                If rsDetail!�Ƿ��� = 1 Then
                    dbl���ۼ� = Get�̵�ʱ�����ۼ�(CLng(rsDetail!ҩƷid), lng�ⷿID, CLng(rsDetail!����), 1, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                End If
                
                If Nvl(rsDetail!����, 0) = -1 Then
                    '����ҩƷû�����ξ��������̵����
                    .TextMatrix(.rows - 1, mconIntCol������) = "1"
                ElseIf CheckNoStock(Val(txtStock.Tag), Val(rsDetail!ҩƷid), Nvl(rsDetail!����, 0)) = True Then
                    '�޿��ʱ�̵���������̵����
                    .TextMatrix(.rows - 1, mconIntCol������) = "1"
                End If
                
                '���۹��������̵����ʱ�Լ۸���д���
                If gtype_UserSysParms.P275_���۹���ģʽ = 2 And .TextMatrix(.rows - 1, mconIntCol������) = "1" Then
                    If IsPriceAdjustMod(Val(rsDetail!ҩƷid)) = True Then
                        If rsDetail!�Ƿ��� = 1 Then
                            'ʱ��ʱ���ۼ�=�ɱ���
                            dbl���ۼ� = dbl�ɱ���
                        Else
                            '����ʱ���ɱ���=�ۼ�
                            dbl�ɱ��� = dbl���ۼ�
                        End If
                    End If
                End If

                '������������и�ʽ��
                .TextMatrix(.rows - 1, 0) = rsDetail!ҩƷid
                
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rsDetail!ͨ����
                Else
                    strҩ�� = IIf(IsNull(rsDetail!��Ʒ��), rsDetail!ͨ����, rsDetail!��Ʒ��)
                End If
                
                .TextMatrix(.rows - 1, mconIntColҩƷ���������) = rsDetail!ҩƷ���� & strҩ��
                .TextMatrix(.rows - 1, mconIntColҩƷ����) = rsDetail!ҩƷ����
                .TextMatrix(.rows - 1, mconIntColҩƷ����) = strҩ��
                
                If mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ����)
                Else
                    .TextMatrix(.rows - 1, mconIntColҩ��) = .TextMatrix(.rows - 1, mconIntColҩƷ���������)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol��Ʒ��) = IIf(IsNull(rsDetail!��Ʒ��), "", rsDetail!��Ʒ��)
                
                .TextMatrix(.rows - 1, mconIntCol��Դ) = zlStr.Nvl(rsDetail!ҩƷ��Դ)
                .TextMatrix(.rows - 1, mconIntCol����ҩ��) = IIf(IsNull(rsDetail!����ҩ��), "", rsDetail!����ҩ��)
                .TextMatrix(.rows - 1, mconIntCol���) = IIf(IsNull(rsDetail!���), "", rsDetail!���)
                .TextMatrix(.rows - 1, mconIntCol����) = IIf(IsNull(rsDetail!����), "", rsDetail!����)
                .TextMatrix(.rows - 1, mconIntCol�ⷿ��λ) = IIf(IsNull(rsDetail!�ⷿ��λ), "", rsDetail!�ⷿ��λ)
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol��λ) = IIf(IsNull(rsDetail!��λ), "", rsDetail!��λ)
                End If
                .TextMatrix(.rows - 1, mconIntCol����) = IIf(IsNull(rsDetail!����), "", rsDetail!����)
                .TextMatrix(.rows - 1, mconIntCol����) = IIf(IsNull(rsDetail!����), "", rsDetail!����)
                .TextMatrix(.rows - 1, mconIntColЧ��) = IIf(IsNull(rsDetail!Ч��), "", Format(rsDetail!Ч��, "yyyy-MM-dd"))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(.rows - 1, mconIntColЧ��) <> "" Then
                    '����Ϊ��Ч��
                    .TextMatrix(.rows - 1, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.rows - 1, mconIntColЧ��)), "yyyy-mm-dd")
                End If
                
                .TextMatrix(.rows - 1, mconIntCol��׼�ĺ�) = IIf(IsNull(rsDetail!��׼�ĺ�), "", rsDetail!��׼�ĺ�)
                
'                If mint�༭״̬ <> 5 Then
'                    .TextMatrix(.rows - 1, mconintCol������) =Str.FormatEx(rsDetail!������, mintNumberDigit)
'                End If
                If mint�༭״̬ = 5 Then
                    If mintUnit > 0 Then
                        .TextMatrix(.rows - 1, mconintColʵ������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�̵�����, 0) / rsDetail!����ϵ��, mintNumberDigit, , True)
                    Else
                        .TextMatrix(.rows - 1, mconintColʵ������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�̵�����, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                        .TextMatrix(.rows - 1, mconintCol�ϼ�) = .TextMatrix(.rows - 1, mconintColʵ������) & rsDetail!С��װ��λ
                    End If
                Else
                    '����������Ϊ0ʱ�������ľ���λ�����������ʾ
                    mintNumberDigit = 5
                    mintNumberDigit0 = 5
                    .TextMatrix(.rows - 1, mconintColʵ������) = zlStr.FormatEx(0, mintNumberDigit, , True)
                End If
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��, mintPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��, mintNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol����ϵ��) = zlStr.Nvl(rsDetail!����ϵ��, 0)
                Else
                    .TextMatrix(.rows - 1, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��С, mintPriceDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    
                    .TextMatrix(.rows - 1, mconIntCol����ϵ����) = zlStr.Nvl(rsDetail!����ϵ����, 0)
                    .TextMatrix(.rows - 1, mconIntCol����ϵ��С) = zlStr.Nvl(rsDetail!����ϵ��С, 0)
                    .TextMatrix(.rows - 1, mconIntCol����������λ��) = rsDetail!���װ��λ
                    .TextMatrix(.rows - 1, mconIntCol����������λС) = rsDetail!С��װ��λ
                    .TextMatrix(.rows - 1, mconIntColʵ��������λ��) = rsDetail!���װ��λ
                    .TextMatrix(.rows - 1, mconIntColʵ��������λС) = rsDetail!С��װ��λ
                    .TextMatrix(.rows - 1, mconintCol���װ��������) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ����), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintCol���װʵ������) = zlStr.FormatEx(Int(zlStr.Nvl(rsDetail!�̵�����, 0) / rsDetail!����ϵ����), mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintColС��װ��������) = zlStr.FormatEx((Val(rsDetail!��������) - Val(.TextMatrix(.rows - 1, mconintCol���װ��������)) * Val(rsDetail!����ϵ����)) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                    .TextMatrix(.rows - 1, mconintColС��װʵ������) = zlStr.FormatEx(.TextMatrix(.rows - 1, mconintColС��װ��������), mintNumberDigit0, , True)
                    
                    '����������Ϊ0ʱ�������ľ���λ�����������ʾ
                    If mint�༭״̬ = 6 Then
                        mintNumberDigit = 5
                        mintNumberDigit0 = 5
                        .TextMatrix(.rows - 1, mconintCol���װʵ������) = zlStr.FormatEx(0, mintNumberDigit, , True)
                        .TextMatrix(.rows - 1, mconintColС��װʵ������) = zlStr.FormatEx(0, mintNumberDigit, , True)
                    End If
                End If
                
                .TextMatrix(.rows - 1, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintColʵ������)) * Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconIntCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��, mintNumberDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconIntCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                End If
                
                .TextMatrix(.rows - 1, mconIntColʵ�ʽ��) = zlStr.Nvl(rsDetail!ʵ�ʽ��, 0)
                .TextMatrix(.rows - 1, mconIntColʵ�ʲ��) = zlStr.Nvl(rsDetail!ʵ�ʲ��, 0)
                .TextMatrix(.rows - 1, mconIntcol�ӳ���) = rsDetail!�ӳ��� / 100 & "||" & rsDetail!�Ƿ��� & "||" & rsDetail!ҩ����������
                
                If Val(.TextMatrix(.rows - 1, mconintCol��������)) > Val(.TextMatrix(.rows - 1, mconintColʵ������)) Then
                    .TextMatrix(.rows - 1, mconintCol��־) = "��"
                ElseIf Val(.TextMatrix(.rows - 1, mconintCol��������)) < Val(.TextMatrix(.rows - 1, mconintColʵ������)) Then
                    .TextMatrix(.rows - 1, mconintCol��־) = "ӯ"
                Else
                    .TextMatrix(.rows - 1, mconintCol��־) = "ƽ"
                End If
                
                .TextMatrix(.rows - 1, mconintCol������) = zlStr.FormatEx(Abs(Val(.TextMatrix(.rows - 1, mconintColʵ������)) - Val(.TextMatrix(.rows - 1, mconintCol��������))), mintNumberDigit, , True)
                .TextMatrix(.rows - 1, mconintCol�������) = zlStr.Nvl(rsDetail!��������, 0)
                
                If mintUnit > 0 Then
                    .TextMatrix(.rows - 1, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�ɱ���, 0) * rsDetail!����ϵ��, mintCostDigit, , True)
                Else
                    .TextMatrix(.rows - 1, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�ɱ���, 0) * rsDetail!����ϵ��С, mintCostDigit0, , True)
                End If
                
                If Val(.TextMatrix(.rows - 1, mconintColʵ������)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(.rows - 1, 0))) = True And Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) = Val(.TextMatrix(.rows - 1, mconintCol�ɱ���))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                .TextMatrix(.rows - 1, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) * Val(.TextMatrix(.rows - 1, mconintColʵ������)) - Val(.TextMatrix(.rows - 1, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                .TextMatrix(.rows - 1, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(.rows - 1, mconIntCol�ۼ�)) - Val(.TextMatrix(.rows - 1, mconintCol�ɱ���))) * Val(.TextMatrix(.rows - 1, mconintColʵ������)) - Val(.TextMatrix(.rows - 1, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                dbl���� = Val(.TextMatrix(.rows - 1, mconintCol����))
                dbl��۲� = Val(.TextMatrix(.rows - 1, mconintCol��۲�))
                
                If .TextMatrix(.rows - 1, mconintCol��־) = "��" Then
                    .TextMatrix(.rows - 1, mconintCol����) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol����)), intMoneyBit, , True)
                    .TextMatrix(.rows - 1, mconintCol��۲�) = zlStr.FormatEx(-1 * Val(.TextMatrix(.rows - 1, mconintCol��۲�)), intMoneyBit, , True)
                End If
                
                '.TextMatrix(.rows - 1, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol�ɱ���)) * Val(.TextMatrix(.rows - 1, mconintColʵ������)), mintMoneyDigit)
                '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                .TextMatrix(.rows - 1, mconintCol�̵�ɱ����) = zlStr.FormatEx((zlStr.Nvl(rsDetail!ʵ�ʽ��, 0) + dbl����) - (zlStr.Nvl(rsDetail!ʵ�ʲ��, 0) + dbl��۲�), mintMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(.rows - 1, mconintCol����)) - Val(.TextMatrix(.rows - 1, mconintCol��۲�)), mintMoneyDigit, , True)
                '�̿���ӯ������ɫ����
                Call SetStocktakingColor(vsfBill, .rows - 1)
                
                '���÷�������
                Call GetҩƷ��������(.rows - 1)
                
                .Col = mconintColʵ������
                rsDetail.MoveNext
            Loop
            Call zlControl.StaShowPercent(rsPhysic.AbsolutePosition / rsPhysic.RecordCount, staThis.Panels(2), frmCheckCard)
            rsPhysic.MoveNext
        Loop
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, 1, mconintColʵ������, .rows - 1, mconintColʵ������) = True
        Else
            .Cell(flexcpFontBold, 1, mconintCol���װʵ������, .rows - 1, mconintCol���װʵ������) = True
            .Cell(flexcpFontBold, 1, mconintColС��װʵ������, .rows - 1, mconintColС��װʵ������) = True
        End If
        
        Call SetSortCode
        
        .Redraw = flexRDDirect
    End With
    Call FS.StopFlash
    staThis.Panels(2).Text = ""
    vsfBill.Row = 1: vsfBill.Col = mconintColʵ������
    If Me.Visible = True Then
        vsfBill.SetFocus
    End If
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mlngFindCurrRow = 1
    mbln���������� = (Val(zlDatabase.GetPara("�̿�ʱ����������", glngSys, ģ���.ҩƷ�̵�)) = 1)
    mblnNoStock = (Val(zlDatabase.GetPara("�洢�ⷿ", glngSys, ģ���.ҩƷ�̵�)) = 1)
    mintMaxMoneyBit = gtype_UserDrugDigits.Digit_���
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    mbln���Է������ = (Val(zlDatabase.GetPara("����ҩƷ�������", glngSys, ģ���.ҩƷ�̵�)) = 1)
    
    txtStock = mfrmMain.cboStock.Text
    txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    mint����� = MediWork_GetCheckStockRule(Val(txtStock.Tag))
    mlng�ⷿ = txtStock.Tag
    Call Get��С��λ
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�̵����", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    mblnLoadData = False
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    initCard
    
    mstrTime_Start = GetBillInfo(12, mstr���ݺ�)
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim n As Integer
    Dim intOld As Integer
    Dim intMoneyBit As Integer
    Dim strҩ�� As String
    Dim strSqlOrder As String
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    
    On Error GoTo ErrHandle
    strOrder = zlDatabase.GetPara("����", glngSys, ģ���.ҩƷ�̵�)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        strSqlOrder = "ҩƷ����"
    ElseIf strCompare = "2" Then
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strSqlOrder = "ͨ����"
        Else
            strSqlOrder = "Nvl(��Ʒ��, ͨ����)"
        End If
    ElseIf strCompare = "3" Then
        strSqlOrder = "�ⷿ��λ"
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC") & ",ҩƷ����,���"
    
    Select Case mint�༭״̬
        Case 1, 5, 6
            Txt������ = UserInfo.�û�����
            Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            
            '�����ȫ����Ϊ�㣬�����Ƿ����δ��˵��̵㵥
'            If mint�༭״̬ = 6 Then
'                If CheckUnVerify(mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)) = False Then
'                    mintParallelRecord = 5
'                    Exit Sub
'                End If
'            End If
            
            cmd�̶���.Visible = (mint�༭״̬ = 1)
        Case 2, 3, 4
            initGrid
            If mint�༭״̬ <> 4 Then
                txtStock = mfrmMain.cboStock.Text
                txtStock.Tag = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
                mint����� = MediWork_GetCheckStockRule(Val(txtStock.Tag))
            Else
                gstrSQL = "select distinct b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id " _
                    & "and A.���� = 12 and a.no=[1] "
                Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                txtStock = rsInitCard!����
                txtStock.Tag = rsInitCard!id
                mint����� = MediWork_GetCheckStockRule(Val(txtStock.Tag))
                rsInitCard.Close
            End If
            
            If mintUnit > 0 Then
                '��С��װ��ͬʱ
                Select Case mintUnit
                    Case mconint�ۼ۵�λ
                        strUnitQuantity = "I.���㵥λ AS ��λ, A.��д���� AS ��������,A.���� AS ʵ������, A.ʵ������ AS ������,'1' as ����ϵ��,a.���ۼ� as �ۼ�,A.���� �ɱ���,"
                    Case mconint���ﵥλ
                        strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д����/ B.�����װ) AS ��������,(A.����/ B.�����װ) AS ʵ������, (A.ʵ������ / B.�����װ) AS ������,B.�����װ as ����ϵ��,a.���ۼ�*B.�����װ as �ۼ�,(A.����* B.�����װ) �ɱ���,"
                    Case mconintסԺ��λ
                        strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д����/ B.סԺ��װ) AS ��������,(A.����/ B.סԺ��װ) AS ʵ������, (A.ʵ������ / B.סԺ��װ) AS ������,B.סԺ��װ as ����ϵ��,a.���ۼ�*B.סԺ��װ as �ۼ�,(A.����*B.סԺ��װ) �ɱ���,"
                    Case mconintҩ�ⵥλ
                        strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д����/ B.ҩ���װ) AS ��������,(A.����/ B.ҩ���װ) AS ʵ������, (A.ʵ������ / B.ҩ���װ) AS ������,B.ҩ���װ as ����ϵ��,a.���ۼ�*B.ҩ���װ as �ۼ�,(A.����* B.ҩ���װ) �ɱ���,"
                End Select
            Else
                'ȡȫ����λ����װ���������ۼۣ��ɱ���ȡԭʼֵ
                strUnitQuantity = "I.���㵥λ As �ۼ۵�λ, B.���ﵥλ, B.סԺ��λ, B.ҩ�ⵥλ, A.��д���� AS ��������, A.���� AS ʵ������, A.ʵ������ AS ������, " & _
                            " '1' As ����ϵ���ۼ�, B.�����װ As ����ϵ������, B.סԺ��װ as ����ϵ��סԺ, B.ҩ���װ as ����ϵ��ҩ��, a.���ۼ� as �ۼ�, A.���� �ɱ���, "
            End If
            
            gstrSQL = "Select *" _
                    & " From " _
                    & "     (SELECT DISTINCT a.ҩƷid,A.���,a.���ϵ��,'[' || I.���� || ']' As ҩƷ����, I.���� As ͨ����, N.���� As ��Ʒ��," _
                    & "             B.ҩƷ��Դ,B.����ҩ��,I.���,A.����,Nvl(A.�ⷿ��λ,C.�ⷿ��λ) As �ⷿ��λ,A.����,a.Ч��,a.����," & strUnitQuantity _
                    & "             A.���۽�� as ����,A.��� as ��۲�, " _
                    & "             a.ժҪ,������,��������,�����,�������,a.Ƶ�� as �̵�ʱ��,a.�ɱ��� as �����,a.�ɱ���� as �����,b.�ӳ���,I.�Ƿ���,b.ҩ������ as ҩ����������,A.��д����,A.��׼�ĺ�,Nvl(A.��ҩ��ʽ,0) As ������, " _
                    & " Nvl(I.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) As ����ʱ�� " _
                    & "      From (Select a.�ⷿid,a.ҩƷid,A.���,a.���ϵ��,A.����,A.�ⷿ��λ,A.����,a.Ч��,a.����,A.��д����,A.����,A.ʵ������,a.���ۼ�,A.����,A.���۽��,A.���,a.ժҪ,������,��������,�����,�������,a.Ƶ��,a.�ɱ���,a.�ɱ����,A.��׼�ĺ�,A.��ҩ��ʽ " _
                    & "            From ҩƷ�շ���¼ A" _
                    & "            Where A.��¼״̬ =[2] AND A.���� =12 AND A.No = [1]) A," _
                    & "           ҩƷ��� b,�շ���ĿĿ¼ I ,�շ���Ŀ���� n,ҩƷ�����޶� C" _
                    & "      Where A.ҩƷid = B.ҩƷid And A.ҩƷid = I.id" _
                    & "            And A.ҩƷid=n.�շ�ϸĿid(+) And n.����(+)=3 " _
                    & "            And A.ҩƷID=C.ҩƷID(+) And A.�ⷿID=C.�ⷿID(+))" _
                    & " ORDER BY " & strSqlOrder
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, mint��¼״̬)
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Txt������ = rsInitCard!������
            If mint�༭״̬ = 2 Then
                Txt������ = UserInfo.�û�����
            End If
            Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
            
            Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
            Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            txtCheckDate.Caption = rsInitCard!�̵�ʱ��
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            intRow = 0
            With vsfBill
                .Redraw = flexRDNone
                Do While Not rsInitCard.EOF
                    
                    intRow = intRow + 1
                    'intRow = rsInitCard!���
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                        strҩ�� = rsInitCard!ͨ����
                    Else
                        strҩ�� = IIf(IsNull(rsInitCard!��Ʒ��), rsInitCard!ͨ����, rsInitCard!��Ʒ��)
                    End If
                    
                    .TextMatrix(intRow, mconIntColҩƷ���������) = rsInitCard!ҩƷ���� & strҩ��
                    .TextMatrix(intRow, mconIntColҩƷ����) = rsInitCard!ҩƷ����
                    .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    Else
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(rsInitCard!��Ʒ��), "", rsInitCard!��Ʒ��)
                    
                    .TextMatrix(intRow, mconIntCol��Դ) = zlStr.Nvl(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = zlStr.Nvl(rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntCol���) = rsInitCard!���
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol�ⷿ��λ) = IIf(IsNull(rsInitCard!�ⷿ��λ), "", rsInitCard!�ⷿ��λ)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mconIntcol�ӳ���) = zlStr.FormatEx(IIf(IsNull(rsInitCard!�ӳ���), 0, rsInitCard!�ӳ���), mintMoneyDigit, , True) / 100 & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol������) = IIf(IsNull(rsInitCard!������), "0", rsInitCard!������)
                    If rsInitCard!ʵ������ = 0 Then
                        mintNumberDigit = 5
                        mintNumberDigit0 = 5
                    End If
                    .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(rsInitCard!��������, mintNumberDigit, , True)
                    .TextMatrix(intRow, mconintColʵ������) = zlStr.FormatEx(rsInitCard!ʵ������, mintNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!�ۼ�, mintPriceDigit, , True)
                    .TextMatrix(intRow, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintColʵ������)) * Val(.TextMatrix(intRow, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    
                    If mintUnit > 0 Then
                        .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!�ɱ���, 0), mintCostDigit, , True)
                    Else
                        .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!�ɱ���, 0), mintCostDigit0, , True)
                    End If
                    
                    If mintUnit > 0 Then
                        .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                        .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                        .TextMatrix(intRow, mconintCol������) = zlStr.FormatEx(rsInitCard!������, mintNumberDigit, , True)
                    Else
                        Select Case mint��λ
'                            Case mconint�ۼ۵�λ
'                                .TextMatrix(intRow, mconIntCol����������λ��) = rsintcard!�ۼ۵�λ
'                                .TextMatrix(intRow, mconIntCol�̵�������λ��) = rsintcard!�ۼ۵�λ
'                                .TextMatrix(intRow, mconIntCol����ϵ����) = rsInitCard!����ϵ���ۼ�
'                                .TextMatrix(intRow, mconintCol���װ��������) =Str.FormatEx(rsInitCard!��������, mintNumberDigit)
'                                .TextMatrix(intRow, mconintCol���װʵ������) =Str.FormatEx(rsInitCard!ʵ������, mintNumberDigit)
                            Case mconint���ﵥλ
                                .TextMatrix(intRow, mconIntCol����������λ��) = rsInitCard!���ﵥλ
                                .TextMatrix(intRow, mconIntColʵ��������λ��) = rsInitCard!���ﵥλ
                                .TextMatrix(intRow, mconIntCol����ϵ����) = rsInitCard!����ϵ������
                                .TextMatrix(intRow, mconintCol���װ��������) = zlStr.FormatEx(Int(rsInitCard!�������� / rsInitCard!����ϵ������), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol���װʵ������) = zlStr.FormatEx(Int(rsInitCard!ʵ������ / rsInitCard!����ϵ������), mintNumberDigit0, , True)
                            Case mconintסԺ��λ
                                .TextMatrix(intRow, mconIntCol����������λ��) = rsInitCard!סԺ��λ
                                .TextMatrix(intRow, mconIntColʵ��������λ��) = rsInitCard!סԺ��λ
                                .TextMatrix(intRow, mconIntCol����ϵ����) = rsInitCard!����ϵ��סԺ
                                .TextMatrix(intRow, mconintCol���װ��������) = zlStr.FormatEx(Int(rsInitCard!�������� / rsInitCard!����ϵ��סԺ), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol���װʵ������) = zlStr.FormatEx(Int(rsInitCard!ʵ������ / rsInitCard!����ϵ��סԺ), mintNumberDigit0, , True)
                            Case mconintҩ�ⵥλ
                                .TextMatrix(intRow, mconIntCol����������λ��) = rsInitCard!ҩ�ⵥλ
                                .TextMatrix(intRow, mconIntColʵ��������λ��) = rsInitCard!ҩ�ⵥλ
                                .TextMatrix(intRow, mconIntCol����ϵ����) = rsInitCard!����ϵ��ҩ��
                                .TextMatrix(intRow, mconintCol���װ��������) = zlStr.FormatEx(Int(rsInitCard!�������� / rsInitCard!����ϵ��ҩ��), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol���װʵ������) = zlStr.FormatEx(Int(rsInitCard!ʵ������ / rsInitCard!����ϵ��ҩ��), mintNumberDigit0, , True)
                        End Select
                        
                        Select Case mintС��λ
                            Case mconint�ۼ۵�λ
                                .TextMatrix(intRow, mconIntCol����������λС) = rsInitCard!�ۼ۵�λ
                                .TextMatrix(intRow, mconIntColʵ��������λС) = rsInitCard!�ۼ۵�λ
                                .TextMatrix(intRow, mconIntCol����ϵ��С) = rsInitCard!����ϵ���ۼ�
                                .TextMatrix(intRow, mconintColС��װ��������) = zlStr.FormatEx(Val(rsInitCard!��������) - Val(.TextMatrix(intRow, mconintCol���װ��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ����)), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintColС��װʵ������) = zlStr.FormatEx(Val(rsInitCard!ʵ������) - Val(.TextMatrix(intRow, mconintCol���װʵ������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ����)), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol������) = zlStr.FormatEx(rsInitCard!������, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!�ۼ� * rsInitCard!����ϵ���ۼ�, mintPriceDigit0, , True)
                                .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!�ɱ���, 0) * rsInitCard!����ϵ���ۼ�, mintCostDigit0, , True)
                                .TextMatrix(intRow, mconintCol�ϼ�) = .TextMatrix(intRow, mconintColʵ������) & rsInitCard!�ۼ۵�λ
                            Case mconint���ﵥλ
                                .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(rsInitCard!�������� / rsInitCard!����ϵ������, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintColʵ������) = zlStr.FormatEx(rsInitCard!ʵ������ / rsInitCard!����ϵ������, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol�ϼ�) = .TextMatrix(intRow, mconintColʵ������) & rsInitCard!���ﵥλ
                                .TextMatrix(intRow, mconIntCol����������λС) = rsInitCard!���ﵥλ
                                .TextMatrix(intRow, mconIntColʵ��������λС) = rsInitCard!���ﵥλ
                                .TextMatrix(intRow, mconIntCol����ϵ��С) = rsInitCard!����ϵ������
                                .TextMatrix(intRow, mconintColС��װ��������) = zlStr.FormatEx((Val(rsInitCard!��������) - Val(.TextMatrix(intRow, mconintCol���װ��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ����))) / Val(.TextMatrix(intRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintColС��װʵ������) = zlStr.FormatEx((Val(rsInitCard!ʵ������) - Val(.TextMatrix(intRow, mconintCol���װʵ������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ����))) / Val(.TextMatrix(intRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol������) = zlStr.FormatEx(rsInitCard!������ / rsInitCard!����ϵ������, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!�ۼ� * rsInitCard!����ϵ������, mintPriceDigit0, , True)
                                .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!�ɱ���, 0) * rsInitCard!����ϵ������, mintCostDigit0, , True)
                            Case mconintסԺ��λ
                                .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(rsInitCard!�������� / rsInitCard!����ϵ��סԺ, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintColʵ������) = zlStr.FormatEx(rsInitCard!ʵ������ / rsInitCard!����ϵ��סԺ, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol�ϼ�) = .TextMatrix(intRow, mconintColʵ������) & rsInitCard!סԺ��λ
                                .TextMatrix(intRow, mconIntCol����������λС) = rsInitCard!סԺ��λ
                                .TextMatrix(intRow, mconIntColʵ��������λС) = rsInitCard!סԺ��λ
                                .TextMatrix(intRow, mconIntCol����ϵ��С) = rsInitCard!����ϵ��סԺ
                                .TextMatrix(intRow, mconintColС��װ��������) = zlStr.FormatEx((Val(rsInitCard!��������) - Val(.TextMatrix(intRow, mconintCol���װ��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ����))) / Val(.TextMatrix(intRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintColС��װʵ������) = zlStr.FormatEx((Val(rsInitCard!ʵ������) - Val(.TextMatrix(intRow, mconintCol���װʵ������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ����))) / Val(.TextMatrix(intRow, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol������) = zlStr.FormatEx(rsInitCard!������ / rsInitCard!����ϵ��סԺ, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!�ۼ� * rsInitCard!����ϵ��סԺ, mintPriceDigit0, , True)
                                .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!�ɱ���, 0) * rsInitCard!����ϵ��סԺ, mintCostDigit0, , True)
                            Case mconintҩ�ⵥλ
                                .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(rsInitCard!�������� / rsInitCard!����ϵ��ҩ��, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintColʵ������) = zlStr.FormatEx(rsInitCard!ʵ������ / rsInitCard!����ϵ��ҩ��, mintNumberDigit0, , True)
                                .TextMatrix(intRow, mconintCol�ϼ�) = .TextMatrix(intRow, mconintColʵ������) & rsInitCard!ҩ�ⵥλ
'                                .TextMatrix(intRow, mconIntCol����������λ��) = rsintcard!ҩ�ⵥλ
'                                .TextMatrix(intRow, mconIntCol�̵�������λ��) = rsintcard!ҩ�ⵥλ
'                                .TextMatrix(intRow, mconIntCol����ϵ����) = rsInitCard!����ϵ��ҩ��
'                                .TextMatrix(intRow, mconintCol���װ��������) =Str.FormatEx(Int(rsInitCard!�������� / rsInitCard!����ϵ��ҩ��), mintNumberDigit)
'                                .TextMatrix(intRow, mconintCol���װʵ������) =Str.FormatEx(Int(rsInitCard!ʵ������ / rsInitCard!����ϵ��ҩ��), mintNumberDigit)
                        End Select
                    End If
                    
                    If rsInitCard!ʵ������ > rsInitCard!�������� Then
                        .TextMatrix(intRow, mconintCol��־) = "ӯ"
                    ElseIf rsInitCard!ʵ������ < rsInitCard!�������� Then
                        .TextMatrix(intRow, mconintCol��־) = "��"
                    Else
                        .TextMatrix(intRow, mconintCol��־) = "ƽ"
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '���ҩƷ���������Ϊ0�������۲�Ϊ0��ҩƷ�޷�ͨ���̵��������¼������
                    '��������µ�ͨ��ҩƷ�������۵�ʵ��λ������ϵͳ���������õĽ��λ��
                    '����취�����ʵ������Ϊ0�������Ͳ�۲�С��λ�����ֺ�ҩƷ�����н��Ͳ��λ��һ��
                    If Val(.TextMatrix(intRow, mconintColʵ������)) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True And Val(.TextMatrix(intRow, mconIntCol�ۼ�)) = Val(.TextMatrix(intRow, mconintCol�ɱ���))) Then
                        intMoneyBit = mintMaxMoneyBit
                    Else
                        intMoneyBit = mintMoneyDigit
                    End If
                    .TextMatrix(intRow, mconIntColʵ�ʲ��) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!�����, 0), intMoneyBit, , True)
                    .TextMatrix(intRow, mconIntColʵ�ʽ��) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!�����, 0), intMoneyBit, , True)
                    .TextMatrix(intRow, mconintCol����) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!����, 0), intMoneyBit, , True)
                    .TextMatrix(intRow, mconintCol��۲�) = zlStr.FormatEx(zlStr.Nvl(rsInitCard!��۲�, 0), intMoneyBit, , True)
                    '���������������Ͳ�۲��㷨һ��
                    dbl���� = Val(.TextMatrix(intRow, mconintCol����)) * rsInitCard!���ϵ�� * IIf(mint��¼״̬ = 1, 1, IIf(mint��¼״̬ Mod 3 = 0, 1, -1))
                    dbl��۲� = Val(.TextMatrix(intRow, mconintCol��۲�)) * rsInitCard!���ϵ�� * IIf(mint��¼״̬ = 1, 1, IIf(mint��¼״̬ Mod 3 = 0, 1, -1))
                    
                    '.TextMatrix(intRow, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol�ɱ���)) * Val(.TextMatrix(intRow, mconintColʵ������)), mintMoneyDigit)
                    '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                    .TextMatrix(intRow, mconintCol�̵�ɱ����) = zlStr.FormatEx((zlStr.Nvl(rsInitCard!�����, 0) + dbl����) - (zlStr.Nvl(rsInitCard!�����, 0) + dbl��۲�), mintMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol����)) - Val(.TextMatrix(intRow, mconintCol��۲�)), mintMoneyDigit, , True)
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .TextMatrix(intRow, mconintCol�������) = zlStr.Nvl(rsInitCard!��д����, 0)
                    
                    '���÷�������
                    Call GetҩƷ��������(intRow)
                                        
                    .Row = intRow
                    
                    '�̿���ӯ������ɫ����
                    Call SetStocktakingColor(vsfBill, intRow)
                   
                    '�����ͣ��ҩƷ�����д�����ʾ
                    If Format(rsInitCard!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                
                If mintUnit > 0 Then
                    .Cell(flexcpFontBold, 1, mconintColʵ������, .rows - 1, mconintColʵ������) = True
                Else
                    .Cell(flexcpFontBold, 1, mconintCol���װʵ������, .rows - 1, mconintCol���װʵ������) = True
                    .Cell(flexcpFontBold, 1, mconintColС��װʵ������, .rows - 1, mconintColС��װʵ������) = True
                End If
                
                Call SetSortCode
                
                .Redraw = flexRDDirect
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    mblnLoadData = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ʼ���༭�ؼ�
Private Sub initGrid()
    Dim i As Integer
    
    With vsfBill
        .Redraw = flexRDNone
        .Cols = mconIntColS
        .SelectionMode = flexSelectionByRow
        .Editable = flexEDNone
        .RowHeightMax = 315
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ⷿ��λ) = "�ⷿ��λ"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        
        .TextMatrix(0, mconIntCol����ϵ����) = "����ϵ����"
        .TextMatrix(0, mconIntCol����ϵ��С) = "����ϵ��С"
        
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        
        .TextMatrix(0, mconintCol��������) = "��������"
        
        .TextMatrix(0, mconintCol���װ��������) = "���װ��������"
        .TextMatrix(0, mconIntCol����������λ��) = "��λ"
        
        .TextMatrix(0, mconintColС��װ��������) = "С��װ��������"
        .TextMatrix(0, mconIntCol����������λС) = "��λ"
        
        .TextMatrix(0, mconintColʵ������) = "ʵ������"
                
        .TextMatrix(0, mconintCol���װʵ������) = "���װʵ������"
        .TextMatrix(0, mconIntColʵ��������λ��) = "��λ"
        
        .TextMatrix(0, mconintColС��װʵ������) = "С��װʵ������"
        .TextMatrix(0, mconIntColʵ��������λС) = "��λ"
        
        .TextMatrix(0, mconintCol�ϼ�) = "�ϼ�"
        .TextMatrix(0, mconintCol��־) = "��־"
        .TextMatrix(0, mconintCol������) = "������"
        .TextMatrix(0, mconintCol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconintCol����) = "����"
        .TextMatrix(0, mconintCol��۲�) = "��۲�"
        .TextMatrix(0, mconintCol�̵���) = "�̵���"
        .TextMatrix(0, mconintCol�̵�ɱ����) = "�̵�ɱ����"
        .TextMatrix(0, mconintCol�̵�ɱ�����) = "�̵�ɱ�����"
        .TextMatrix(0, mconintCol�������) = "�������"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol������) = "������"
        .TextMatrix(0, mconIntCol�������) = "�������"
        .TextMatrix(0, mconIntCol��������) = "��������"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol��������) = 0
        
        .ColWidth(mconIntCol����ϵ��) = 0
        
        .ColWidth(mconIntCol����ϵ����) = 0
        .ColWidth(mconIntCol����ϵ��С) = 0
        
        .ColWidth(mconIntcol�ӳ���) = 0
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol�ⷿ��λ) = 2000
        .ColWidth(mconIntCol��λ) = IIf(mintUnit = 0, 0, 600)
        
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        
        .ColWidth(mconintCol��������) = IIf(mintUnit = 0, 0, 1200)
        
        .ColWidth(mconintCol���װ��������) = IIf(mintUnit = 0, 1400, 0)
        .ColWidth(mconIntCol����������λ��) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintColС��װ��������) = IIf(mintUnit = 0, 1400, 0)
        .ColWidth(mconIntCol����������λС) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintColʵ������) = IIf(mintUnit = 0, 0, 1200)
        
        .ColWidth(mconintCol���װʵ������) = IIf(mintUnit = 0, 1400, 0)
        .ColWidth(mconIntColʵ��������λ��) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintColС��װʵ������) = IIf(mintUnit = 0, 1400, 0)
        .ColWidth(mconIntColʵ��������λС) = IIf(mintUnit = 0, 600, 0)
        
        .ColWidth(mconintCol�ϼ�) = IIf(mintUnit = 0, 1000, 0)
        .ColWidth(mconintCol��־) = 500
        .ColWidth(mconintCol������) = 800
        .ColWidth(mconintCol�ɱ���) = 900
        .ColWidth(mconIntCol�ۼ�) = 900
        .ColWidth(mconintCol����) = 900
        .ColWidth(mconintCol��۲�) = 900
        .ColWidth(mconintCol�̵���) = 900
        .ColWidth(mconintCol�̵�ɱ����) = 1400
        .ColWidth(mconintCol�̵�ɱ�����) = 1500
        .ColWidth(mconintCol�������) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntCol������) = 0
        .ColWidth(mconIntCol�������) = 0
        .ColWidth(mconIntCol��������) = 0
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconintCol��������) = flexAlignRightCenter
        .ColAlignment(mconintCol���װ��������) = flexAlignRightCenter
        .ColAlignment(mconintColС��װ��������) = flexAlignRightCenter
        .ColAlignment(mconIntCol����������λ��) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����������λС) = flexAlignCenterCenter
        .ColAlignment(mconintColʵ������) = flexAlignRightCenter
        .ColAlignment(mconintCol���װʵ������) = flexAlignRightCenter
        .ColAlignment(mconintColС��װʵ������) = flexAlignRightCenter
        .ColAlignment(mconIntColʵ��������λ��) = flexAlignCenterCenter
        .ColAlignment(mconIntColʵ��������λС) = flexAlignCenterCenter
        
        .ColAlignment(mconintCol�ϼ�) = flexAlignRightCenter
        .ColAlignment(mconintCol��־) = flexAlignCenterCenter
        .ColAlignment(mconintCol������) = flexAlignRightCenter
        .ColAlignment(mconintCol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconintCol����) = flexAlignRightCenter
        .ColAlignment(mconintCol��۲�) = flexAlignRightCenter
        .ColAlignment(mconintCol�̵���) = flexAlignRightCenter
        .ColAlignment(mconintCol�̵�ɱ����) = flexAlignRightCenter
        .ColAlignment(mconintCol�̵�ɱ�����) = flexAlignRightCenter
        

        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Or mint�༭״̬ = 6 Then
            txtժҪ.Enabled = True
        Else
            txtժҪ.Enabled = False
        End If
        
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        
        .Redraw = flexRDDirect
    End With
    txtժҪ.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
    
    '�ָ����Ի����ã��������в���Ӱ��
    RestoreWinState Me, App.ProductName, MStrCaption
    
    'Ȩ�޿��Ƶģ��ڸ��Ի��ָ�����Ҫ��һ������
    vsfBill.ColWidth(mconintCol�ɱ���) = IIf(mblnViewCost = True, 900, 0)
    vsfBill.ColWidth(mconintCol��۲�) = IIf(mblnViewCost = True, 900, 0)
    vsfBill.ColWidth(mconintCol�̵�ɱ����) = IIf(mblnViewCost = True, 1400, 0)
    vsfBill.ColWidth(mconintCol�̵�ɱ�����) = IIf(mblnViewCost = True, 1400, 0)
    
    vsfBill.ColWidth(mconIntCol��λ) = IIf(mintUnit = 0, 0, 600)
    vsfBill.ColWidth(mconintCol��������) = IIf(mintUnit = 0, 0, 1200)
    vsfBill.ColWidth(mconintCol���װ��������) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntCol����������λ��) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintColС��װ��������) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntCol����������λС) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintColʵ������) = IIf(mintUnit = 0, 0, 1200)
    vsfBill.ColWidth(mconintCol���װʵ������) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntColʵ��������λ��) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintColС��װʵ������) = IIf(mintUnit = 0, 1400, 0)
    vsfBill.ColWidth(mconIntColʵ��������λС) = IIf(mintUnit = 0, 600, 0)
    vsfBill.ColWidth(mconintCol�ϼ�) = IIf(mintUnit = 0, 1000, 0)
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        vsfBill.ColWidth(mconIntCol��Ʒ��) = IIf(vsfBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, vsfBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        vsfBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 600
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With
    
    With vsfBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNo
        .Left = vsfBill.Left + vsfBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    txtCheckDate.Left = vsfBill.Left + vsfBill.Width - txtCheckDate.Width
    lblCheckDate.Left = txtCheckDate.Left - lblCheckDate.Width - 100
    
    LblStock.Left = vsfBill.Left
    txtStock.Left = LblStock.Left + LblStock.Width + 100
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = vsfBill.Left + 100
    End With
    
    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With
    
    With Lbl��������
        .Top = Lbl������.Top
        .Left = Txt������.Left + Txt������.Width + 250
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Txt�������
        .Top = Lbl������.Top - 80
        .Left = vsfBill.Left + vsfBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = Lbl�������.Left - 200 - .Width
    End With
    
    With Lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = vsfBill.Left + vsfBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With
    
    With lblPurchasePrice
        .Left = vsfBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = Pic����.TextWidth(.Caption) + 200
        
        lblCheckSum.Left = .Left + .Width + 100
        lblCheckSum.Top = .Top
        lblCheckSum.Width = Pic����.TextWidth(lblCheckSum.Caption) + 200
        
'        lblCheckMoney.Left = lblCheckSum.Left + .Width + 100
'        lblCheckMoney.Top = lblCheckSum.Top
'        lblCheckMoney.Width = Pic����.TextWidth(lblCheckSum.Caption) + 200
'
'        lblPricePur.Left = lblCheckMoney.Left + .Width + 100
'        lblPricePur.Top = lblCheckMoney.Top
'        lblPricePur.Width = Pic����.TextWidth(lblCheckSum.Caption) + 200
    End With
    
    With lblCostPrice
        .Top = lblCheckSum.Top
        .Left = lblCheckSum.Left + lblCheckSum.Width + 200
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblCostPrice.Visible = False
    End If
    
    With vsfBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic����.Left + vsfBill.Left + vsfBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 550
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + vsfBill.Left
        .Top = Pic����.Top + Pic����.Height + 130
    End With
        
    With lblCode
        .Top = Pic����.Top + Pic����.Height + 220
    End With
    With txtCode
        .Top = Pic����.Top + Pic����.Height + 150
    End With
    
    With cmd�̶���
        .Left = CmdSave.Left - .Width - 150
        .Top = Pic����.Top + Pic����.Height + 550
    End With
    
    With cmdBatch
        .Left = cmdHelp.Left
        .Top = cmdHelp.Top + cmdHelp.Height + 100
    End With
    
    With cmdReSet
        .Left = cmdBatch.Left + cmdBatch.Width + 150
        .Top = cmdBatch.Top
    End With
       
    If mint�༭״̬ = 1 Then
        cmdBatch.Visible = True
        cmdReSet.Visible = True
    ElseIf mint�༭״̬ = 5 Then
        cmdBatch.Visible = False
        cmdReSet.Visible = True
        cmdReSet.Left = cmdBatch.Left
    Else
        cmdBatch.Visible = False
        cmdReSet.Visible = False
        Pic����.Height = Pic����.Height + 200
        cmdHelp.Top = cmdHelp.Top + 250
        lblCode.Top = lblCode.Top + 240
        txtCode.Top = txtCode.Top + 250
        CmdCancel.Top = txtCode.Top
        CmdSave.Top = txtCode.Top
    End If
        
    cmdSet.Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
    cmdZero.Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
    
    If cmdSet.Visible = True Then
        With cmdSet
            .Left = IIf(cmdReSet.Visible, cmdReSet.Left + cmdReSet.Width + 150, txtCode.Left + txtCode.Width + 150)
            .Top = IIf(cmdReSet.Visible, cmdReSet.Top, txtCode.Top)
        End With
    End If
    
    If cmdZero.Visible = True Then
        With cmdZero
            .Left = IIf(cmdReSet.Visible, cmdReSet.Left + cmdReSet.Width + 150, txtCode.Left + txtCode.Width + 150) + cmdSet.Width + 30
            .Top = IIf(cmdReSet.Visible, cmdReSet.Top, txtCode.Top)
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�̵����", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    mbln���䶯 = False
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        vsfBill.SetFocus
        vsfBill.Row = 1
        vsfBill.Col = mconIntColҩ��
        If txtCheckDate.Caption = "" Then txtCheckDate.Caption = Format(Sys.Currentdate, "yyyy-MM-dd HH:mm:ss")
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
End Sub

Private Function SaveCheck() As Boolean
    Dim strNo As String
    Dim str����� As String
    
    mblnSave = False
    SaveCheck = False
    
    str����� = UserInfo.�û�����
    strNo = txtNo.Tag
    On Error GoTo ErrHandle
    
    gstrSQL = "zl_ҩƷ�̵�_Verify('" & strNo & "','" & str����� & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
        
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function



Private Sub mnuColDrug_Click(index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(index).Checked = True
        
        Call SetDrugName(index)
    End With
End Sub

Private Sub SetDrugName(ByVal intType As Integer)
    'ҩƷ������ʾ��
    'intType��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With vsfBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntColҩ��) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                Else
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ���������)
                End If
            End If
        Next
    End With
End Sub
Private Sub mnuDefault_Click()
    vsfBill.FixedCols = 1
End Sub

Private Sub mnuFirst_Click()
    vsfBill.Redraw = flexRDNone
    vsfBill.FixedCols = mconIntCol��λ
    vsfBill.Refresh
    vsfBill.Redraw = flexRDDirect
End Sub

Private Sub mnuSecond_Click()
    vsfBill.Redraw = flexRDNone
    vsfBill.FixedCols = mconIntColЧ��
    vsfBill.Refresh
    vsfBill.Redraw = flexRDDirect
End Sub

Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtStock_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        Call SetSelectorRS(2, "ҩƷ�̵����", txtStock.Tag, txtStock.Tag, , , , mbln��ͣ��ҩƷ, mblnNoStock, 1, , , mbln���Է������)
    End If
End Sub

Private Sub vsfBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBill
        Select Case Col
            Case mconIntColҩ��
                .ColComboList(mconIntColҩ��) = "..."
        End Select
    End With
End Sub

Private Sub vsfBill_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    Dim lngColor As Long
    
    With vsfBill
        If NewRowSel > 0 And NewRowSel <> OldRowSel Then
            If .TextMatrix(NewRowSel, mconintCol��־) = "ƽ" Then
                lngColor = mlngColor_��ƽ
            ElseIf .TextMatrix(NewRowSel, mconintCol��־) = "ӯ" Then
                lngColor = mlngColor_��ӯ
            ElseIf .TextMatrix(NewRowSel, mconintCol��־) = "��" Then
                lngColor = mlngColor_�̿�
            End If
            
            .ForeColorSel = lngColor
        End If
    End With
End Sub

Private Sub vsfBill_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfBill
        If Col = mconIntColҩ�� Then
            .Col = mconIntCol�������
            .Sort = Order
        End If
    End With
    
    Call RefreshListSN
End Sub

Private Sub vsfBill_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    If Button = 1 Then
        If y <= vsfBill.RowHeight(0) Then '�������ͷʱ������ͷ��ʼ���²�ѯ
            mlngFindCurrRow = 1
            If Not mrsFindName Is Nothing Then
                mrsFindName.MoveFirst
            End If
        End If
    End If
End Sub

Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim rsProvider As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    
    intOldRow = vsfBill.Row
    With vsfBill
        Select Case Col
        Case mconIntColҩ��
            If mblnNotTrigger <> True Then
                mblnNotTrigger = True
                
                If grsMaster.State = adStateClosed Then
                    Call SetSelectorRS(2, "ҩƷ�̵����", txtStock.Tag, txtStock.Tag, , , , mbln��ͣ��ҩƷ, mblnNoStock, 1, , , mbln���Է������)
                End If
                
                Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , txtStock.Tag, txtStock.Tag, , 0, False, True, True, IIf(mbln��ͣ��ҩƷ, 1, 0))
                If RecReturn.RecordCount > 0 Then
                    Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                End If
                
                mblnNotTrigger = False
            Else
                Exit Sub
            End If
        
            '��"FrmҩƷѡ����"�еĴ�����ִ����
            DoEvents
                            
            If RecReturn.RecordCount > 0 Then
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    intCurRow = .Row
                    Call SetPhiscRows(RecReturn!ҩƷid, IIf(IsNull(RecReturn!����), 0, RecReturn!����), Val(RecReturn!�ɱ���), IIf(mintUnit > 0, Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), 0), _
                            IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�))
                    
                    vsfBill_MoveNextCell Row, Col
                    
                    If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If
                    .Row = .rows - 1
                    RecReturn.MoveNext
                Next
                .Row = intOldRow
            End If
        Case mconIntCol����
            vRect = zlControl.GetControlRect(vsfBill.hWnd)
            dblLeft = vRect.Left + vsfBill.CellLeft
            dblTop = vRect.Top + vsfBill.CellTop
            
            gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
            Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
            True, dblLeft, dblTop, 300, blnCancel, False, True, gstrNodeNo)
            
            If rsProvider Is Nothing Then
                Exit Sub
            End If
            If Not rsProvider.EOF Then
                .TextMatrix(.Row, mconIntCol����) = rsProvider!����
            End If
        End Select
    End With
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ���ʱ��ҩƷ�Ƿ��п��

    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim rs����ʱ�� As ADODB.Recordset
    Dim str��� As String
    Dim strSQL As String
    Dim strDub As String    '�ظ�ҩƷ
    Dim str�ظ�ҩ�� As String
    Dim strNotPrice As String  '�޼۸�ҩƷ
    Dim strNotPriceҩ�� As String   '������¼�ظ�ѡ���˵�ҩƷ����
    Dim strPriceҩ�� As String
    Dim rsDetail As ADODB.Recordset
    Dim str�̵�ʱ�� As String
    Dim str�̵�ʱ���ҩƷ As String       '��¼���̵�ʱ�������ҩƷ
    Dim strSql�̵� As String   '�����̵�ʱ�������ҩƷ
    
    rsTemp.MoveFirst
    str�̵�ʱ���ҩƷ = ""
    strSql�̵� = ""
    str���� = ""
    strTemp = ""
    str�̵�ʱ�� = txtCheckDate.Caption
    
    On Error GoTo ErrHandle
    Do While Not rsTemp.EOF
        str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        If InStr(1, strTemp, rsTemp!ҩƷid & "," & str����) = 0 Then
            If Val(str����) <> -1 Then strTemp = strTemp & rsTemp!ҩƷid & "," & str���� & "," & rsTemp!ͨ���� & "|"
        End If
        
        gstrSQL = "select �ּ� from �շѼ�Ŀ where ִ������(+)<=[1] AND NVL(��ֹ����(+),SYSDATE)>=[1] and �շ�ϸĿid=[2]" & _
                GetPriceClassString("")
        
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�ּ�", CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")), rsTemp!ҩƷid)
        If Not rsDetail.EOF Then
            If IsNull(rsDetail!�ּ�) Then
                strNotPrice = strNotPrice & rsTemp!ҩƷid & "," & rsTemp!ͨ���� & "|"
            End If
        End If
        
        gstrSQL = "Select a.����ʱ�� From �շ���ĿĿ¼ A Where a.Id =[1]"
        Set rs����ʱ�� = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ����ʱ��", rsTemp!ҩƷid)
        If Format(rs����ʱ��!����ʱ��, "yyyy-MM-dd HH:mm:ss") > Format(txtCheckDate.Caption, "yyyy-MM-dd HH:mm:ss") Then
            str�̵�ʱ���ҩƷ = str�̵�ʱ���ҩƷ & ";" & "[" & rsTemp!ҩƷ���� & "]" & rsTemp!ͨ����
            strSql�̵� = strSql�̵� & "ҩƷid<>" & rsTemp!ҩƷid & " and "
        End If
        
        rsTemp.MoveNext
    Loop
           
    If strSql�̵� <> "" Then
        MsgBox Mid(str�̵�ʱ���ҩƷ, 2) & vbCrLf & "����ҩƷΪ�̵�ʱ����������Բ��ᱻ��ӣ�", vbInformation, gstrSysName
        rsTemp.Filter = Mid(strSql�̵�, 1, Len(strSql�̵�) - 4)
    End If
    
    With vsfBill    '���ظ��Ĳ�ѯ����
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol����)) > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
            End If
        Next
        
        If strInfo <> "" Then   'Ϊ��������ƴ��sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "ҩƷid<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str�ظ�ҩ��, ",")) <= 2 Then
                    str�ظ�ҩ�� = str�ظ�ҩ�� & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        If strNotPrice <> "" Then
            strPriceҩ�� = ""
            For i = 0 To UBound(Split(strNotPrice, "|")) - 1
                strPriceҩ�� = strPriceҩ�� & "ҩƷid<>" & Split(Split(strNotPrice, "|")(i), ",")(0) & " and "
                If UBound(Split(strNotPriceҩ��, ",")) <= 2 Then
                    strNotPriceҩ�� = strNotPriceҩ�� & Split(Split(strNotPrice, "|")(i), ",")(1) & ","
                End If
            Next
            If strPriceҩ�� <> "" Then
                strPriceҩ�� = Mid(strPriceҩ��, 1, Len(strPriceҩ��) - 4)
            End If
        End If
        '�ж���ʲô��ʽƴ��sql
        
        If str�ظ�ҩ�� <> "" And strNotPriceҩ�� <> "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & strNotPriceҩ�� & "�ڱ����̵�ʱ��ʱ���ۼ���Ϣ��" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strDub & " and " & strPriceҩ��
        End If
        If str�ظ�ҩ�� <> "" And strNotPriceҩ�� = "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If str�ظ�ҩ�� = "" And strNotPriceҩ�� <> "" Then
            MsgBox strNotPriceҩ�� & "�ڱ����̵�ʱ��ʱ���ۼ���Ϣ��" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strPriceҩ��
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsfBill_EnterCell()
    Dim lng����  As Long
    Dim bln������ As Boolean
        
    With vsfBill
        .Editable = flexEDNone
        
        Select Case .Col
            Case mconIntColҩ��
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntColҩ��) = "..."
                End If
                
            Case mconIntCol����
                .EditMaxLength = mintBatchNoLen
                
                lng���� = Val(.TextMatrix(.Row, mconIntCol����))
                bln������ = (Val(.TextMatrix(.Row, mconIntCol������)) = 1 And (mint�༭״̬ = 1 Or mint�༭״̬ = 2))
                
                If IIf(lng���� = -1 Or bln������ = True, 1, 0) = 1 Then
                    .Editable = flexEDKbdMouse
                End If
            Case mconIntCol����
                lng���� = Val(.TextMatrix(.Row, mconIntCol����))
                bln������ = (Val(.TextMatrix(.Row, mconIntCol������)) = 1 And (mint�༭״̬ = 1 Or mint�༭״̬ = 2))
                
                If IIf(lng���� = -1 Or bln������ = True, 1, 0) = 1 Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(mconIntCol����) = "..."
                End If
            Case mconIntColЧ��
                .EditMaxLength = 10
                
                lng���� = Val(.TextMatrix(.Row, mconIntCol����))
                bln������ = (Val(.TextMatrix(.Row, mconIntCol������)) = 1 And (mint�༭״̬ = 1 Or mint�༭״̬ = 2))
                
                If IIf(lng���� = -1 Or bln������ = True, 1, 0) = 1 Then
                    .Editable = flexEDKbdMouse
                End If
                 
                If .TextMatrix(.Row, mconIntCol����) <> "" And .TextMatrix(.Row, mconIntColЧ��) = "" Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                        If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                            strxq = TranNumToDate(strxq)
                            If strxq = "" Then Exit Sub
                            
                            .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", .RowData(.Row), strxq), "yyyy-mm-dd")
                            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                                '����Ϊ��Ч��
                                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mconintColʵ������, mconintCol���װʵ������, mconintColС��װʵ������
                .EditMaxLength = 16
                If Val(.TextMatrix(.Row, 0)) <> 0 Then
                    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                        If (.Col = mconintColʵ������ And mintUnit > 0) Or ((.Col = mconintCol���װʵ������ Or .Col = mconintColС��װʵ������) And mintUnit = 0) Then
                            .Editable = flexEDKbdMouse
                        End If
                    End If
                End If
            Case mconintCol�ɱ���
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
                    If Val(.TextMatrix(.Row, mconintCol��������)) = 0 Then
                       .Editable = flexEDKbdMouse
                    End If
                End If
        End Select
        
        If mlongCurrRow <> .Row Then
            mlongCurrRow = .Row
            Call ��ʾ�ϼƽ��
            Call ��ʾ�����
        End If
    End With
End Sub

Private Sub vsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBill
        If KeyCode = vbKeyDelete Then
            If .rows = 2 Then Exit Sub
            If .TextMatrix(.Row, mconIntCol�к�) = "" Then Exit Sub
            If InStr(1, "3456", mint�༭״̬) <> 0 Then Exit Sub
            
            If MsgBox("�Ƿ�ɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .RemoveItem .Row
                Call RefreshRowNO(vsfBill, mconIntCol�к�, .Row)
            End If
        End If
        
        If txtCode.Visible And KeyCode = vbKeyF3 Then
            Call txtCode_KeyPress(13)
        End If
        
        Select Case .Col
            Case mconIntColҩ��
                If KeyCode <> vbKeyReturn Then
                    .ColComboList(mconIntColҩ��) = ""
                ElseIf .EditText = "" Then
'                    mblnNotTrigger = True
                    If .TextMatrix(.Row, mconIntColҩ��) = "" Then
                        txtժҪ.SetFocus
                    End If
                End If
            Case mconIntCol����
                If KeyCode <> vbKeyReturn Then
                    .ColComboList(mconIntCol����) = ""
                End If
        End Select
    End With
End Sub

Private Sub vsfBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim strKey As String
    Dim strTmp As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim rsProvider As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblTop, dblLeft As Double
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    intOldRow = vsfBill.Row
    With vsfBill
        .Redraw = flexRDNone
        
        .EditText = Trim(.EditText)
        strKey = Trim(.EditText)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        
        Select Case Col
            Case mconIntColҩ��
                strTmp = .TextMatrix(Row, Col)
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + vsfBill.Left + vsfBill.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + vsfBill.Top + vsfBill.CellTop + vsfBill.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - vsfBill.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, txtStock.Tag, txtStock.Tag, , strkey, sngLeft, sngTop, False, True, True, True, True, 0, mblnNoStock, 0, mbln��ͣ��ҩƷ, mbln���Է������)
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, "ҩƷ�̵����", txtStock.Tag, txtStock.Tag, , , , mbln��ͣ��ҩƷ, mblnNoStock, 1, , , mbln���Է������)
                    End If
                    
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strKey, sngLeft, sngTop, txtStock.Tag, txtStock.Tag, , 0, False, True, True, IIf(mbln��ͣ��ҩƷ, 1, 0))
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                    End If
                    '��"FrmҩƷ��ѡѡ����"�еĴ�����ִ����
                    DoEvents
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            Call SetPhiscRows(RecReturn!ҩƷid, IIf(IsNull(RecReturn!����), 0, RecReturn!����), Val(RecReturn!�ɱ���), IIf(mintUnit > 0, Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), 0), IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�))
                            
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    End If

                    Call ��ʾ�����
                End If
            Case mconIntCol����
                vRect = zlControl.GetControlRect(vsfBill.hWnd)
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top + vsfBill.CellTop
                
                gstrSQL = "Select ���� as id,����,���� From ҩƷ������ " _
                            & "Where (վ�� = [3] Or վ�� is Null) And (upper(����) like [1] or Upper(����) like [1] or Upper(����) like [2]) Order By ����"
                
                Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                True, dblLeft, dblTop, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", strKey & "%", gstrNodeNo)
                
                If rsProvider Is Nothing Then
                    .EditText = ""
                    .TextMatrix(.Row, .Col) = ""
                    Exit Sub
                End If
                If Not rsProvider.EOF Then
                    .TextMatrix(.Row, mconIntCol����) = rsProvider!����
                    .EditText = rsProvider!����
                End If
        End Select
        
        vsfBill_MoveNextCell vsfBill.Row, vsfBill.Col
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        vsfBill_MoveNextCell vsfBill.Row, vsfBill.Col
    End If
End Sub

Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    If InStr(" '", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    If KeyAscii = 13 Then
        mblnKeyPressReturn = True
    Else
        mblnKeyPressReturn = False
    End If
    
    With vsfBill
        Select Case Col
            Case mconintColʵ������, mconintCol���װʵ������, mconintColС��װʵ������
                If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                ElseIf KeyAscii = Asc(".") Then
                    If InStr(.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                    End If
                End If
                
                strKey = .EditText
                If strKey = "" Then
                    strKey = .TextMatrix(.Row, .Col)
                End If
                Select Case .Col
                    Case mconintColʵ������
                        intDigit = mintNumberDigit
                    Case mconintCol���װʵ������
                        intDigit = mintNumberDigit1
                    Case mconintColС��װʵ������
                        intDigit = mintNumberDigit0
                End Select
                
                If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                    If .EditSelLength = Len(strKey) Then Exit Sub
                    If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Case mconIntColЧ��
                If InStr("1234567890-" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
        End Select
    End With
End Sub

Private Sub vsfBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With vsfBill
            If .Col = mconIntColҩ�� Then
                If .Row < 1 Then Exit Sub
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    If KeyAscii = 13 Then
        Call FindGridRow(txtCode.Text)
    End If
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long
    
    '����ҩƷ
    On Error GoTo ErrHandle
    If strInput <> txtCode.Tag Then
        '��ʾ�µĲ���
        txtCode.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub
    
    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����
        Else
            strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        End If
        lngFindRow = vsfBill.FindRow(strҩ��, mlngFindCurrRow, CLng(mconIntColҩƷ���������), True, True)
        
        If lngFindRow > 0 Then '��ѯ�����ݺ���ƶ��µ���һ�У����������һ���Ƿ�����ͬ��ҩƷ
            vsfBill.Select lngFindRow, 1, lngFindRow, vsfBill.Cols - 1
            vsfBill.TopRow = lngFindRow
                        
            If lngFindRow < vsfBill.rows - 1 Then
                mlngFindCurrRow = lngFindRow + 1
            Else
                mlngFindCurrRow = 1
                mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
            End If
            Exit For
        Else
            mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
            mlngFindCurrRow = 1 '�����ӵ�һ�п�ʼ�Ƚ�����ҩƷ
        End If
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    Dim lngЧ�� As Long
    Dim dblδ��ҩ���� As Double
    Dim dbl����ϵ�� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim lngҩƷID As Long
    Dim str���� As String, str���� As String, dbl�ɱ��� As Double
    Dim intRow As Integer
    
    On Error GoTo ErrHandle
    With vsfBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, 0)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconintColʵ������))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ��ʵ������Ϊ���ˣ����飡", vbInformation, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintColʵ������
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintColʵ������)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ��ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintColʵ������
                        .EditCell
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconintCol����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�Ľ�����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintColʵ������
                        .EditCell
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintCol������)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ����������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        .Col = mconintColʵ������
                        .EditCell
                        Exit Function
                    End If
                    
                    '����ҩƷ����¼����غ�����
                    If Val(.TextMatrix(intLop, mconIntCol��������)) = 1 And Val(.TextMatrix(intLop, mconIntCol����)) = -1 And (.TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntCol����) = "") Then
                        MsgBox "��" & intLop & "�е�ҩƷ���������η���ҩƷ,������Ĳ��غ�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                        vsfBill.SetFocus
                        .Row = intLop
                        .TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol����) = "" Then
                            .Col = mconIntCol����
                        Else
                            .Col = mconIntCol����
                        End If
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol����)) = -1 Then
                        If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol����))), vbFromUnicode)) > mintBatchNoLen Then
                            MsgBox "��" & intLop & "��ҩƷ�����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                            .SetFocus
                            .Row = intLop
                            .TopRow = intLop
                            .Col = mconIntCol����
                            .EditCell
                            Exit Function
                        End If
                        
                        '�ж��Ƿ�ΪЧ��ҩƷ
                        gstrSQL = "Select Nvl(���Ч��,0) Ч�� From ҩƷ��� Where ҩƷID=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж��Ƿ�ΪЧ��ҩƷ]", Val(.TextMatrix(intLop, 0)))
                        
                        lngЧ�� = rsTemp!Ч��
                        If lngЧ�� <> 0 Then
                            If Val(.TextMatrix(intLop, mconintColʵ������)) <> 0 Then
                                If Trim(.TextMatrix(intLop, mconIntCol����)) = "" Or Trim(.TextMatrix(intLop, mconIntColЧ��)) = "" Then
                                    MsgBox "��" & intLop & "�е�ҩƷ��Ч��ҩƷ,����������ż�Ч��" & vbCrLf & "��Ϣ�������뵥���У�", vbInformation, gstrSysName
                                    vsfBill.SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    If .TextMatrix(intLop, mconIntCol����) = "" Then
                                        .Col = mconIntCol����
                                    Else
                                        .Col = mconIntColЧ��
                                    End If
                                    .EditCell
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol������)) = 0 Then
                        '���۹�������Ƿ���ڲ��������۵�ҩƷ
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And (Val(.TextMatrix(intLop, mconIntCol����)) >= 0 And Val(.TextMatrix(intLop, mconIntCol������)) = 0) Then
                            If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                                If CheckPriceAdjust(Val(.TextMatrix(intLop, 0)), Val(txtStock.Tag), Val(.TextMatrix(intLop, mconIntCol����))) = False Then
                                    MsgBox "��" & intLop & "��ҩƷ���������۹���������¼���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                                    .SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        '����ʱ
                        If .TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                            If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                                '��������۹����������ۼۺͳɱ��۹�ϵ
                                If Val(.TextMatrix(intLop, mconintCol�ɱ���)) <> Val(.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                                    MsgBox "��" & intLop & "��ҩƷ���������۹������̵������ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                                    .SetFocus
                                    .Row = intLop
                                    .TopRow = intLop
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                                        
                End If
            Next
            
            
            '������ҩƷ�������εĲ��أ������Ƿ��ظ�
            For intLop = 1 To .rows - 1
                If Val(.TextMatrix(intLop, mconIntCol����)) = -1 Or Val(.TextMatrix(intLop, mconIntCol������)) = 1 Then
                    lngҩƷID = Val(.TextMatrix(intLop, 0))
                    str���� = .TextMatrix(intLop, mconIntCol����)
                    str���� = .TextMatrix(intLop, mconIntCol����)
                    dbl�ɱ��� = Val(.TextMatrix(intLop, mconintCol�ɱ���))
                    
                    For intRow = 1 To .rows - 1
                        If intLop <> intRow And _
                            lngҩƷID = Val(.TextMatrix(intRow, 0)) And _
                            str���� = .TextMatrix(intRow, mconIntCol����) And _
                            str���� = .TextMatrix(intRow, mconIntCol����) And _
                            dbl�ɱ��� = Val(.TextMatrix(intRow, mconintCol�ɱ���)) Then
                            
                            MsgBox "��" & intLop & "�е�ҩƷ(" & Trim(.TextMatrix(intLop, mconIntColҩ��)) & ")�������εĲ��أ����ţ��ɱ��ۺ͵�" & intRow & "���ظ��ˣ�" & vbCrLf & "������¼����غ�������Ϣ��", vbInformation, gstrSysName
                            
                            vsfBill.SetFocus
                            .Row = intLop
                            .TopRow = intLop
                            .Col = mconIntCol����
                            .EditCell
                            Exit Function
                        End If
                    Next
                End If
                
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function SaveCard() As Boolean
    Dim lng������id As Long
    Dim int���ϵ�� As Integer
    Dim lng������ID As Integer
    Dim lng�������ID As Integer
    
    Dim chrNo As Variant
    Dim lng��� As Long
    Dim lng�ⷿID As Long
    Dim lngҩƷID As Long
    Dim str���� As String
    Dim lng����ID As Long
    Dim str���� As String
    Dim datЧ�� As String
    Dim dbl�������� As Double
    Dim dblʵ������ As Double
    Dim dbl������ As Double
    Dim dbl�ۼ� As Double
    Dim dbl�ɱ��� As Double
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim dat�������� As String
    Dim str�̵�ʱ�� As String
    Dim dbl����� As Double
    Dim dbl����� As Double
    Dim rs������ As New Recordset
    Dim intRow As Integer
    Dim str��׼�ĺ� As String
    Dim int������ As Integer
    Dim arrSql As Variant
    Dim i As Integer
    
    Dim str���ݺ�() As String
    Dim n As Long
    
    Dim intMoneyBit As Integer
    Dim dbl����ϵ�� As Double
    Dim str�ⷿ��λ As String
    
    arrSql = Array()
    SaveCard = False
    On Error GoTo ErrHandle
    '����������������ID����Ҫ������ҩƷ��Ҫ����
    gstrSQL = "SELECT b.ϵ��,b.id AS ���id " _
            & "FROM ҩƷ�������� a, ҩƷ������ b " _
            & "Where a.���id = b.ID AND a.���� = 12 "
    Set rs������ = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption)
    If rs������.EOF Then
        MsgBox "�Բ���û������ҩƷ�̵���������������ҩƷ�������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    lng������ID = 0
    lng�������ID = 0
    
    rs������.MoveFirst
    Do While Not rs������.EOF
        If rs������!ϵ�� = 1 Then
            lng������ID = rs������!���id
        Else
            lng�������ID = rs������!���id
        End If
        rs������.MoveNext
    Loop
    rs������.Close
    
    If lng������ID = 0 Then
        MsgBox "�Բ���û������ҩƷ�̵���������������ҩƷ�������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If lng�������ID = 0 Then
        MsgBox "�Բ���û������ҩƷ�̵����ĳ����������ҩƷ�������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    With vsfBill
        chrNo = Trim(txtNo)
        lng�ⷿID = txtStock.Tag
        If chrNo = "" Then chrNo = Sys.GetNextNo(29, lng�ⷿID)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        
        strժҪ = Trim(txtժҪ.Text)
        str������ = Txt������
        dat�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str�̵�ʱ�� = txtCheckDate.Caption
        
        If mint�༭״̬ = 2 Or mbln���䶯 = True Then       '�޸�
            gstrSQL = "zl_ҩƷ�̵�_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                int������ = 0
                If Val(.TextMatrix(intRow, mconIntCol����)) = -1 Or Val(.TextMatrix(intRow, mconIntCol������)) = 1 Then
                    int������ = 1
                End If
                
                lngҩƷID = .TextMatrix(intRow, 0)
                dbl����ϵ�� = IIf(mintUnit > 0, Val(.TextMatrix(intRow, mconIntCol����ϵ��)), Val(.TextMatrix(intRow, mconIntCol����ϵ��С)))
                str���� = .TextMatrix(intRow, mconIntCol����)
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng����ID = IIf(.TextMatrix(intRow, mconIntCol����) = "", 0, .TextMatrix(intRow, mconIntCol����))
                datЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datЧ�� <> "" Then
                    '����ΪʧЧ��������
                    datЧ�� = Format(DateAdd("D", 1, datЧ��), "yyyy-mm-dd")
                End If
                
                dbl�������� = Val(.TextMatrix(intRow, mconintCol�������))
                dblʵ������ = zlStr.FormatEx(.TextMatrix(intRow, mconintColʵ������) * dbl����ϵ��, gtype_UserDrugDigits.Digit_����, , True)

                If Val(.TextMatrix(intRow, mconintCol��������)) = Val(.TextMatrix(intRow, mconintColʵ������)) Then
                    If dbl�������� <> Val(.TextMatrix(intRow, mconintCol��������)) * dbl����ϵ�� Then
                        '��ʵ������������ͽ����������������Ĳ�һ��ʱ(���ھ���ȡ�ᵼ�µģ����ܵ����̵���޷��õ�Ԥ�ڵ�ʵ������)
                        'ʹ����ʵ�����������ʵ����������������
                        dbl������ = Val(.TextMatrix(intRow, mconintColʵ������)) * dbl����ϵ�� - dbl��������
                    Else
                        dbl������ = 0
                    End If
                    dblʵ������ = Val(.TextMatrix(intRow, mconintCol�������))
                Else
                    dbl������ = zlStr.FormatEx(Abs(.TextMatrix(intRow, mconintColʵ������) * dbl����ϵ�� - Val(.TextMatrix(intRow, mconintCol�������))), gtype_UserDrugDigits.Digit_����, , True)
                End If
                
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                              
                dbl�ۼ� = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / dbl����ϵ��, gtype_UserDrugDigits.Digit_���ۼ�)
                dbl�ɱ��� = zlStr.FormatEx(.TextMatrix(intRow, mconintCol�ɱ���) / dbl����ϵ��, gtype_UserDrugDigits.Digit_�ɱ���)

                If Val(Split(.TextMatrix(intRow, mconIntcol�ӳ���), "||")(1)) = 0 Or int������ = 0 Then
                    '����ҩƷ������������ȡԭʼ�ۼ�
                    dbl�ۼ� = Get�̵�ʱ���ۼ�(Split(.TextMatrix(intRow, mconIntcol�ӳ���), "||")(1) = 1, lngҩƷID, lng�ⷿID, lng����ID, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                Else
                    '��������ʱ�۰�����۸���󱣴�
                    dbl�ۼ� = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / dbl����ϵ��, gtype_UserDrugDigits.Digit_���ۼ�)
                End If

                If int������ = 0 Then
                    '������������ȡԭʼ�ɱ���
                    dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����ID, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                Else
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(lngҩƷID) = True Then
                        dbl�ɱ��� = dbl�ۼ�
                    Else
                        '�������ΰ�����۸���󱣴�
                        dbl�ɱ��� = zlStr.FormatEx(.TextMatrix(intRow, mconintCol�ɱ���) / dbl����ϵ��, gtype_UserDrugDigits.Digit_�ɱ���)
                    End If
                End If
      
                str�ⷿ��λ = IIf(Trim(.TextMatrix(intRow, mconIntCol�ⷿ��λ)) = "", "", .TextMatrix(intRow, mconIntCol�ⷿ��λ))
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '���ҩƷ���������Ϊ0�������۲�Ϊ0��ҩƷ�޷�ͨ���̵��������¼������
                '��������µ�ͨ��ҩƷ�������۵�ʵ��λ������ϵͳ���������õĽ��λ��
                '����취�����ʵ������Ϊ0�������Ͳ�۲�С��λ�����ֺ�ҩƷ�����н��Ͳ��λ��һ��
                If int������ = 1 Then
                    intMoneyBit = mintMoneyDigit
                ElseIf dblʵ������ = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True And Val(.TextMatrix(intRow, mconIntCol�ۼ�)) = Val(.TextMatrix(intRow, mconintCol�ɱ���))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
        
                dbl���� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol����)), intMoneyBit, , True)
                dbl��۲� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol��۲�)), intMoneyBit, , True)
                dbl����� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                dbl����� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                If dbl�������� <= dblʵ������ Then
                    lng������id = lng������ID
                    int���ϵ�� = 1
                Else
                    lng������id = lng�������ID
                    int���ϵ�� = -1
                End If
                 
                lng��� = intRow
                
                gstrSQL = "zl_ҩƷ�̵�_INSERT('" & chrNo & "'," & lng��� & "," & lng�ⷿID & "," & lng����ID & "," _
                    & lng������id & "," & int���ϵ�� & "," & lngҩƷID & "," & dbl�������� & "," _
                    & dblʵ������ & "," & dbl������ & "," & dbl�ۼ� & "," & dbl���� & "," & dbl��۲� & ",'" _
                    & str������ & "',to_date('" & dat�������� & "','yyyy-mm-dd HH24:MI:SS'),'" _
                    & strժҪ & "','" & str���� & "','" & str���� & "'," & IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" _
                    & str�̵�ʱ�� & "'," & dbl����� & "," & dbl����� & "," & dbl�ɱ��� & ",'" & str��׼�ĺ� & "'," & int������ & ",'" & str�ⷿ��λ & "')"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                
            End If
            recSort.MoveNext
        Next
        
        If mint�༭״̬ = 5 Then
            If InStr(mstr�̵㵥��, ",") = 0 Then
                ReDim str���ݺ�(0)
                str���ݺ�(0) = mstr�̵㵥��
            Else
                str���ݺ� = Split(mstr�̵㵥��, ",")
            End If
            
            If mblnɾ���̵㵥 Then
                For n = 0 To UBound(str���ݺ�)
                    gstrSQL = "Zl_ҩƷ�̵��¼��_DELETE(" & str���ݺ�(n) & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                Next
            Else
                For n = 0 To UBound(str���ݺ�)
                    gstrSQL = "Zl_ҩƷ�̵��¼��_Update(" & str���ݺ�(n) & ")"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                Next
            End If
        End If
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim dbl���� As Double
    Dim dbl�̵��� As Double
    Dim intLop As Integer
    Dim dbl�ɱ���� As Double
    
    dbl���� = 0
    dbl�̵��� = 0
    dbl�ɱ���� = 0
    
    With vsfBill
        For intLop = 1 To .rows - 1
            If .TextMatrix(intLop, 0) <> "" Then
                dbl���� = dbl���� + Val(.TextMatrix(intLop, mconintCol����)) * IIf(.TextMatrix(intLop, mconintCol��־) = "��", -1, 1)
                dbl�̵��� = dbl�̵��� + Val(.TextMatrix(intLop, mconIntColʵ�ʽ��)) + Val(.TextMatrix(intLop, mconintCol����)) * IIf(.TextMatrix(intLop, mconintCol��־) = "��", -1, 1) * IIf(mint��¼״̬ = 1, 1, IIf(mint��¼״̬ Mod 3 = 0, 1, -1))
                dbl�ɱ���� = dbl�ɱ���� + Val(.TextMatrix(intLop, mconintCol�̵�ɱ����))
            End If
        Next
    End With
    
    lblPurchasePrice.Caption = "����ϼƣ�" & zlStr.FormatEx(dbl����, mintMoneyDigit, , True)
    lblPurchasePrice.Width = Pic����.TextWidth(lblPurchasePrice.Caption)
    lblCheckSum.Left = lblPurchasePrice.Left + lblPurchasePrice.Width + 200

    lblCheckSum.Caption = "�̵���ϼƣ�" & zlStr.FormatEx(dbl�̵���, mintMoneyDigit, , True)
    lblCheckSum.Width = Pic����.TextWidth(lblCheckSum.Caption)
    
    lblCostPrice.Top = lblCheckSum.Top
    lblCostPrice.Left = lblCheckSum.Left + lblCheckSum.Width + 200
    lblCostPrice.Caption = "�̵�ɱ����ϼƣ�" & zlStr.FormatEx(dbl�ɱ����, mintMoneyDigit, , True)
    lblCostPrice.Width = Pic����.TextWidth(lblCostPrice.Caption)
End Sub

Private Sub ��ʾ�����()
    Dim rsUseCount As New Recordset
    Dim dbl���װ���� As Double
    Dim dblС��װ���� As Double
    Dim dbl���װʵ������ As Double
    Dim dblС��װʵ������ As Double
    
    On Error GoTo ErrHandle
    With vsfBill
        If .TextMatrix(.Row, mconIntColҩ��) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(vsfBill.Row, 0) = "" Then Exit Sub
        
        gstrSQL = "select Nvl(��������,0) ��������,nvl(ʵ������,0) ʵ������ from ҩƷ��� " _
                & "where �ⷿid=[1] " _
                & "  and ҩƷid=[2] " _
                & "  and ����=1 " _
                & "  and nvl(����,0)=[3]"
        Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", txtStock.Tag, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
        
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol��������) = 0
        Else
            If mintUnit > 0 Then
                dbl���װ���� = rsUseCount!�������� / Val(.TextMatrix(.Row, mconIntCol����ϵ��))
                dbl���װʵ������ = rsUseCount!ʵ������ / Val(.TextMatrix(.Row, mconIntCol����ϵ��))
                
                .TextMatrix(.Row, mconIntCol��������) = dbl���װ����
            Else
                dbl���װ���� = Int(rsUseCount!�������� / Val(.TextMatrix(.Row, mconIntCol����ϵ����)))
                dbl���װʵ������ = Int(rsUseCount!ʵ������ / Val(.TextMatrix(.Row, mconIntCol����ϵ����)))
                
                dblС��װ���� = zlStr.FormatEx((Val(rsUseCount!��������) - dbl���װ���� * Val(.TextMatrix(.Row, mconIntCol����ϵ����))) / Val(.TextMatrix(.Row, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                dblС��װʵ������ = zlStr.FormatEx((Val(rsUseCount!ʵ������) - dbl���װʵ������ * Val(.TextMatrix(.Row, mconIntCol����ϵ����))) / Val(.TextMatrix(.Row, mconIntCol����ϵ��С)), mintNumberDigit0, , True)
                
               .TextMatrix(.Row, mconIntCol��������) = rsUseCount!�������� / Val(.TextMatrix(.Row, mconIntCol����ϵ��С))
            End If
        End If
        rsUseCount.Close
        
        If mintUnit > 0 Then
            staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & zlStr.FormatEx(dbl���װʵ������, mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol��λ)
        Else
            staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & zlStr.FormatEx(dbl���װʵ������, mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol����������λ��) & _
                ",[" & zlStr.FormatEx(dblС��װʵ������, mintNumberDigit0, , True) & "]" & .TextMatrix(.Row, mconIntCol����������λС)
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtժҪ_Change()
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    OS.OpenIme True
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
     OS.OpenIme
End Sub

Private Function SetPhiscRows(ByVal lngID As Long, ByVal lng���� As Long, ByVal dbl��ʼ�ɱ��� As Double, ByVal dbl����ϵ�� As Double, ByVal str��׼�ĺ� As String) As Boolean
'���ܣ�����ҩƷID���̴������ʾ�������ҩƷ�ĳ�ʼ�̴���Ϣ
'˵����
'   1.����ǷǷ�������ҩ,���Ѿ�������,����ʾ���˳���
'   2.����Ƿ�������ҩ����ֱ����ҩ��δ����ĸ����ο���С�
    Dim i As Integer, lngRow As Long
    Dim rsDetail As ADODB.Recordset
    Dim intRecordCount As Integer
    Dim intCurrentRow As Integer
    Dim intRow As Integer
    Dim bln�ⷿ As Boolean
    Dim dbl�ɱ��� As Double, dbl���ۼ� As Double, dbl�ӳ��� As Double
    Dim str���� As String
    Dim lngBatch As Long
    Dim intMoneyBit As Integer
    Dim intOld As Integer
    Dim n As Integer
    Dim rsʱ�۷��� As ADODB.Recordset
    Dim rsDingPrice As ADODB.Recordset
    Dim strҩ�� As String
    Dim bln�̵���� As Boolean
    Dim str�̵�ʱ�� As String
     
    On Error GoTo errH
    
    str�̵�ʱ�� = txtCheckDate.Caption
    
    Set rsDetail = GetPhysicDetail(txtStock.Tag, lngID)
    intRecordCount = rsDetail.RecordCount
    If intRecordCount = 0 Then Exit Function
    
    mstrMsg = ""
    
    '��������ҩƷ
    If lng���� <> -1 Then
        rsDetail.MoveFirst
        rsDetail.Find "����=" & lng����
        If rsDetail.EOF Then Exit Function
    End If
    
    bln�ⷿ = CheckPartProp(Val(txtStock.Tag))
    With vsfBill
        vsfBill.Redraw = flexRDNone
        intRow = .Row
        .TextMatrix(intRow, 0) = rsDetail!ҩƷid
        
        dbl�ɱ��� = zlStr.Nvl(rsDetail!ƽ���ɱ���, 0)
        dbl���ۼ� = IIf(IsNull(rsDetail!�ۼ�), 0, rsDetail!�ۼ�)
        '�������̵���������˵�ҩƷ
        If rsDetail!�Ƿ��� = 0 And IsNull(rsDetail!�ۼ�) Then
            gstrSQL = "select �ּ� from �շѼ�Ŀ where �շ�ϸĿid=[1] and sysdate between ִ������ and ��ֹ����" & _
                    GetPriceClassString("")
            
            Set rsDingPrice = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ۸�", rsDetail!ҩƷid)
            If rsDingPrice.EOF = False Then
                dbl���ۼ� = rsDingPrice!�ּ�
            End If
        End If
        
        If rsDetail!�Ƿ��� = 1 Then
            dbl���ۼ� = Get�̵�ʱ�����ۼ�(Val(.TextMatrix(intRow, 0)), Val(txtStock.Tag), lng����, 1, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
        End If
        
        '�ж����޿�棬����޿����Ϊ����ҩƷ
        If lng���� = 0 Then
            If CheckNoStock(Val(txtStock.Tag), Val(.TextMatrix(intRow, 0))) = True Then
                '�޿��ʱΪ�̵����
                bln�̵���� = True
                If IsPriceAdjustMod(rsDetail!ҩƷid) = True Then
                    If rsDetail!�Ƿ��� = 1 Then
                        '���۹���ʱ��ҩƷ�ۼ�Ҫ���ڳɱ���
                        dbl���ۼ� = dbl�ɱ���
                    Else
                        '���۹�������ҩƷ�ɱ���Ҫ�����ۼ�
                        dbl�ɱ��� = dbl���ۼ�
                    End If
                End If
            End If
        End If
        
        '�������������ʱ
        If lng���� = -1 Then
            If rsDetail!�Ƿ��� = 0 Then
                '����
                If IsPriceAdjustMod(rsDetail!ҩƷid) = True Then
                    '���۹����ɱ���Ҫ�����ۼ�
                    dbl�ɱ��� = dbl���ۼ�
                End If
            Else
                'ʱ��
                If IsPriceAdjustMod(rsDetail!ҩƷid) = True Then
                    '���۹����ۼ�Ҫ���ڳɱ���
                    dbl���ۼ� = dbl�ɱ���
                Else
                    dbl���ۼ� = Get�̵�ʱ�����ۼ�(Val(.TextMatrix(intRow, 0)), Val(txtStock.Tag), lng����, 1, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")))
                End If
            End If
        End If
        
        str���� = zlStr.Nvl(rsDetail!ȱʡ����, "")
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = rsDetail!ͨ����
        Else
            strҩ�� = IIf(IsNull(rsDetail!��Ʒ��), rsDetail!ͨ����, rsDetail!��Ʒ��)
        End If
        
        .TextMatrix(intRow, mconIntColҩƷ���������) = rsDetail!ҩƷ���� & strҩ��
        .TextMatrix(intRow, mconIntColҩƷ����) = rsDetail!ҩƷ����
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        Else
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(rsDetail!��Ʒ��), "", rsDetail!��Ʒ��)
        
        If .Col = mconIntColҩ�� Then
            .EditText = .TextMatrix(intRow, mconIntColҩ��)
        End If
        
        .TextMatrix(intRow, mconIntCol��Դ) = zlStr.Nvl(rsDetail!ҩƷ��Դ)
        .TextMatrix(intRow, mconIntCol����ҩ��) = zlStr.Nvl(rsDetail!����ҩ��)
        .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsDetail!���), "", rsDetail!���)
        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsDetail!����), "", rsDetail!����)
        If .TextMatrix(intRow, mconIntCol����) = "" Then .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol�ⷿ��λ) = IIf(IsNull(rsDetail!�ⷿ��λ), "", rsDetail!�ⷿ��λ)
        
        If mintUnit > 0 Then
            '������������и�ʽ��
            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��, mintPriceDigit, , True)
            
            .TextMatrix(intRow, mconIntCol��λ) = IIf(IsNull(rsDetail!��λ), "", rsDetail!��λ)
            .TextMatrix(intRow, mconIntCol����ϵ��) = rsDetail!����ϵ��
            
            If rsDetail!�Ƿ��� = 1 Then
                .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(Get�̵�ʱ�̳ɱ���(rsDetail!ҩƷid, Val(txtStock.Tag), CLng(rsDetail!����), CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss"))) * dbl����ϵ��, mintCostDigit, , True)
                If IsPriceAdjustMod(rsDetail!ҩƷid) = True Then
                    '���۹����ۼ�Ҫ���ڳɱ���
                    .TextMatrix(intRow, mconIntCol�ۼ�) = .TextMatrix(intRow, mconintCol�ɱ���)
                End If
            Else
                If IsPriceAdjustMod(rsDetail!ҩƷid) = True Then
                    .TextMatrix(intRow, mconintCol�ɱ���) = .TextMatrix(intRow, mconIntCol�ۼ�)
                Else
                    .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(dbl��ʼ�ɱ��� * dbl����ϵ��, mintCostDigit, , True)
                End If
            End If
        Else
            '������������и�ʽ��
            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ� * rsDetail!����ϵ��С, mintPriceDigit0, , True)
            
            .TextMatrix(intRow, mconIntCol����������λ��) = rsDetail!���װ��λ
            .TextMatrix(intRow, mconIntCol����������λС) = rsDetail!С��װ��λ
            .TextMatrix(intRow, mconIntColʵ��������λ��) = rsDetail!���װ��λ
            .TextMatrix(intRow, mconIntColʵ��������λС) = rsDetail!С��װ��λ
            
            .TextMatrix(intRow, mconIntCol����ϵ����) = zlStr.Nvl(rsDetail!����ϵ����, 0)
            .TextMatrix(intRow, mconIntCol����ϵ��С) = zlStr.Nvl(rsDetail!����ϵ��С, 0)
            
            If rsDetail!�Ƿ��� = 1 Then
                .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(Get�̵�ʱ�̳ɱ���(rsDetail!ҩƷid, Val(txtStock.Tag), CLng(rsDetail!����), CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss"))) * rsDetail!����ϵ��С, mintCostDigit0, , True)
                If IsPriceAdjustMod(rsDetail!ҩƷid) = True Then
                    '���۹����ۼ�Ҫ���ڳɱ���
                    .TextMatrix(intRow, mconIntCol�ۼ�) = .TextMatrix(intRow, mconintCol�ɱ���)
                End If
            Else
                If IsPriceAdjustMod(rsDetail!ҩƷid) = True Then
                    .TextMatrix(intRow, mconintCol�ɱ���) = .TextMatrix(intRow, mconIntCol�ۼ�)
                Else
                    .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(dbl��ʼ�ɱ��� * rsDetail!����ϵ��С, mintCostDigit0, , True)
                End If
            End If
        End If
            
        .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsDetail!����), "0", rsDetail!����)
        If CheckPhysicBatch(bln�ⷿ, rsDetail!��������, rsDetail!ҩ����������) And Val(.TextMatrix(intRow, mconIntCol����)) = 0 Then
            lng���� = -1
        End If
        
        If lng���� = -1 Or bln�̵���� = True Then
            .TextMatrix(intRow, mconIntCol������) = 1
            .TextMatrix(intRow, mconIntCol����) = lng����
            .TextMatrix(intRow, mconIntCol����) = ""
            .TextMatrix(intRow, mconIntColЧ��) = ""
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
            
            .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconintColʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol��������), mintNumberDigit, , True)
            
            If mintUnit = 0 Then
                .TextMatrix(intRow, mconintCol���װ��������) = zlStr.FormatEx(0, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintColС��װ��������) = zlStr.FormatEx(0, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol���װʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol���װ��������), mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintColС��װʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconintColС��װ��������), mintNumberDigit0, , True)
            End If
            
            .TextMatrix(intRow, mconintCol�̵���) = zlStr.FormatEx(0, mintMoneyDigit, , True)
            .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconIntColʵ�ʽ��) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconintCol�������) = zlStr.FormatEx(0, mintNumberDigit, , True)
            .TextMatrix(intRow, mconIntColʵ�ʲ��) = zlStr.FormatEx(0, mintMoneyDigit, , True)
            If mintUnit <= 0 Then
                .TextMatrix(intRow, mconintCol�ϼ�) = .TextMatrix(intRow, mconintColʵ������) & rsDetail!С��װ��λ
            End If
        Else
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsDetail!����), "0", rsDetail!����)
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsDetail!����), "", rsDetail!����)
            .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsDetail!Ч��), "", Format(rsDetail!Ч��, "yyyy-MM-dd"))
            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                '����Ϊ��Ч��
                .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
            End If
            
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsDetail!��׼�ĺ�), "", rsDetail!��׼�ĺ�)
            
            If mintUnit > 0 Then
                .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��, mintNumberDigit, , True)
                .TextMatrix(intRow, mconintColʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol��������), mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0), mintNumberDigit, , True)
                
                .TextMatrix(intRow, mconintCol�ɱ���) = zlStr.FormatEx(zlStr.Nvl(rsDetail!�ɱ���, 0) * rsDetail!����ϵ��, mintCostDigit, , True)
            Else
                .TextMatrix(intRow, mconintCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintColʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol��������), mintNumberDigit0, , True)
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(zlStr.Nvl(rsDetail!��������, 0), mintNumberDigit0, , True)
                
                .TextMatrix(intRow, mconintCol���װ��������) = zlStr.FormatEx(Int(rsDetail!�������� / rsDetail!����ϵ����), mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintCol���װʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconintCol���װ��������), mintNumberDigit0, , True)

                .TextMatrix(intRow, mconintColС��װ��������) = zlStr.FormatEx((Val(rsDetail!��������) - Val(.TextMatrix(intRow, mconintCol���װ��������)) * Val(rsDetail!����ϵ����)) / rsDetail!����ϵ��С, mintNumberDigit0, , True)
                .TextMatrix(intRow, mconintColС��װʵ������) = zlStr.FormatEx(.TextMatrix(intRow, mconintColС��װ��������), mintNumberDigit0, , True)

                If mintUnit <= 0 Then
                    .TextMatrix(intRow, mconintCol�ϼ�) = .TextMatrix(intRow, mconintColʵ������) & rsDetail!С��װ��λ
                End If
            End If
            .TextMatrix(intRow, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintColʵ������)) * Val(.TextMatrix(intRow, mconIntCol�ۼ�)), mintMoneyDigit, , True)
            .TextMatrix(intRow, mconIntColʵ�ʽ��) = zlStr.Nvl(rsDetail!ʵ�ʽ��, 0)
            .TextMatrix(intRow, mconintCol�������) = zlStr.Nvl(rsDetail!��������, 0)
            .TextMatrix(intRow, mconIntColʵ�ʲ��) = zlStr.Nvl(rsDetail!ʵ�ʲ��, 0)
        End If
        
        .TextMatrix(intRow, mconIntcol�ӳ���) = rsDetail!�ӳ��� / 100 & "||" & rsDetail!�Ƿ��� & "||" & rsDetail!ҩ����������
        .TextMatrix(intRow, mconintCol��־) = "ƽ"
        .TextMatrix(intRow, mconintCol������) = zlStr.FormatEx("0", mintNumberDigit, , True)
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '���ҩƷ���������Ϊ0�������۲�Ϊ0��ҩƷ�޷�ͨ���̵��������¼������
        '��������µ�ͨ��ҩƷ�������۵�ʵ��λ������ϵͳ���������õĽ��λ��
        '����취�����ʵ������Ϊ0�������Ͳ�۲�С��λ�����ֺ�ҩƷ�����н��Ͳ��λ��һ��
        If Val(.TextMatrix(intRow, mconintColʵ������)) = 0 And lng���� <> -1 And bln�̵���� = False Then
            intMoneyBit = mintMaxMoneyBit
        Else
            intMoneyBit = mintMoneyDigit
        End If
        
        '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
        '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
        .TextMatrix(intRow, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) * Val(.TextMatrix(intRow, mconintColʵ������)) - Val(.TextMatrix(intRow, mconIntColʵ�ʽ��)), intMoneyBit, , True)
        .TextMatrix(intRow, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol�ۼ�)) - Val(.TextMatrix(intRow, mconintCol�ɱ���))) * Val(.TextMatrix(intRow, mconintColʵ������)) - Val(.TextMatrix(intRow, mconIntColʵ�ʲ��)), intMoneyBit, , True)
        
        '.TextMatrix(intRow, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol�ɱ���)) * Val(.TextMatrix(intRow, mconintColʵ������)), mintMoneyDigit)
        '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
        .TextMatrix(intRow, mconintCol�̵�ɱ����) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntColʵ�ʽ��)) + Val(.TextMatrix(intRow, mconintCol����))) - (Val(.TextMatrix(intRow, mconIntColʵ�ʲ��)) + Val(.TextMatrix(intRow, mconintCol��۲�))), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol����)) - Val(.TextMatrix(intRow, mconintCol��۲�)), mintMoneyDigit, , True)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If mbln��ͣ��ҩƷ = True Then
            '�����ͣ��ҩƷ�����д�����ʾ
            If Format(rsDetail!����ʱ��, "YYYY-MM-DD") <> "3000-01-01" Then
                .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
            End If
        End If
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, intRow, mconintColʵ������, intRow, mconintColʵ������) = True
        Else
            .Cell(flexcpFontBold, intRow, mconintCol���װʵ������, intRow, mconintCol���װʵ������) = True
            .Cell(flexcpFontBold, intRow, mconintColС��װʵ������, intRow, mconintColС��װʵ������) = True
        End If
        
        '���÷�������
        Call GetҩƷ��������(intRow)
        
        Call RefreshRowNO(vsfBill, mconIntCol�к�, 1)
        vsfBill.Redraw = flexRDDirect
    End With
    rsDetail.Close
    SetPhiscRows = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'��һ���в���
Private Sub InsertRow(ByVal intRow As Integer, ByVal intRecordCount As Integer)
    Dim blnHaveData As Boolean
    Dim intOldRows As Integer
    Dim intLop As Integer
    Dim intExchange As Integer
    Dim intCol As Integer
    
    With vsfBill
        blnHaveData = False
        intOldRows = .rows - 1
        .rows = .rows + intRecordCount
        For intLop = intRow + 1 To intRecordCount
            If .TextMatrix(intLop, 0) <> "" Then
                blnHaveData = True
                Exit For
            End If
        Next
        If blnHaveData = True Then
            For intExchange = .rows - 1 To intOldRows Step -1
                For intCol = 0 To .Cols - 1
                    .TextMatrix(intExchange, intCol) = .TextMatrix(intExchange - intRecordCount, intCol)
                    .TextMatrix(intExchange - intRecordCount, intCol) = ""
                Next
            Next
        End If
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'��ӡ����
Private Sub printbill()
    Dim int��λϵ�� As Integer
    Dim strNo As String
    
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
    
    strNo = txtNo.Tag
    Call FrmBillPrint.ShowME(Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1307", "zl8_bill_1307"), mint��¼״̬, int��λϵ��, 1307, "ҩƷ�̵��", strNo)
End Sub

Private Function CheckPartProp(ByVal lng�ⷿID As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    '���ⷿ���ԣ������ҩ�⣬������
    On Error GoTo ErrHandle
    gstrSQL = "SELECT count(*) " _
            & "From ��������˵�� " _
            & "WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')) " _
            & "  AND ����id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж���ҩ��/ҩ��]", lng�ⷿID)
    
    If rsTemp.Fields(0) > 0 Then
        CheckPartProp = False
    Else
        CheckPartProp = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckPhysicBatch(ByVal bln�ⷿ As Boolean, ByVal intҩ����� As Integer, ByVal intҩ������ As Integer) As Boolean
    '���ظ�ҩƷ�Ƿ�����ı�ʶ
    CheckPhysicBatch = (bln�ⷿ And (intҩ����� = 1)) Or (Not bln�ⷿ And (intҩ������ = 1))
End Function

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Set rsBatchNolen = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "-ȡ���ų���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPhysic(ByVal lng�ⷿID As Long, ByVal str�̵����� As String, _
        ByVal str���� As String, Optional ByVal str�ⷿ��λ As String = "����", _
        Optional ByVal bln���޿��ҩƷ As Boolean = True, _
        Optional ByVal bln�����̵㵥 As Boolean = False, _
        Optional ByVal bln�̵㵥 As Boolean = False, _
        Optional ByVal bln���޿���н��ҩƷ As Boolean = False) As ADODB.Recordset
    '��ȡ������������ҩƷ��ͬʱ�����λ���װϵ����
    'bln���޿��ҩƷ=�Ƿ��޿��ҩƷҲ��ȡ����
    'bln�����̵㵥=�Ƿ���Ҫ����ָ���̵�ʱ����̵㵥�γ��̵��
    'bln�̵㵥=�Ƿ������̵㵥�����̵�����Ϊ�٣�˵��Ҫ�����п��һ����������ܣ������̵㵥�е�ҩƷ��ʵ������������ʾ
    Dim str��λ As String, str�̵�ʱ�� As String, str�����̵㵥 As String
    Dim strOrder As String, strCompare As String
    Dim rsTemp As New ADODB.Recordset
    Dim strNo�� As String
    Dim str�̵㵥NO As String
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    If str�ⷿ��λ = "" Then
        str�ⷿ��λ = "����"
    ElseIf str�ⷿ��λ <> "����" Then
        str�ⷿ��λ = Replace(str�ⷿ��λ, "'", "")
    End If
    
    If str���� = "" Then str���� = "'zyb'"          '��֤����ļ���Ϊ��ʱ��������κ�ҩƷ
    
    str�̵�ʱ�� = txtCheckDate.Caption
    strOrder = zlDatabase.GetPara("����", glngSys, ģ���.ҩƷ�̵�)
    strCompare = Mid(strOrder, 1, 1)

    '����ָ���̵�ʱ�̵��̵㵥
    str�����̵㵥 = " Union " & _
             " Select A.ҩƷID,B.����,B.����,E.�ⷿ��λ" & _
             " From (select DISTINCT a.ҩƷID,a.�ⷿID FROM ҩƷ�շ���¼ a " & _
             " Where a.����=14 And a.�ⷿID+0=[1] And a.No in (select * from Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)))) A, " & _
             " �շ���ĿĿ¼ B,ҩƷ�����޶� E " & _
             " Where A.ҩƷID+0=B.ID and A.�ⷿid=E.�ⷿid(+) and A.ҩƷid+0=E.ҩƷid(+) "
    If mbln���Է������ = False Then
         str�����̵㵥 = str�����̵㵥 & " And(Decode(B.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(1,3))" & _
                " or Decode(B.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(2,3)) " & _
                " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[1]) )"
    End If
    
    '��ȡ�����̵�����������ҩƷ�嵥
    gstrSQL = "SELECT " & IIf(str�ⷿ��λ <> "����", " /*+rule*/ ", "") & " Distinct A.ҩƷID,B.����,B.����,E.�ⷿ��λ" & _
             " FROM ҩƷ��� A,�շ���ĿĿ¼ B,ҩƷ���� C,������ĿĿ¼ K,���Ʒ���Ŀ¼ L," & _
             "     (SELECT ҩƷID,Nvl(ʵ������,0) ʵ������,Nvl(ʵ�ʽ��,0) ʵ�ʽ��,Nvl(ʵ�ʲ��,0) ʵ�ʲ�� " & _
             "      FROM ҩƷ��� WHERE �ⷿID=[1] AND ����=1 " & IIf(bln���޿���н��ҩƷ = True, " And ʵ������=0 And (ʵ�ʽ��<>0 Or ʵ�ʲ��<>0)", " And (Nvl(ʵ������,0)<>0 Or Nvl(ʵ�ʽ��,0)<>0 Or Nvl(ʵ�ʲ��,0)<>0 )") & ") D, "
    If bln�����̵㵥 Then
        gstrSQL = gstrSQL & "(SELECT �ⷿid, ҩƷid, ����, ����, �̵�����, �ⷿ��λ FROM ҩƷ�����޶� WHERE �ⷿID=[1]) E, " & _
             "     (SELECT �շ�ϸĿid, ������Դ, ��������id, ִ�п���id FROM �շ�ִ�п��� WHERE ִ�п���ID=[1]) F " & _
             " WHERE A.ҩƷID=B.ID And A.ҩ��ID=K.ID And K.����ID=L.ID and L.���� in (1,2,3) And A.ҩ��ID=C.ҩ��ID AND A.ҩƷID=F.�շ�ϸĿID" & IIf(mblnNoStock, "(+)", "") & _
             "  AND (B.����ʱ��=TO_DATE('3000-01-01','yyyy-MM-dd') OR B.����ʱ�� IS NULL Or B.����ʱ�� BETWEEN To_Date('" & str�̵�ʱ�� & "', 'yyyy-mm-dd hh24:mi:ss') AND SYSDATE) " & _
             IIf(mstr����ID = "", "", " AND L.ID in (select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) ") & _
             IIf(str���� = "����", "", " AND C.ҩƷ���� in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist))) ") & _
             "  AND A.ҩƷID=D.ҩƷID" & IIf(bln���޿��ҩƷ, "(+)", "") & " AND A.ҩƷID=E.ҩƷID(+)"
        If mbln���Է������ = False Then
            gstrSQL = gstrSQL & " And(Decode(B.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(1,3))" & _
                " or Decode(B.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(2,3)) " & _
                " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[1]) )"
        End If
    Else
        If str�ⷿ��λ <> "����" Then
'            gstrSQL = gstrSQL & "(SELECT A.ҩƷid, A.�ⷿ��λ FROM ҩƷ�����޶� A WHERE A.�ⷿID=[1] " & IIf(str�̵����� = "����", "", str�̵�����) & " And A.�ⷿ��λ in (select * from Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)))) E, "
            gstrSQL = gstrSQL & "(Select a.ҩƷid, a.�ⷿ��λ" & vbNewLine & _
                            "From ҩƷ�����޶� A, (Select * From Table(Cast(f_Str2list([2]) As Zltools.t_Strlist))) B" & vbNewLine & _
                            "Where a.�ⷿid = [1] " & IIf(str�̵����� = "����", "", str�̵�����) & " And (Instr(',' || a.�ⷿ��λ || ',', ',' || b.Column_Value || ',') > 0)) E, "
        Else
            gstrSQL = gstrSQL & "(SELECT A.ҩƷid, A.�ⷿ��λ FROM ҩƷ�����޶� A WHERE A.�ⷿID=[1] " & IIf(str�̵����� = "����", "", str�̵�����) & " ) E, "
        End If
        
        gstrSQL = gstrSQL & " (SELECT �շ�ϸĿid, ������Դ, ��������id, ִ�п���id FROM �շ�ִ�п��� WHERE ִ�п���ID=[1]) F " & _
             " WHERE A.ҩƷID=B.ID And A.ҩ��ID=K.ID And K.����ID=L.ID and L.���� in (1,2,3) And A.ҩ��ID=C.ҩ��ID AND A.ҩƷID=F.�շ�ϸĿID" & IIf(mblnNoStock, "(+)", "") & " " & _
             IIf(mbln��ͣ��ҩƷ = True, "", " AND (B.����ʱ��=TO_DATE('3000-01-01','yyyy-MM-dd') OR B.����ʱ�� IS NULL Or B.����ʱ�� BETWEEN To_Date('" & str�̵�ʱ�� & "', 'yyyy-mm-dd hh24:mi:ss') AND SYSDATE) ") & _
             IIf(mstr����ID = "", "", " AND L.ID in (select * from Table(Cast(f_Num2List([3]) As zlTools.t_NumList))) ") & _
             IIf(str���� = "����", "", " AND C.ҩƷ���� in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist))) ") & _
             "  AND A.ҩƷID=D.ҩƷID" & IIf(bln���޿��ҩƷ, "(+)", "") & " AND" & IIf(str�̵����� = "����", " A.ҩƷID=E.ҩƷID(+)", " A.ҩƷID=E.ҩƷID")
        If mbln���Է������ = False Then
            gstrSQL = gstrSQL & " And(Decode(B.�������,1,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(1,3))" & _
                " or Decode(B.�������,2,1,3,1,0)=(select distinct '1' from ��������˵�� where �������� like '%ҩ��' and ����id=[1] and ������� in(2,3)) " & _
                " or exists(select 1 from ��������˵�� where �������� like '%ҩ��' and ����id=[1]) )"
        End If
    End If
    If bln�����̵㵥 Then
        str�̵㵥NO = mstr�̵㵥�� & ","
        For i = 0 To UBound(Split(str�̵㵥NO, ","))
            If Split(str�̵㵥NO, ",")(i) <> "" Then
                strNo�� = IIf(strNo�� = "", "", strNo�� & ",") & Replace(Split(str�̵㵥NO, ",")(i), "'", "")
            End If
        Next
        
        If bln�̵㵥 = False Then
            gstrSQL = gstrSQL & str�����̵㵥
        Else
            gstrSQL = Replace(str�����̵㵥, " Union", "")
        End If
    End If
    
    gstrSQL = gstrSQL & " and b.����ʱ�� <=To_Date('" & str�̵�ʱ�� & "', 'yyyy-mm-dd hh24:mi:ss') "

    gstrSQL = gstrSQL & " Order by " & _
              IIf(strCompare = "0", "����", IIf(strCompare = "1", "����", IIf(strCompare = "2", "����", "�ⷿ��λ"))) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc") & ",����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ������������ҩƷ]", lng�ⷿID, str�ⷿ��λ, mstr����ID, str����, strNo��)
    
    Set GetPhysic = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPhysicDetail(ByVal lng�ⷿID As Long, ByVal lngҩƷID As Long, _
    Optional ByVal bln���޿��ҩƷ As Boolean = True, Optional ByVal bln�����̵㵥 As Boolean = False, Optional ByVal bln���޿���н��ҩƷ As Boolean = False) As ADODB.Recordset
    'bln���޿��ҩƷ=�Ƿ��޿��ҩƷҲ��ȡ����
    'bln�����̵㵥=�Ƿ���Ҫ����ָ���̵�ʱ����̵㵥�γ��̵��
    '��ȡ��ҩƷ��ǰ�ⷿ����������ϸ��¼
    Dim str��λ As String, str�̵�ʱ�� As String, str�����̵㵥 As String, str�����̵㵥�������� As String
    Dim rsTemp As New ADODB.Recordset
    Dim strSql���װ As String
    Dim strSqlС��װ As String
    Dim strSql�̵�ʱ��֮���� As String
    Dim str�̵㵥NO As String
    Dim strNo�� As String
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    str�̵�ʱ�� = txtCheckDate.Caption
    
    If mintUnit > 0 Then
        Select Case mintUnit
            Case mconint�ۼ۵�λ
                str��λ = ",E.���㵥λ As ��λ,1 As ����ϵ��"
            Case mconint���ﵥλ
                str��λ = ",A.���ﵥλ As ��λ,A.�����װ As ����ϵ��"
            Case mconintסԺ��λ
                str��λ = ",A.סԺ��λ As ��λ,A.סԺ��װ As ����ϵ��"
            Case mconintҩ�ⵥλ
                str��λ = ",A.ҩ�ⵥλ As ��λ,A.ҩ���װ As ����ϵ��"
        End Select
    Else
        Select Case mint��λ
            Case mconint�ۼ۵�λ
                strSql���װ = ",E.���㵥λ As ���װ��λ,1 As ����ϵ����"
            Case mconint���ﵥλ
                strSql���װ = ",A.���ﵥλ As ���װ��λ,A.�����װ As ����ϵ����"
            Case mconintסԺ��λ
                strSql���װ = ",A.סԺ��λ As ���װ��λ,A.סԺ��װ As ����ϵ����"
            Case mconintҩ�ⵥλ
                strSql���װ = ",A.ҩ�ⵥλ As ���װ��λ,A.ҩ���װ As ����ϵ����"
        End Select
        
        Select Case mintС��λ
            Case mconint�ۼ۵�λ
                strSqlС��װ = ",E.���㵥λ As С��װ��λ,1 As ����ϵ��С"
            Case mconint���ﵥλ
                strSqlС��װ = ",A.���ﵥλ As С��װ��λ,A.�����װ As ����ϵ��С"
            Case mconintסԺ��λ
                strSqlС��װ = ",A.סԺ��λ As С��װ��λ,A.סԺ��װ As ����ϵ��С"
            Case mconintҩ�ⵥλ
                strSqlС��װ = ",A.ҩ�ⵥλ As С��װ��λ,A.ҩ���װ As ����ϵ��С"
        End Select
        
        str��λ = strSql���װ & strSqlС��װ
    End If
    
    '�����̵㵥��SQL
    If bln�����̵㵥 Then
        str�̵㵥NO = mstr�̵㵥�� & ","
        For i = 0 To UBound(Split(str�̵㵥NO, ","))
            If Split(str�̵㵥NO, ",")(i) <> "" Then
                strNo�� = IIf(strNo�� = "", "", strNo�� & ",") & Replace(Split(str�̵㵥NO, ",")(i), "'", "")
            End If
        Next
        
        '35.60֧���̵㵥¼������������
        str�����̵㵥 = "" & _
            " UNION ALL" & _
            " SELECT A.�ⷿID,A.ҩƷID,NVL(A.����, 0) AS ����,0 AS ʵ������,A.���� As �̵�����," & _
                    " 0 AS ʵ�ʽ��,0 AS ʵ�ʲ��,0 AS ��������,A.����,A.����,A.Ч��,A.��׼�ĺ� " & _
            " FROM ҩƷ�շ���¼ A " & _
            " Where A.����=14 AND A.�ⷿID+0=[1] And Nvl(a.����, 0) <> -1 AND a.No in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist))) "
            
        
        str�����̵㵥�������� = "" & _
            " UNION ALL" & _
            " Select �ⷿid, ҩƷid, ����, Sum(ʵ������) As ��������, Sum(�̵�����) As �̵�����, Sum(ʵ�ʽ��) As ʵ�ʽ��, Sum(ʵ�ʲ��) As ʵ�ʲ��," & _
            " Sum(��������) As ��������, Max(����) As ����, Max(����) As ����, Max(Ч��) As Ч��, Max(��׼�ĺ�) As ��׼�ĺ�, �ɱ��� " & _
            " from (SELECT A.�ⷿID,A.ҩƷID,NVL(A.����, 0) AS ����,0 AS ʵ������,A.���� As �̵�����," & _
                    " 0 AS ʵ�ʽ��,0 AS ʵ�ʲ��,0 AS ��������,A.����,A.����,A.Ч��,A.��׼�ĺ�, a.���� As �ɱ��� " & _
            " FROM ҩƷ�շ���¼ A " & _
            " Where A.����=14 AND A.�ⷿID+0=[1] And Nvl(a.����, 0) = -1 AND a.No in (select * from Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)))) " & _
            " GROUP BY �ⷿID, ҩƷID, ����, ����, ����, �ɱ��� "
    End If
    
    If mbln�����̵�ʱ�� = False Then
        strSql�̵�ʱ��֮���� = "" & _
            " Union All" & _
            " SELECT A.�ⷿID,A.ҩƷID,NVL(A.����,0) AS ����,-1*A.���ϵ��*A.ʵ������*A.���� AS ʵ������,0 �̵�����," & _
            " -1*A.���ϵ��*A.���۽�� AS ʵ�ʽ��, -1*A.���ϵ��*A.��� AS ʵ�ʲ��,0 AS ��������,A.����,A.����,A.Ч��,A.��׼�ĺ� " & _
            " FROM ҩƷ�շ���¼ A" & _
            " Where A.�ⷿID+0=[1] And A.ҩƷID+0=[2] " & _
            " AND A.������� > [3] "
    End If
    
    'ȡҩƷ��ǰ��漰�̵�ʱ���Ժ�ľ�������
    gstrSQL = "" & _
        " SELECT DISTINCT A.ҩƷID,A.�ɱ��� As ƽ���ɱ���,E.���� ȱʡ����,'[' || E.���� || ']' As ҩƷ����, E.���� As ͨ����, C.���� As ��Ʒ��,A.ҩ����� AS ��������,A.ҩ������ AS ҩ����������,E.�Ƿ���,A.�ӳ���," & _
        "        NVL(B.ʵ�ʽ��,0) ʵ�ʽ��,NVL(B.ʵ�ʲ��,0) ʵ�ʲ��,D.�ּ� �ۼ�,NVL(B.����,0) ����,A.ҩƷ��Դ,A.����ҩ��,Decode(b.����, Null, a.�ϴ�����, b.����) As ����,B.Ч��,F.�ⷿ��λ,E.���,decode(b.����,null,decode(a.�ϴβ���,null,e.����,a.�ϴβ���),b.����) as ����," & _
        "        B.��׼�ĺ�,Nvl(B.��������,0) ��������,B.�̵�����,B.��������" & str��λ & ",Decode(b.����, -1, b.�ɱ���, Decode(x.�ּ�, Null, Decode(k.�ɱ���, Null, a.�ɱ���, k.�ɱ���), x.�ּ�)) As �ɱ���, " & _
        "        Nvl(E.����ʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) As ����ʱ�� " & _
        " FROM (SELECT �ⷿID, ҩƷID, ����, SUM (ʵ������) AS ��������,SUM (�̵�����) AS �̵�����,SUM (ʵ�ʽ��) AS ʵ�ʽ��," & _
        "         SUM (ʵ�ʲ��) AS ʵ�ʲ��, SUM(��������) AS ��������,MAX(����) As ����, MAX(����) AS ���� ,MAX(Ч��) AS Ч��, MAX(��׼�ĺ�) As ��׼�ĺ�, 0 As �ɱ��� " & _
        "         From" & _
        "             ( SELECT A.�ⷿID,A.ҩƷID,NVL(����,0) AS ����,Nvl(A.ʵ������,0) ʵ������,0 �̵�����,Nvl(A.ʵ�ʽ��,0) ʵ�ʽ��,Nvl(A.ʵ�ʲ��,0) ʵ�ʲ��,Nvl(A.��������,0) ��������,A.�ϴ����� AS ����,A.�ϴβ��� AS ����,A.Ч��,A.��׼�ĺ� " & _
        "             FROM ҩƷ��� A" & _
        "             Where A.���� = 1 And A.�ⷿID=[1] And A.ҩƷID=[2] " & IIf(bln���޿���н��ҩƷ = True, " And A.ʵ������=0 And (A.ʵ�ʽ��<>0 Or A.ʵ�ʲ��<>0)", " And (Nvl(A.ʵ������,0)<>0 Or Nvl(A.ʵ�ʽ��,0)<>0 Or Nvl(A.ʵ�ʲ��,0)<>0 )") & _
        IIf(mbln�����̵�ʱ�� = True, "", strSql�̵�ʱ��֮����) & _
        IIf(Not bln�����̵㵥, "", str�����̵㵥) & _
        "     ) GROUP BY �ⷿID, ҩƷID, ���� " & IIf(Not bln�����̵㵥, "", str�����̵㵥��������) & _
        ") B, �շѼ�Ŀ D, �շ���Ŀ���� C, �շ���ĿĿ¼ E, ҩƷ��� A," & _
        "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 2 and [3] between x.ִ������ and x.��ֹ����) X," & _
        "      (Select ҩƷid,����,ƽ���ɱ��� �ɱ��� From ҩƷ��� Where ���� = 1 And �ⷿid =[1] " & IIf(bln���޿���н��ҩƷ = True, " And ʵ������=0 And (ʵ�ʽ��<>0 Or ʵ�ʲ��<>0)", "") & ") K,ҩƷ�����޶� F " & _
        " Where A.ҩƷID=E.ID And A.ҩƷID=B.ҩƷID" & IIf(bln���޿��ҩƷ, "(+)", "") & _
        " AND A.ҩƷID=F.ҩƷID(+) And B.ҩƷid=K.ҩƷid(+) And Nvl(B.����, 0)=nvl(K.����(+),0) " & _
        " AND b.ҩƷid = x.ҩƷid(+) And b.�ⷿid = x.�ⷿid(+) And Nvl(b.����, 0) = Nvl(x.����(+), 0) " & _
        " AND A.ҩƷID=C.�շ�ϸĿID(+) AND C.����(+)=3 AND A.ҩƷID=D.�շ�ϸĿID(+)  " & _
        " AND F.�ⷿID(+)=[1] And A.ҩƷID+0=[2] AND D.ִ������(+)<=[3] AND NVL(D.��ֹ����(+),SYSDATE)>=[3] " & _
        GetPriceClassString("D") & _
        " and e.����ʱ��<=[3]  Order by ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ��ҩƷ��ǰ�ⷿ����������ϸ��¼]", lng�ⷿID, lngҩƷID, CDate(Format(str�̵�ʱ��, "yyyy-mm-dd hh:mm:ss")), strNo��)
    
    Set GetPhysicDetail = rsTemp
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function ʱ��ҩƷ���ۼ�(ByVal lngҩƷID As Long, ByVal sin�ӳ��� As Double, ByVal sin�ɹ��� As Single) As Double
    Dim sin���ۼ� As Single, sinָ�����ۼ� As Single, sin��������� As Single
    Dim rsTemp As New ADODB.Recordset
    'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
    '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
    On Error GoTo ErrHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������
    
    ʱ��ҩƷ���ۼ� = 0
    
    sin���ۼ� = sin�ɹ��� * (1 + sin�ӳ���)
    sin���ۼ� = sin���ۼ� + (sinָ�����ۼ� - sin���ۼ�) * (1 - sin��������� / 100)
    ʱ��ҩƷ���ۼ� = IIf(sin���ۼ� > sinָ�����ۼ�, sinָ�����ۼ�, sin���ۼ�)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub vsfBill_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsfBill.EditSelStart = 0
    vsfBill.EditSelLength = zlStr.ActualLen(vsfBill.EditText)
End Sub

Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intMoneyBit As Integer
    Dim intNumber As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim dbl�ɱ��� As Double
    Dim dblSumNum As Double
    Dim dbl���� As Double
    Dim dbl��۲� As Double
    
    On Error GoTo ErrHandle
    With vsfBill
        .Redraw = flexRDNone
        
        .EditText = Trim(.EditText)
        strKey = Trim(.EditText)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        
        Select Case Col
            Case mconIntCol����
                .TextMatrix(Row, Col) = strKey
            Case mconIntColЧ��
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            .EditText = ""
                            MsgBox "�Բ���ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Exit Sub
                        End If
                        .EditText = strKey
                    ElseIf Not IsDate(strKey) Then
                        .EditText = ""
                        MsgBox "�Բ���ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Exit Sub
                    End If
                End If
                .TextMatrix(Row, Col) = strKey
            Case mconintColʵ������
                If .TextMatrix(Row, Col) = "" Or strKey = "" Then
                    MsgBox "�Բ���ʵ�������������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ���ʵ����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                
                If strKey <> "" And .TextMatrix(Row, 0) <> "" And Val(strKey) <> Val(.TextMatrix(Row, mconintColʵ������)) Then
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .EditText = strKey
                    
                    .TextMatrix(Row, mconintCol������) = zlStr.FormatEx(Abs(Val(strKey) - Val(.TextMatrix(Row, mconintCol��������))), mintNumberDigit, , True)
                    If Val(strKey) > Val(.TextMatrix(Row, mconintCol��������)) Then
                        .TextMatrix(Row, mconintCol��־) = "ӯ"
                    ElseIf Val(strKey) < Val(.TextMatrix(Row, mconintCol��������)) Then
                        .TextMatrix(Row, mconintCol��־) = "��"
                    Else
                        .TextMatrix(Row, mconintCol��־) = "ƽ"
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '���ҩƷ���������Ϊ0�������۲�Ϊ0��ҩƷ�޷�ͨ���̵��������¼������
                    '��������µ�ͨ��ҩƷ�������۵�ʵ��λ������ϵͳ���������õĽ��λ��
                    '����취�����ʵ������Ϊ0�������Ͳ�۲�С��λ�����ֺ�ҩƷ�����н��Ͳ��λ��һ��
                    If Val(.TextMatrix(Row, mconIntCol������)) = 1 Then
                        intMoneyBit = mintMoneyDigit
                    ElseIf Val(strKey) = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True And Val(.TextMatrix(Row, mconIntCol�ۼ�)) = Val(.TextMatrix(Row, mconintCol�ɱ���))) Then
                        '��0��������ҩƷ�̵�ʱ
                        intMoneyBit = mintMaxMoneyBit
                    Else
                        intMoneyBit = mintMoneyDigit
                    End If
                    
                    '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                    '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                    .TextMatrix(Row, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol�ۼ�)) * Val(strKey) - Val(.TextMatrix(Row, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                    .TextMatrix(Row, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol�ۼ�)) - Val(.TextMatrix(Row, mconintCol�ɱ���))) * Val(strKey) - Val(.TextMatrix(Row, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                    dbl���� = Val(.TextMatrix(Row, mconintCol����))
                    dbl��۲� = Val(.TextMatrix(Row, mconintCol��۲�))
                    If .TextMatrix(Row, mconintCol��־) = "��" Then
                        .TextMatrix(Row, mconintCol����) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol����)), intMoneyBit, , True)
                        .TextMatrix(Row, mconintCol��۲�) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol��۲�)), intMoneyBit, , True)
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    .TextMatrix(Row, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol�ۼ�)) * Val(strKey), mintMoneyDigit, , True)
                    .TextMatrix(Row, mconintColʵ������) = strKey
                    
                    '.TextMatrix(Row, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol�ɱ���)) * Val(.TextMatrix(Row, mconintColʵ������)), mintMoneyDigit)
                    '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                    .TextMatrix(Row, mconintCol�̵�ɱ����) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntColʵ�ʽ��)) + dbl����) - (Val(.TextMatrix(Row, mconIntColʵ�ʲ��)) + dbl��۲�), mintMoneyDigit, , True)
                    .TextMatrix(Row, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol����)) - Val(.TextMatrix(Row, mconintCol��۲�)), mintMoneyDigit, , True)
                    
                    '�̿���ӯ������ɫ����
                    Call SetStocktakingColor(vsfBill, .Row)
                End If
                
                Call ��ʾ�ϼƽ��
        Case mconintCol���װʵ������, mconintColС��װʵ������
            If .TextMatrix(Row, Col) = "" Or strKey = "" Then
                MsgBox "�Բ���ʵ�������������룡", vbOKOnly + vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If Not IsNumeric(strKey) And strKey <> "" Then
                MsgBox "�Բ���ʵ����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                Cancel = True
                Exit Sub
            End If
            If mintUnit > 0 Then
                intNumber = mintNumberDigit
            Else
                intNumber = mintNumberDigit0
            End If
               
            If strKey <> "" And .TextMatrix(Row, 0) <> "" Then
                strKey = zlStr.FormatEx(strKey, intNumber, , True)
                .EditText = strKey
                
                '�����С��װ��λ������ʵ������
                If .Col = mconintCol���װʵ������ Then
                    dblSumNum = Val(strKey) * Val(.TextMatrix(Row, mconIntCol����ϵ����)) / Val(.TextMatrix(Row, mconIntCol����ϵ��С)) + Val(.TextMatrix(Row, mconintColС��װʵ������))
                Else
                    dblSumNum = Val(.TextMatrix(Row, mconintCol���װʵ������)) * Val(.TextMatrix(Row, mconIntCol����ϵ����)) / Val(.TextMatrix(Row, mconIntCol����ϵ��С)) + Val(strKey)
                End If
                
                .TextMatrix(Row, mconintColʵ������) = zlStr.FormatEx(dblSumNum, intNumber, , True)
                .TextMatrix(Row, mconintCol�ϼ�) = .TextMatrix(Row, mconintColʵ������) & .TextMatrix(Row, mconIntColʵ��������λС)
                .TextMatrix(Row, mconintCol������) = zlStr.FormatEx(Abs(Val(.TextMatrix(Row, mconintColʵ������)) - Val(.TextMatrix(Row, mconintCol��������))), intNumber, , True)
                
                If dblSumNum > Val(.TextMatrix(Row, mconintCol��������)) Then
                    .TextMatrix(Row, mconintCol��־) = "ӯ"
                ElseIf dblSumNum < Val(.TextMatrix(Row, mconintCol��������)) Then
                    .TextMatrix(Row, mconintCol��־) = "��"
                Else
                    .TextMatrix(Row, mconintCol��־) = "ƽ"
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '���ҩƷ���������Ϊ0�������۲�Ϊ0��ҩƷ�޷�ͨ���̵��������¼������
                '��������µ�ͨ��ҩƷ�������۵�ʵ��λ������ϵͳ���������õĽ��λ��
                '����취�����ʵ������Ϊ0�������Ͳ�۲�С��λ�����ֺ�ҩƷ�����н��Ͳ��λ��һ��
                If Val(.TextMatrix(Row, mconIntCol������)) = 1 Then
                    intMoneyBit = mintMoneyDigit
                ElseIf dblSumNum = 0 Or (IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True And Val(.TextMatrix(Row, mconIntCol�ۼ�)) = Val(.TextMatrix(Row, mconintCol�ɱ���))) Then
                    intMoneyBit = mintMaxMoneyBit
                Else
                    intMoneyBit = mintMoneyDigit
                End If
                
                '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                .TextMatrix(Row, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol�ۼ�)) * dblSumNum - Val(.TextMatrix(Row, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                .TextMatrix(Row, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol�ۼ�)) - Val(.TextMatrix(Row, mconintCol�ɱ���))) * Val(dblSumNum) - Val(.TextMatrix(Row, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                dbl���� = Val(.TextMatrix(Row, mconintCol����))
                dbl��۲� = Val(.TextMatrix(Row, mconintCol��۲�))
                If .TextMatrix(Row, mconintCol��־) = "��" Then
                    .TextMatrix(Row, mconintCol����) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol����)), intMoneyBit, , True)
                    .TextMatrix(Row, mconintCol��۲�) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol��۲�)), intMoneyBit, , True)
                End If
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                .TextMatrix(Row, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol�ۼ�)) * dblSumNum, mintMoneyDigit, , True)
                '.TextMatrix(Row, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol�ɱ���)) * Val(.TextMatrix(Row, mconintColʵ������)), mintMoneyDigit)
                '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                .TextMatrix(Row, mconintCol�̵�ɱ����) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntColʵ�ʽ��)) + dbl����) - (Val(.TextMatrix(Row, mconIntColʵ�ʲ��)) + dbl��۲�), mintMoneyDigit, , True)
                .TextMatrix(Row, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol����)) - Val(.TextMatrix(Row, mconintCol��۲�)), mintMoneyDigit, , True)
                
                 '�̿���ӯ������ɫ����
                 Call SetStocktakingColor(vsfBill, .Row)
            End If
            
            Call ��ʾ�ϼƽ��
        Case mconintCol�ɱ���
            If .TextMatrix(Row, Col) = "" Or strKey = "" Then
                    MsgBox "�Բ��𣬼۸�������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬼۸����Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                
                If strKey <> "" And .TextMatrix(Row, 0) <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .EditText = strKey
                    
                    If Split(.TextMatrix(Row, mconIntcol�ӳ���), "||")(1) = 1 Then
                        'ʱ��ҩƷʱ
                        If IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                            '���۹����ۼ۵��ڳɱ���
                            .TextMatrix(Row, mconIntCol�ۼ�) = strKey
                        End If
                    Else
                        '����ҩƷ
                        If IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                            '���۹���Ҫ�жϳɱ����Ƿ�����ۼ�
                            If Val(strKey) <> Val(.TextMatrix(Row, mconIntCol�ۼ�)) Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ɱ���Ӧ���ۼ�(" & .TextMatrix(Row, mconIntCol�ۼ�) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(Row, mconIntCol�ۼ�)
                                .TextMatrix(.Row, mconintCol�ɱ���) = zlStr.FormatEx(strKey, mintCostDigit, , True)
                                .EditText = strKey
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(Row, mconIntCol������)) = 1 Then
                        intMoneyBit = mintMoneyDigit
                    ElseIf IsPriceAdjustMod(Val(.TextMatrix(Row, 0))) = True Then
                        intMoneyBit = mintMaxMoneyBit
                    Else
                        intMoneyBit = mintMoneyDigit
                    End If
                    
                    If mintUnit > 0 Then
                        dblSumNum = Val(.TextMatrix(Row, mconintColʵ������))
                    Else
                        dblSumNum = Val(.TextMatrix(Row, mconintCol���װʵ������)) * Val(.TextMatrix(Row, mconIntCol����ϵ����)) / Val(.TextMatrix(Row, mconIntCol����ϵ��С)) + Val(.TextMatrix(Row, mconintColС��װʵ������))
                    End If
                                       
                    '����=��ǰ�ۼ�*ʵ������-ʵ�ʽ��
                    '��۲�=����*iif(ʵ�ʽ��=0,ָ�������,(ʵ�ʲ��/ʵ�ʽ��))
                    .TextMatrix(Row, mconintCol����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol�ۼ�)) * dblSumNum - Val(.TextMatrix(Row, mconIntColʵ�ʽ��)), intMoneyBit, , True)
                    .TextMatrix(Row, mconintCol��۲�) = zlStr.FormatEx((Val(.TextMatrix(Row, mconIntCol�ۼ�)) - Val(strKey)) * Val(dblSumNum) - Val(.TextMatrix(Row, mconIntColʵ�ʲ��)), intMoneyBit, , True)
                    dbl���� = Val(.TextMatrix(Row, mconintCol����))
                    dbl��۲� = Val(.TextMatrix(Row, mconintCol��۲�))
                    If .TextMatrix(Row, mconintCol��־) = "��" Then
                        .TextMatrix(Row, mconintCol����) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol����)), intMoneyBit, , True)
                        .TextMatrix(Row, mconintCol��۲�) = zlStr.FormatEx(-1 * Val(.TextMatrix(Row, mconintCol��۲�)), intMoneyBit, , True)
                    End If
                                        
                    .TextMatrix(Row, mconintCol�̵���) = zlStr.FormatEx(Val(.TextMatrix(Row, mconIntCol�ۼ�)) * dblSumNum, mintMoneyDigit, , True)
                    '.TextMatrix(Row, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol�ɱ���)) * Val(.TextMatrix(Row, mconintColʵ������)), mintMoneyDigit)
                    '�ɱ����=�ɱ���*ʵ������=(������+����) -(������+��۲�) �ú�����Ϊ�˿��Ʊ������������̵㵥�ܶ���
                    .TextMatrix(Row, mconintCol�̵�ɱ����) = zlStr.FormatEx(Val(strKey) * dblSumNum, mintMoneyDigit, , True)
                    .TextMatrix(Row, mconintCol�̵�ɱ�����) = zlStr.FormatEx(Val(.TextMatrix(Row, mconintCol����)) - Val(.TextMatrix(Row, mconintCol��۲�)), mintMoneyDigit, , True)
                
                End If
        End Select
        
        If mintUnit > 0 Then
            .Cell(flexcpFontBold, Row, mconintColʵ������, Row, mconintColʵ������) = True
        Else
            .Cell(flexcpFontBold, Row, mconintCol���װʵ������, Row, mconintCol���װʵ������) = True
            .Cell(flexcpFontBold, Row, mconintColС��װʵ������, Row, mconintColС��װʵ������) = True
        End If
        
        If mblnKeyPressReturn = True Then
            vsfBill_MoveNextCell vsfBill.Row, vsfBill.Col
        End If

        mblnKeyPressReturn = False
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetҩƷ��������(ByVal intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim int�������� As Integer      '0-������;1-����
    Dim intҩ����� As Integer      '0-������;1-����
    Dim intҩ������ As Integer      '0-������;1-����
    Dim bln�Ƿ����ҩ������ As Boolean  'True-����ҩ������;False-������ҩ������
    
    If Val(vsfBill.TextMatrix(intBillRow, 0)) = 0 Then Exit Sub
    On Error GoTo ErrHandle
    strSQL = "SELECT NVL(ҩ�����, 0) ҩ�����,NVL(ҩ������, 0) ҩ������ " & _
            " From ҩƷ��� WHERE ҩƷID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡҩƷ�ⷿ��������", Val(vsfBill.TextMatrix(intBillRow, 0)))
    
    If rsTemp.RecordCount > 0 Then
        intҩ����� = rsTemp!ҩ�����
        intҩ������ = rsTemp!ҩ������
    End If
    
    If intҩ������ = 1 Then     '���ҩ�����������������Ϊ1
        int�������� = 1
    Else
        If intҩ����� = 1 Then
            strSQL = "SELECT ����ID From ��������˵�� " & _
                    " WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')) AND ����ID = [1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "ȡ��������", txtStock.Tag)
            
            bln�Ƿ����ҩ������ = (rsTemp.RecordCount > 0)
                    
            If bln�Ƿ����ҩ������ Then
                int�������� = 0
            Else
                int�������� = 1
            End If
        End If
    End If
    
    vsfBill.TextMatrix(intBillRow, mconIntCol��������) = int��������
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Get�̵�ʱ�����ۼ�(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double, ByVal date�̵�ʱ�� As Date) As Double
    '���ܣ���ȡָ��ʱ��ʱ��ҩƷ��ǰҩƷ�����ۼ�
    '����:ҩƷid,�ⷿid,����,�̵�ʱ��
    '����ֵ�����ۼ�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo ErrHandle
    '1���ж�ҩƷ�۸��¼�Ƿ�������
    gstrSQL = "select �ּ� as ���ۼ� from ҩƷ�۸��¼ where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and �۸����� = 1 and [4] between ִ������ and ��ֹ����"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
    
    If rsData.EOF Then '�޶�Ӧ��ҩƷ�۸��¼
    
        gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
            '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
            gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Get�̵�ʱ�����ۼ� = 0
            dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Get�̵�ʱ�����ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
        Else
            If rsData!���ۼ� = 0 Then
                gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
                dblָ�����ۼ� = rsData!ָ�����ۼ�
                dbl��������� = rsData!���������
                
                Get�̵�ʱ�����ۼ� = 0
                dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
                dbl�ӳ��� = rsData!�ӳ��� / 100
                dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                Get�̵�ʱ�����ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
            Else
                Get�̵�ʱ�����ۼ� = rsData!���ۼ� * dbl����ϵ��
            End If
        End If
    Else '�ж�ӦҩƷ�۸��¼
        Get�̵�ʱ�����ۼ� = rsData!���ۼ� * dbl����ϵ��
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get�̵�ʱ���ۼ�(ByVal bln�Ƿ�ʱ�� As Boolean, lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal date�̵�ʱ�� As Date) As Double
    '���ܣ���ȡԭʼ���ۼ۵�λ�ۼۣ���Ҫ���ڳ���
    '����: bln�Ƿ�ʱ��:false-����,true-ʱ��
    '����ֵ����С��λ�ļ۸�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo ErrHandle

    'ȡ����ҩƷ�ۼ�
    If bln�Ƿ�ʱ�� = False Then
        gstrSQL = "Select �ּ� " & _
            " From �շѼ�Ŀ A, ҩƷ��� B " & _
            " Where A.�շ�ϸĿid = B.ҩƷid And A.�շ�ϸĿID=[1] And to_date('" & date�̵�ʱ�� & "','YYYY-MM-DD HH24:MI:SS') Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) " & GetPriceClassString("A")
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Get�̵�ʱ���ۼ�-ȡ����ҩƷ�ۼ�", lngҩƷID)
        
        If Not rsData.EOF Then
            Get�̵�ʱ���ۼ� = rsData!�ּ�
        End If
    Else
        'ȡʱ��ҩƷ�ۼ�
        '1���ж�ҩƷ�۸��¼�Ƿ�������
        gstrSQL = "select �ּ� as ���ۼ� from ҩƷ�۸��¼ where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and �۸����� = 1 and [4] between ִ������ and ��ֹ����"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
        
        If rsData.EOF Then '�޶�Ӧ��ҩƷ�۸��¼
        
            gstrSQL = "select Decode(Nvl(���ۼ�, 0), 0, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� " & _
                " from ҩƷ��� where ����=1 and  ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lngҩƷID, lng�ⷿID, lng����)
            
            If rsData.EOF Then
                'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
                '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
                '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
                gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
                dblָ�����ۼ� = rsData!ָ�����ۼ�
                dbl��������� = rsData!���������
                
                Get�̵�ʱ���ۼ� = 0
                dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
                dbl�ӳ��� = rsData!�ӳ��� / 100
                dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                Get�̵�ʱ���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
            Else
                If rsData!���ۼ� = 0 Then
                    gstrSQL = "Select ָ�����ۼ�,nvl(�ӳ���,15) as �ӳ���,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1]"
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lngҩƷID)
                    dblָ�����ۼ� = rsData!ָ�����ۼ�
                    dbl��������� = rsData!���������
                    
                    Get�̵�ʱ���ۼ� = 0
                    dbl�ɱ��� = Get�̵�ʱ�̳ɱ���(lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
                    dbl�ӳ��� = rsData!�ӳ��� / 100
                    dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                    dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                    Get�̵�ʱ���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
                Else
                    Get�̵�ʱ���ۼ� = rsData!���ۼ�
                End If
            End If
        Else
            Get�̵�ʱ���ۼ� = rsData!���ۼ�
        End If
        
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get�̵�ʱ�̳ɱ���(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal date�̵�ʱ�� As Date) As Double
'���ܣ���ȡ��ǰҩƷ�ĳɱ��۸�
'������ҩƷid,�ⷿid,����
'����ֵ�� �ɱ��۸�
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo ErrHandle
    
    '1���ж�ҩƷ�۸��¼�Ƿ�������
    gstrSQL = "select �ּ� as �ɱ��� from ҩƷ�۸��¼ where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and �۸����� = 2 and [4] between ִ������ and ��ֹ����"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID, lng�ⷿID, lng����, date�̵�ʱ��)
    
    If rsData.EOF Then '�޶�Ӧ��ҩƷ�۸��¼
    
        gstrSQL = "select ƽ���ɱ��� from ҩƷ��� where ����=1 and ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            blnNullPrice = True
        ElseIf IsNull(rsData!ƽ���ɱ���) = True Then
            blnNullPrice = True
        ElseIf Val(rsData!ƽ���ɱ���) < 0 Then
            blnNullPrice = True
        End If
        
        If Not blnNullPrice Then
            Get�̵�ʱ�̳ɱ��� = rsData!ƽ���ɱ���
        Else
            '����޷��ӿ����ȡ�ɱ��ۣ����ҩƷ�����ȡ
            gstrSQL = "select �ɱ��� from ҩƷ��� where ҩƷid=[1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lngҩƷID)
            If Not rsData.EOF Then
                If Val(Nvl(rsData!�ɱ���, 0)) > 0 Then
                    Get�̵�ʱ�̳ɱ��� = rsData!�ɱ���
                End If
            End If
        End If
    Else
        Get�̵�ʱ�̳ɱ��� = rsData!�ɱ���
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
