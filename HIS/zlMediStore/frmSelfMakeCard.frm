VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSelfMakeCard 
   AutoRedraw      =   -1  'True
   Caption         =   "ҩƷ������ⵥ"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmSelfMakeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   11
      Top             =   5970
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   10
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   9
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   7
      Top             =   5880
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   8
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5805
      Left            =   0
      ScaleHeight     =   5745
      ScaleWidth      =   11655
      TabIndex        =   12
      Top             =   0
      Width           =   11715
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDrug 
         Height          =   3000
         Left            =   600
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   5292
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   1230
         Left            =   195
         TabIndex        =   4
         Top             =   945
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   2170
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   6
         Top             =   4920
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin ZL9BillEdit.BillEdit mshStructure 
         Height          =   1935
         Left            =   120
         TabIndex        =   28
         Top             =   2640
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3413
         Enabled         =   -1  'True
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   27
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:asdfasdfasdfsadfsadfsdfasdfsadfasdfsdf"
         Height          =   180
         Left            =   2040
         TabIndex        =   26
         Top             =   2280
         Width           =   4590
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   2280
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7230
         TabIndex        =   23
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   9330
         TabIndex        =   22
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   21
         Top             =   5280
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   20
         Top             =   5280
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   5
         Top             =   4995
         Width           =   645
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ������ⵥ"
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
         TabIndex        =   17
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   16
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   15
         Top             =   5340
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   6645
         TabIndex        =   14
         Top             =   5340
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   8520
         TabIndex        =   13
         Top             =   5340
         Width           =   720
      End
      Begin VB.Label LblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƽ���(&T)"
         Height          =   180
         Left            =   8220
         TabIndex        =   2
         Top             =   660
         Width           =   810
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
            Picture         =   "frmSelfMakeCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1000
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
            Picture         =   "frmSelfMakeCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelfMakeCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   6615
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSelfMakeCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13758
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSelfMakeCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmSelfMakeCard.frx":3080
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
      Caption         =   "����"
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
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
Attribute VB_Name = "frmSelfMakeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mstrPrivs As String                     'Ȩ��
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private mbln�¿������� As Boolean           '��Ƿ��¿�������
Private mstrWay�ɱ���   As String            '�ɱ�����Դ��ʽ        0-����ԭ��ҩƷ�ĳɱ��ۼ��㣨Ĭ�ϣ���1-��������ҩƷ���һ�����ȷ��
Private mcolUseCount As Collection
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��
Private mbln��ʾ��ʽ As Boolean             '��ʾ��ʽ true-ֻ��ʾһ�Σ�false-������ʾ

Private mlng���ⷿ As Long
Private mlng�Ƽ��� As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mintStruUnit As Integer             '�����Ƽ���

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���

Private Const MStrCaption As String = "ҩƷ����������"

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

'�Ӳ�������ȡҩƷ�۸����������С��λ�� ����
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

'�������ҩƷ�����ҩƷ���ۼ۵�λ��ȡС��λ�� ����
Private mintStruCostDigit As Integer        '�ɱ���С��λ��
Private mintStruPriceDigit As Integer       '�ۼ�С��λ��
Private mintStruNumberDigit As Integer      '����С��λ��
Private mintStruMoneyDigit As Integer       '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

'=========================================================================================

Private Const mconIntColҩ�� As Integer = 1
Private Const mconIntCol��Ʒ�� As Integer = 2
Private Const mconIntCol��� As Integer = 3
Private Const mconIntCol����ҩ�� As Integer = 4
Private Const mconIntColԭ���� As Integer = 5
Private Const mconIntCol����ϵ�� As Integer = 6
Private Const mconIntCol��λ As Integer = 7
Private Const mconIntCol���� As Integer = 8
Private Const mconIntColЧ�� As Integer = 9
Private Const mconIntCol���� As Integer = 10
Private Const mconIntCol�ɹ��� As Integer = 11
Private Const mconIntCol�ɹ���� As Integer = 12
Private Const mconintColƫ��ɱ���� As Integer = 13
Private Const mconIntCol�ۼ� As Integer = 14
Private Const mconIntCol�ۼ۽�� As Integer = 15
Private Const mconintCol��� As Integer = 16
Private Const mconIntColҩƷ��������� = 17
Private Const mconIntColҩƷ���� = 18
Private Const mconIntColҩƷ���� = 19
Private Const mconIntCol����ϵ�� = 20
Private Const mconIntColS As Integer = 21       '������
'=========================================================================================


'=========================================================================================
'����ҩƷ����
Private Const mconIntCol��ҩ�� As Integer = 0
Private Const mconIntCol����Ʒ�� As Integer = 1
Private Const mconIntCol����� As Integer = 2
Private Const mconIntCol������ As Integer = 3
Private Const mconIntCol����λ As Integer = 4
Private Const mconIntCol������ As Integer = 5
Private Const mconIntCol��������� As Integer = 6
Private Const mconIntCol���������� As Integer = 7
Private Const mconintCol���������� As Integer = 8
Private Const mconIntcol�ӳ��� As Integer = 9
Private Const mconintcol��ʵ�ʲ�� As Integer = 10
Private Const mconintcol��ʵ�ʽ�� As Integer = 11
Private Const mconintcol��ҩƷid As Integer = 12
Private Const mconIntCol���ɹ��� As Integer = 13
Private Const mconIntCol���ɹ���� As Integer = 14
Private Const mconIntCol���ۼ� As Integer = 15
Private Const mconIntCol���ۼ۽�� As Integer = 16
Private Const mconintCol����� As Integer = 17
Private Const mconIntCol��ҩƷ��������� = 18
Private Const mconIntCol��ҩƷ���� = 19
Private Const mconIntCol��ҩƷ���� = 20
Private Const mconIntCol������ϵ�� = 21
Private Const mconintColRalation = 22   '��¼����ҩ��Ӧ��ԭ��ҩ������������ҩƷʱɾ��������ҩƷ��Ӧ��ԭ��ҩƷ
Private Const mconInt��ColS As Integer = 23             '������
'=========================================================================================

Private Sub SetDrugName(ByVal intType As Integer)
    'ҩƷ������ʾ��
    'intType��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
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
    
    With mshStructure
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntCol��ҩ��) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntCol��ҩ��) = .TextMatrix(lngRow, mconIntCol��ҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntCol��ҩ��) = .TextMatrix(lngRow, mconIntCol��ҩƷ����)
                Else
                    .TextMatrix(lngRow, mconIntCol��ҩ��) = .TextMatrix(lngRow, mconIntCol��ҩƷ���������)
                End If
            End If
        Next
    End With
End Sub
Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
                
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = n
                !ҩƷID = Val(mshBill.TextMatrix(n, 0))
                               
                .Update
            End If
        Next
        
    End With
End Sub
Private Sub GetSysParm()
    mbln�¿������� = (gtype_UserSysParms.P96_ҩƷ��¿��ÿ�� = 1)
End Sub
'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim strsql As String
    
    On Error GoTo errHandle
    GetDepend = False
    strsql = "SELECT B.Id,b.���� " _
           & "FROM ҩƷ�������� A, ҩƷ������ B " _
           & "Where A.���id = B.ID AND A.���� = 2 AND B.ϵ�� = 1 "
    Set rsDepend = zldatabase.OpenSQLRecord(strsql, MStrCaption)
    
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ������������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    strsql = "SELECT B.Id,b.���� " _
           & "FROM ҩƷ�������� A, ҩƷ������ B " _
           & "Where A.���id = B.ID AND A.���� = 2  and b.ϵ�� = -1 "
    Set rsDepend = zldatabase.OpenSQLRecord(strsql, MStrCaption)
    
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ�������ĳ����������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    strsql = "SELECT DISTINCT a.id, a.���� " _
           & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
           & "Where (a.վ�� = [1] Or a.վ�� is Null) And c.�������� = b.���� " _
           & "  AND b.���� ='K'" _
           & "  AND a.id = c.����id " _
           & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
    Set rsDepend = zldatabase.OpenSQLRecord(strsql, MStrCaption, gstrNodeNo)
    
    If rsDepend.EOF Then
        MsgBox "����������û������Ϊ�Ƽ��ҵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    rsDepend.Close
    strsql = " SELECT a.����ҩƷid FROM ����ҩƷ���� a,ҩƷ��� b Where a.����ҩƷid = b.ҩƷid "
    Set rsDepend = zldatabase.OpenSQLRecord(strsql, MStrCaption)
    
    If rsDepend.EOF Then
        MsgBox "û��һ�־���ԭ��ҩ��ɵ�����ҩƷ,��鿴ҩƷĿ¼����", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    rsDepend.Close
    
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1301)
    
    Set mfrmMain = FrmMain
    
    If mint�༭״̬ = 1 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If Not zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    End If

    If Not GetDepend Then Exit Sub
   
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub


Private Sub cboStock_Click()
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
    
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���ӦҩƷ�ĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '����ҩƷ��λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                                    
                    mlng���ⷿ = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng���ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    
    With mshBill
        .SetFocus
        .Row = 1
        .Col = mconIntColҩ��
    End With
        
End Sub

Private Sub cboType_Validate(Cancel As Boolean)
    mlng�Ƽ��� = Me.cboType.ItemData(Me.cboType.ListIndex)
    Call GetDrugDigit(mlng�Ƽ���, MStrCaption, mintStruUnit, mintStruCostDigit, mintStruPriceDigit, mintStruNumberDigit, mintStruMoneyDigit)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntColҩƷ���������, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    mblnChange = False
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
    End Select
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zldatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntColҩ��, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    Dim intLop As Integer
    
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
        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub
        
        With mshBill
            For intLop = 1 To .rows - 1
                '���۹�������Ƿ���ڲ��������۵�ҩƷ������ҩ
                If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                    If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                        If Val(.TextMatrix(intLop, mconIntCol�ɹ���)) <> Val(.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                            MsgBox "��" & intLop & "������ҩƷ���������۹�������ⵥ���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End With
        
        '���۹�������Ƿ���ڲ��������۵�ҩƷ��ԭ��ҩ
        With mshStructure
            For intLop = 1 To .rows - 1
                '���۹�������Ƿ���ڲ��������۵�ҩƷ��ԭ��ҩ
                If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                    If IsPriceAdjustMod(Val(.TextMatrix(intLop, mconintcol��ҩƷid))) = True Then
                        If Val(.TextMatrix(intLop, mconIntCol���ɹ���)) <> Val(.TextMatrix(intLop, mconIntCol���ۼ�)) Then
                            MsgBox "��" & intLop & "��ԭ��ҩƷ���������۹������ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                            mshStructure.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End With
        
        If SaveCheck = True Then
            If Val(zldatabase.GetPara("��˴�ӡ", glngSys, ģ���.�������)) = 1 Then
                '��ӡ
                If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
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
            
        If Val(zldatabase.GetPara("���̴�ӡ", glngSys, ģ���.�������)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
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
    mshBill.ClearBill
    mshStructure.ClearBill
    Call ��ʾ�ϼƽ��
    SetEdit
    
    txtժҪ.Text = ""
    cboType.SetFocus
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Load()
    Dim rsMakeDrugDepart As New Recordset
    
    On Error GoTo errHandle
    mstrWay�ɱ��� = zldatabase.GetPara("ҩƷ�������ɱ��ۼ��㷽ʽ", glngSys, ģ���.�������)
    mintBatchNoLen = GetBatchNoLen()
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ����������", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    Call GetSysParm
    
    With cboType
        gstrSQL = "SELECT DISTINCT a.id, a.���� " _
                & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
                & "Where (a.վ�� = [1] Or a.վ�� is Null) And c.�������� = b.���� " _
                & " AND b.���� ='K'" _
                & " AND a.id = c.����id " _
                & " AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'"
'        Call SQLTest(App.Title, MStrCaption, gstrSQL)
'        rsMakeDrugDepart.Open gstrSQL, gcnOracle
'        Call SQLTest
        Set rsMakeDrugDepart = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, gstrNodeNo)
        
        If rsMakeDrugDepart.EOF Then Exit Sub
        .Clear
        Do While Not rsMakeDrugDepart.EOF
            .AddItem rsMakeDrugDepart.Fields(1)
            .ItemData(.NewIndex) = rsMakeDrugDepart.Fields(0)
            rsMakeDrugDepart.MoveNext
        Loop
        rsMakeDrugDepart.Close
        .ListIndex = 0
    End With
    
    mlng�Ƽ��� = Me.cboType.ItemData(Me.cboType.ListIndex)
    mlng���ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    Call GetDrugDigit(mlng�Ƽ���, MStrCaption, mintStruUnit, mintStruCostDigit, mintStruPriceDigit, mintStruNumberDigit, mintStruMoneyDigit)
    Call GetDrugDigit(mlng���ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(2, mstr���ݺ�)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '����Ȩ���жϣ��Ƿ���ʾ�ɱ���
    mshBill.ColWidth(mconIntCol�ɹ���) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol�ɹ����) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol���) = IIf(mblnViewCost, 900, 0)
    mshStructure.ColWidth(mconIntCol���ɹ���) = IIf(mblnViewCost, 1000, 0)
    mshStructure.ColWidth(mconIntCol���ɹ����) = IIf(mblnViewCost, 900, 0)
    mshStructure.ColWidth(mconintCol�����) = IIf(mblnViewCost, 900, 0)
    
    If mstrWay�ɱ��� = 1 Then
        mshBill.ColWidth(mconintColƫ��ɱ����) = 0
        mshStructure.ColWidth(mconintCol����������) = 0
    End If
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
        mshStructure.ColWidth(mconIntCol����Ʒ��) = IIf(mshStructure.ColWidth(mconIntCol����Ʒ��) = 0, 2000, mshStructure.ColWidth(mconIntCol����Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
        mshStructure.ColWidth(mconIntCol����Ʒ��) = 0
    End If
    Call mshbill_EnterCell(1, 1)
    
'    mshBill.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim rsStructure As ADODB.Recordset
    Dim rstemp As ADODB.Recordset
    Dim strUnitQuantity As String
    Dim str��װϵ�� As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPricedigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    '�ⷿ
    On Error GoTo errHandle
    strOrder = zldatabase.GetPara("����", glngSys, ģ���.�������)
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
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    intCostDigit = mintCostDigit
    intPricedigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
            Next
            mintcboIndex = .ListIndex
            cboStock.ListIndex = .ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û�����
            Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        
        Case 2, 3, 4
                
            initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "select b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id and A.���� = 2 and a.no=[1] "
                Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
                
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsInitCard!����
                    .ItemData(.NewIndex) = rsInitCard!id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint�ۼ۵�λ
                    strUnitQuantity = "D.���㵥λ AS ��λ, A.��д���� AS ����,'1' as ����ϵ��,"
                    str��װϵ�� = "1"
                Case mconint���ﵥλ
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ����,B.�����װ as ����ϵ��, "
                    str��װϵ�� = "B.�����װ"
                Case mconintסԺ��λ
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ����,B.סԺ��װ as ����ϵ��,"
                    str��װϵ�� = "B.סԺ��װ"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ����, b.ҩ���װ as ����ϵ��, "
                    str��װϵ�� = "B.ҩ���װ"
            End Select

            gstrSQL = " SELECT * FROM (SELECT DISTINCT ���,A.ҩƷID,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��,D.���,A.����, A.����, A.Ч��," & _
                strUnitQuantity & _
                " (A.�ɱ���*" & str��װϵ�� & ") AS �ɱ���,A.�ɱ���� ," & _
                " (A.���ۼ�*" & str��װϵ�� & ") AS ���ۼ�,A.���۽�� AS ���۽��," & _
                " A.��� AS ���,A.������,A.��������,A.�����,A.�������,A.ժҪ,D.���� AS ԭ����," & _
                " B.����ҩ��,B.���Ч��,A.�Է�����ID,D.�Ƿ���,B.�ӳ���,B.ҩ������ As ҩ����������,B.����ϵ��,nvl(a.����,0) as ƫ��ɱ���� " & _
                " FROM ҩƷ�շ���¼ A, ҩƷ��� B, �շ���Ŀ���� E, �շ���ĿĿ¼ D " & _
                " WHERE A.ҩƷID = B.ҩƷID And B.ҩƷID=D.ID " & _
                " AND B.ҩƷID = E.�շ�ϸĿID(+) And E.����(+)=3 " & _
                " AND A.��¼״̬ = [2] " & _
                " AND A.���� = 2 AND A.���ϵ��=1 " & _
                " AND A.NO = [1])" & _
                " ORDER BY " & strSqlOrder
            Set rsInitCard = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, mint��¼״̬)
            
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
            
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsInitCard!�Է�����id Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
            End With
            
            With mshBill
                Do While Not rsInitCard.EOF
                    
                    intRow = rsInitCard.AbsolutePosition
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard!ҩƷID
                    
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
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    Else
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = IIf(IsNull(rsInitCard!����ҩ��), "", rsInitCard!����ҩ��)

                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    .TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(rsInitCard!����, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(rsInitCard!�ɱ���, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(rsInitCard!�ɱ����, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPricedigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintColƫ��ɱ����) = zlStr.FormatEx(rsInitCard!ƫ��ɱ����, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol���������) = ""
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�ӳ��� & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    
                    .RowData(intRow) = rsInitCard!���
                    
                    'ԭ��ҩ�б�ֵ
                    If .TextMatrix(intRow, 0) <> "" Then
                        mshStructure.Redraw = False
                        
                        gstrSQL = "Select Distinct a.ҩƷid, '[' || f.���� || ']' As ����, f.���� As ͨ������, e.���� As ��Ʒ����, f.���, " & _
                            " a.����, f.���㵥λ As ��λ, a.ʵ������, a.�ɱ���,a.�ɱ����, a.���ۼ�, a.���۽��, a.���, Nvl(a.����, 0) As �������, c.���� / c.��ĸ As ���, b.����ϵ�� " & _
                            " From ҩƷ�շ���¼ A, ҩƷ��� B, ����ҩƷ���� C, �շ���Ŀ���� E, �շ���ĿĿ¼ F " & _
                            " Where a.ҩƷid = b.ҩƷid And a.ҩƷid = c.ԭ��ҩƷid And b.ҩƷid = f.Id And b.ҩƷid = e.�շ�ϸĿid(+) And e.����(+) = 3 And e.����(+) = 1 And " & _
                            " a.No = [1] And a.���� = 2 And a.��¼״̬ = [3] And a.���ϵ�� = -1 And a.���� = [4] And a.����id = [2] And c.����ҩƷid = [5] "
                            
                        Set rsStructure = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNo.Tag, Val(.TextMatrix(intRow, 0)), mint��¼״̬, mshBill.RowData(intRow), Val(.TextMatrix(intRow, 0)))
                        
                        If rsStructure.EOF Then
                            mshStructure.Redraw = True
                            Exit Sub
                        End If
                        With mshStructure
                            Do While Not rsStructure.EOF
                                .rows = .rows + 1
                                .TextMatrix(.rows - 1, mconintColRalation) = Val(mshBill.TextMatrix(intRow, 0)) 'ԭ��ҩƷ��Ӧ������ҩƷ
                                
                                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                                    strҩ�� = rsStructure!ͨ������
                                Else
                                    strҩ�� = IIf(IsNull(rsStructure!��Ʒ����), rsStructure!ͨ������, rsStructure!��Ʒ����)
                                End If
                                                                
                                .TextMatrix(.rows - 1, mconIntCol��ҩƷ���������) = rsStructure!���� & strҩ��
                                .TextMatrix(.rows - 1, mconIntCol��ҩƷ����) = rsStructure!����
                                .TextMatrix(.rows - 1, mconIntCol��ҩƷ����) = strҩ��
                                
                                If mintDrugNameShow = 0 Then
                                    .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ���������)
                                ElseIf mintDrugNameShow = 1 Then
                                    .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ����)
                                Else
                                    .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ����)
                                End If
                                
                                .TextMatrix(.rows - 1, mconIntCol����Ʒ��) = IIf(IsNull(rsStructure!��Ʒ����), "", rsStructure!��Ʒ����)
                                
                                .TextMatrix(.rows - 1, mconIntCol�����) = IIf(IsNull(rsStructure!���), "", rsStructure!���)
                                .TextMatrix(.rows - 1, mconIntCol������) = IIf(IsNull(rsStructure!����), "", rsStructure!����)
                                .TextMatrix(.rows - 1, mconIntCol����λ) = rsStructure!��λ
                                .TextMatrix(.rows - 1, mconIntCol������) = zlStr.FormatEx(rsStructure!ʵ������ - rsStructure!�������, mintStruNumberDigit, , True)
                                .TextMatrix(.rows - 1, mconIntCol���ɹ���) = zlStr.FormatEx(rsStructure!�ɱ���, mintStruCostDigit, , True)
                                .TextMatrix(.rows - 1, mconIntCol���ɹ����) = zlStr.FormatEx(IIf(IsNull(rsStructure!�ɱ����), 0, rsStructure!�ɱ����), mintStruMoneyDigit, , True)
                                .TextMatrix(.rows - 1, mconIntCol���ۼ�) = zlStr.FormatEx(rsStructure!���ۼ�, mintStruPriceDigit, , True)
                                .TextMatrix(.rows - 1, mconIntCol���ۼ۽��) = zlStr.FormatEx(IIf(IsNull(rsStructure!���۽��), 0, rsStructure!���۽��), mintStruMoneyDigit, , True)
                                .TextMatrix(.rows - 1, mconintCol�����) = zlStr.FormatEx(IIf(IsNull(rsStructure!���), 0, rsStructure!���), mintStruMoneyDigit, , True)
                                .TextMatrix(.rows - 1, mconintcol��ҩƷid) = rsStructure!ҩƷID
                                .TextMatrix(.rows - 1, mconintCol����������) = zlStr.FormatEx(rsStructure!�������, mintStruNumberDigit, , True)
                                .TextMatrix(.rows - 1, mconIntCol���������) = rsStructure!���
                                .TextMatrix(.rows - 1, mconIntCol������ϵ��) = rsStructure!����ϵ��
        
                                rsStructure.MoveNext
                            Loop
                        End With
                        rsStructure.Close
                        mshStructure.Redraw = True
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select

    SetEdit         '���ñ༭����
    Call ��ʾ�ϼƽ��
    If mint�༭״̬ = 2 And mint����� <> 0 Then
        SetUseCountCol
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'�����޸�ǰԭ��ҩ��ʹ���������Ա������޸Ĺ����жԿ���������жϸ�׼ȷ
Private Sub SetUseCountCol()
    Dim rsUseCount As New Recordset
    Dim numUsedCount As Double
    Dim vardrug As Variant
    
    On Error GoTo errHandle
    gstrSQL = "select ҩƷid,��д����,����id from ҩƷ�շ���¼ where no=[1] and ����=2 and ��¼״̬=1 and ���ϵ��=-1 "
    Set rsUseCount = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
    
    If rsUseCount.EOF Then Exit Sub
    Set mcolUseCount = New Collection
    With mcolUseCount
        
        Do While Not rsUseCount.EOF
            numUsedCount = 0
            For Each vardrug In mcolUseCount
                If vardrug(0) = rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0)) Then
                    numUsedCount = vardrug(1)
                    .Remove vardrug(0)
                    Exit For
                End If
            Next
            .Add Array(rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0)), rsUseCount.Fields(1)), rsUseCount.Fields(2) & "!" & CStr(rsUseCount.Fields(0))
            rsUseCount.MoveNext
        Loop
        rsUseCount.Close
        
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            cboStock.Enabled = False
            cboType.Enabled = False
            txtժҪ.Enabled = False
        Else
            .ColData(0) = 5
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol���) = 5
            
            .ColData(mconIntCol��λ) = 5
            .ColData(mconIntCol����) = 4
            .ColData(mconIntColЧ��) = 5
            .ColData(mconIntCol����) = 4
            If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                .ColData(mconIntCol�ɹ���) = 4
            Else
                .ColData(mconIntCol�ɹ���) = 5
            End If
            .ColData(mconIntCol�ɹ����) = 5
            .ColData(mconIntCol�ۼ�) = 5
            .ColData(mconIntCol�ۼ۽��) = 5
            .ColData(mconintCol���) = 5
            
            
            .ColData(mconIntColԭ����) = 5
            .ColData(mconIntCol����ϵ��) = 5
            
            .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol���) = flexAlignLeftCenter
            
            .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
            .ColAlignment(mconintCol���) = flexAlignRightCenter
            
            cboStock.Enabled = True
            
            cboType.Enabled = True
            txtժҪ.Enabled = True
        End If
    End With
End Sub

Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconintColƫ��ɱ����) = "ƫ��ɱ����"
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        .TextMatrix(0, mconIntColԭ����) = "ԭЧ��"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        
        .TextMatrix(1, 0) = ""
        
        .ColWidth(0) = 0
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol����) = 1100
        .ColWidth(mconintColƫ��ɱ����) = 1200
        .ColWidth(mconIntCol�ɹ���) = 1000
        .ColWidth(mconIntCol�ɹ����) = 900
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = 800
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntColҩ��) = 1
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 4
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol����) = 4
        .ColData(mconintColƫ��ɱ����) = 5
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 0
        .ColData(mconIntColԭ����) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntCol����ϵ��) = 5
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconintColƫ��ɱ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
    End With
    
    With mshStructure
        .Cols = mconInt��ColS
        
        .TextMatrix(0, mconIntCol��ҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol����Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol�����) = "���"
        .TextMatrix(0, mconIntCol������) = "����"
        .TextMatrix(0, mconIntCol����λ) = "��λ"
        .TextMatrix(0, mconIntCol������) = "����"
        .TextMatrix(0, mconIntCol���������) = "�������"
        .TextMatrix(0, mconIntCol����������) = "��������"
        .TextMatrix(0, mconintCol����������) = "��������"
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconintcol��ʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconintcol��ʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconintcol��ҩƷid) = "ҩƷid"
        .TextMatrix(0, mconIntCol���ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol���ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol���ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol���ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol�����) = "���"
        .TextMatrix(0, mconIntCol��ҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol��ҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol������ϵ��) = "����ϵ��"
        
        .ColData(mconintCol����������) = 4
        
        .ColWidth(mconIntCol��ҩ��) = 2500
        .ColWidth(mconIntCol����Ʒ��) = 2000
        .ColWidth(mconIntCol�����) = 1000
        .ColWidth(mconIntCol������) = 1000
        .ColWidth(mconIntCol����λ) = 500
        .ColWidth(mconIntCol������) = 1100
        .ColWidth(mconIntCol���������) = 0
        .ColWidth(mconIntCol����������) = 0
        .ColWidth(mconintCol����������) = 1100
        .ColWidth(mconIntcol�ӳ���) = 0
        .ColWidth(mconintcol��ʵ�ʲ��) = 0
        .ColWidth(mconintcol��ʵ�ʽ��) = 0
        .ColWidth(mconintcol��ҩƷid) = 0
        .ColWidth(mconIntCol���ɹ���) = 1000
        .ColWidth(mconIntCol���ɹ����) = 1200
        .ColWidth(mconIntCol���ۼ�) = 1000
        .ColWidth(mconIntCol���ۼ۽��) = 1200
        .ColWidth(mconintCol�����) = 1000
        .ColWidth(mconIntCol��ҩƷ���������) = 0
        .ColWidth(mconIntCol��ҩƷ����) = 0
        .ColWidth(mconIntCol��ҩƷ����) = 0
        .ColWidth(mconIntCol������ϵ��) = 0
        .ColWidth(mconintColRalation) = 0
        
        .ColAlignment(mconIntCol��ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol���ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol���ɹ����) = flexAlignRightCenter
        .ColAlignment(mconintCol�����) = flexAlignRightCenter
        .ColAlignment(mconIntCol���ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol���ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconIntCol������) = flexAlignRightCenter
        .ColAlignment(mconintCol����������) = flexAlignRightCenter
        .ColAlignment(mconIntCol�����) = flexAlignLeftCenter
        .rows = 1
    End With
    
    txtժҪ.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With
    
    
    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cboType.Left = mshBill.Left + mshBill.Width - cboType.Width
    
    LblType.Left = cboType.Left - LblType.Width - 100
    
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
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
        .Left = mshBill.Left + mshBill.Width - .Width
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
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
    End With
    
    With mshStructure
        .Left = mshBill.Left
        .Width = mshBill.Width
        .Top = txtժҪ.Top - 60 - .Height
    End With
    
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = mshStructure.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    'Pic����.Visible = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ����������", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
        
    If mshDrug.Visible Then
        mshDrug.Visible = False
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
End Sub

Private Function CheckBuildupNumStore() As String
    '�������ҩƷ��ԭ��ҩ��������Ƿ��㹻
    '����ֵ����-��ʾ�����㹻����Ϊ��-��ʾ��������
    Dim intRow As Integer
    Dim dblNum��� As Double
    Dim dblNum As Double
    Dim rstemp As ADODB.Recordset
    Dim strKey As String
    Dim collNum As Collection
    Dim vardrug As Variant
    Dim strArray As String
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim lngҩƷid As Long
    
    With mshBill
        If .rows <= 1 Then Exit Function
        
        Set collNum = New Collection
        
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, 0)) <> 0 Then
                dblNum = 0
                dblNum��� = 0
                
                gstrSQL = "Select Distinct b.ҩƷid As ԭ��ҩid, (a.���� / a.��ĸ) As ���, b.����ϵ�� As ԭ��ҩ����ϵ��, c.ʵ������ As ԭ��ҩ���" & vbNewLine & _
                    "From ����ҩƷ���� A, ҩƷ��� B, ҩƷ��� C" & vbNewLine & _
                    "Where a.ԭ��ҩƷid = b.ҩƷid And b.ҩƷid = c.ҩƷid(+) And a.����ҩƷid = [1] And c.�ⷿid = [2]"
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ���ϵ��", Val(.TextMatrix(intRow, 0)), cboType.ItemData(cboType.ListIndex))
                If rstemp.RecordCount > 0 Then
                    If rstemp!ԭ��ҩ����ϵ�� <> 0 Then
                        dblNum��� = rstemp!��� * Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)) / rstemp!ԭ��ҩ����ϵ��
                    End If
                    
                    For Each vardrug In collNum
                        If vardrug(0) = rstemp!ԭ��ҩid & "" Then
                            dblNum = vardrug(1)
                            collNum.Remove vardrug(0)
                            Exit For
                        End If
                    Next
                    strKey = rstemp!ԭ��ҩid
                    '����С��λ�����������������ʱ�����������ݱȽ�
                    strArray = dblNum + dblNum���
                    collNum.Add Array(strKey, strArray), strKey
                End If
            End If
        Next
        
        For Each varNum In collNum
            lngҩƷid = varNum(0)  '��ʽ��ҩƷid,����
            dblNum = varNum(1)
            
            'ֻ�����������ж�
            If dblNum > 0 Then
                gstrSQL = "Select (a.ʵ������ - [1]) As ʣ������, b.����" & vbNewLine & _
                            "From ҩƷ��� A, �շ���ĿĿ¼ B" & vbNewLine & _
                            "Where a.ҩƷid = b.Id And a.ҩƷid = [2] And a.�ⷿid = [3] And Nvl(a.����, 0) = [4] And b.��� In ('5', '6', '7') And a.���� = 1"
                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�����", dblNum, lngҩƷid, cboType.ItemData(cboType.ListIndex), 0)
                If rstemp.RecordCount = 0 Then
                    gstrSQL = "select ���� from �շ���ĿĿ¼ where id=[1]"
                    Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "�����", lngҩƷid)
                    CheckBuildupNumStore = rstemp!����
                    Exit Function
                Else
                    If rstemp!ʣ������ >= 0 Then
                        CheckBuildupNumStore = ""
                    Else
                        CheckBuildupNumStore = rstemp!����
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Private Function SaveCheck() As Boolean
    Dim strҩƷ As String
    
    mblnSave = False
    SaveCheck = False
    
    mstrTime_End = GetBillInfo(2, mstr���ݺ�)
    If mstrTime_End = "" Then
        MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
        Exit Function
    End If

    If mstrTime_End > mstrTime_Start Then
        MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    '�����
    strҩƷ = CheckBuildupNumStore
    If strҩƷ <> "" Then
        If mint����� = 1 Then '��������
            If MsgBox("ԭ��ҩƷ��" & strҩƷ & "����治�㣬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            Else
                mbln��ʾ��ʽ = True
            End If
        ElseIf mint����� = 2 Then '�����ֹ
            MsgBox "ԭ��ҩƷ��" & strҩƷ & "����治�㣬������ˣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    gstrSQL = "zl_�������_Verify('" & txtNo.Tag & "','" & UserInfo.�û����� & "')"
    On Error GoTo errHandle
    Call zldatabase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
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

Private Sub mshBill_AfterDeleteRow()
    Dim intRow As Integer
    
    With mshBill
        If .Row > 1 Then
            .Row = .Row - 1
        Else
            .Row = 1
        End If
        If .TextMatrix(.Row, 0) = "" Then
            mshStructure.ClearBill
        Else
'            Dim dblCostPrice As Double
'
'            If SetStructure(.TextMatrix(.Row, 0)) Then
'                If .TextMatrix(.Row, mconIntCol����) <> "" Then
'                    GetStructureNum .TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol����ϵ��), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), dblCostPrice, False
'                End If
'            End If
            
            For intRow = 1 To .rows - 1
            
            Next
        End If
    End With
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Dim intRow As Integer
    
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            Else
                 For intRow = mshStructure.rows - 1 To 1 Step -1
                    If Val(mshBill.TextMatrix(Row, 0)) = Val(mshStructure.TextMatrix(intRow, mconintColRalation)) Then
                        mshStructure.MsfObj.RemoveItem intRow
                    End If
                 Next
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As New Recordset
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim intStockID As Long
    Dim strUnitQuantity As String
    
    On Error GoTo errHandle
    
    mblnChange = True
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strUnitQuantity = "D.���㵥λ AS ��λ, trim(to_char(s.�������,'99999999990.00000')) AS ����,'1' as ����ϵ��," _
                & "trim(to_char(p.�ּ�,'99999999990.00000')) as �ۼ�,"
        Case mconint���ﵥλ
            strUnitQuantity = "d.���ﵥλ AS ��λ, trim(to_char(s.������� / d.�����װ,'99999999990.00000')) AS ����,TRIM(d.�����װ) as ����ϵ��," _
                & "trim(to_char(p.�ּ�*d.�����װ,'99999999990.00000')) as �ۼ�, "
        Case mconintסԺ��λ
            strUnitQuantity = "d.סԺ��λ AS ��λ, trim(to_char(s.������� / d.סԺ��װ,'99999999990.00000')) AS ����,TRIM(d.סԺ��װ) as ����ϵ��," _
                & "trim(to_char(p.�ּ�*d.סԺ��װ,'99999999990.00000')) as �ۼ�,"
        Case mconintҩ�ⵥλ
            strUnitQuantity = "d.ҩ�ⵥλ AS ��λ, trim(to_char(s.������� / d.ҩ���װ,'99999999990.00000')) AS ����,TRIM(d.ҩ���װ) as ����ϵ��," _
                & "trim(to_char(p.�ּ�*d.ҩ���װ,'99999999990.00000')) as �ۼ� , "
    End Select
        
    intStockID = cboStock.ItemData(cboStock.ListIndex)
    
    sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
    sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50

    gstrSQL = "" & _
        " SELECT DECODE(D.���,5,'����ҩ',6,'�г�ҩ','�в�ҩ') AS ���ʷ���,D.����,D.����,D.ͨ������,D.��Ʒ��,D.���,D.����ҩ��,D.����,D.ҩƷID," & _
             strUnitQuantity & _
        "    S.�����, D.���Ч��,D.�Ƿ���,D.�ӳ���,D.ҩ����������,E.�ⷿ��λ, D.����ϵ�� " & _
        " FROM  " & _
        "    (SELECT DISTINCT C.ҩƷ���� ����,M.���,M.����,M.���� ͨ������,A.���� ��Ʒ��," & _
        "        M.���,M.����, D.����ҩ��, D.ҩ��ID, D.ҩƷID, M.���㵥λ,NVL (TO_CHAR (D.���Ч��, '9999990'), 0) ���Ч��,D.���ﵥλ," & _
        "        TO_CHAR (D.�����װ, '999999999990.99999') �����װ,D.סԺ��λ,TO_CHAR (D.סԺ��װ, '999999999990.99999') סԺ��װ," & _
        "        D.ҩ�ⵥλ,TO_CHAR(D.ҩ���װ, '999999999990.99999') ҩ���װ,M.�Ƿ���,D.�ӳ���,D.ҩ������ AS ҩ����������, D.����ϵ�� " & _
        "    FROM ����ҩƷ���� F, ҩƷ���� C, ҩƷ��� D,�շ���ĿĿ¼ M,�շ���Ŀ���� A " & _
        "    WHERE F.����ҩƷID = D.ҩƷID AND D.ҩƷID=M.ID AND D.ҩ��ID=C.ҩ��ID " & _
        "    AND D.ҩƷID = A.�շ�ϸĿID(+) AND A.����(+)=3 AND A.����(+)=1 AND NVL(D.����ҩƷ,0)=1 And (M.վ�� = '" & gstrNodeNo & "' Or M.վ�� is Null) " & _
        "    AND (EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID =[1]) " & _
        "        OR M.��� =(SELECT DISTINCT 5 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID =[1]) " & _
        "        OR M.��� =(SELECT DISTINCT 6 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID =[1]) "
    gstrSQL = gstrSQL & _
        "        OR M.��� =(SELECT DISTINCT 7 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID =[1])) " & _
        "    AND ( EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID = [1]) " & _
        "        OR EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID =[1]) " & _
        "        OR DECODE (�������,1,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID =[1] AND ������� IN (1, 3)) " & _
        "        OR DECODE (�������,2,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID =[1] AND ������� IN (2, 3))) " & _
        "    AND ( M.����ʱ�� IS NULL OR TO_CHAR (M.����ʱ��, 'YYYY-MM-DD') = '3000-01-01') ) D,�շѼ�Ŀ P," & _
        "    (SELECT ҩƷID,TRIM(TO_CHAR(SUM(��������), '99999999999990.00000')) ��������," & _
        "        TRIM(TO_CHAR(SUM (ʵ������), '99999999999990.00000')) �������," & _
        "        TRIM(TO_CHAR(SUM (ʵ�ʽ��), '99999999999990.00')) ����� " & _
        "    FROM ҩƷ��� " & _
        "    WHERE �ⷿID =[1] AND ����=1 " & _
        "    GROUP BY ҩƷID) S,ҩƷ�����޶� E,(Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1]) F " & _
        " WHERE D.ҩƷID=P.�շ�ϸĿID AND SYSDATE BETWEEN P.ִ������ AND NVL(P.��ֹ����,SYSDATE)" & _
        GetPriceClassString("P") & _
        " AND D.ҩƷID=S.ҩƷID(+) AND D.ҩƷID=E.ҩƷID(+) And D.ҩƷid = F.�շ�ϸĿid AND E.�ⷿID(+)=[1]" & _
        " ORDER BY D.����"
    Set RecReturn = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, intStockID)
    
    If RecReturn.EOF Then Exit Sub
    Set mshDrug.Recordset = RecReturn
    RecReturn.Close
    Call SetDrugWidth(sngLeft, sngTop)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'����ҩƷѡ�����Ŀ�ȼ��������
Private Sub SetDrugWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    
    With mshDrug
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
        If RestoreFlexState(mshDrug, MStrCaption) = False Then
            .ColWidth(0) = 1000
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1000
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 0
            
            .ColWidth(8) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 0
            .ColWidth(11) = 1000
            .ColWidth(12) = 1000
            .ColWidth(13) = 1000
            .ColWidth(.Cols - 1) = 1500
        End If
        .ColAlignment(8) = flexAlignCenterCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(11) = flexAlignRightCenter
        .ColAlignment(12) = flexAlignRightCenter
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub


Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        strKey = .Text
        If strKey = "" Then
            strKey = .TextMatrix(.Row, .Col)
        End If
        
        If .Col = mconIntCol���� Or .Col = mconIntCol�ɹ��� Then
            Select Case .Col
                Case mconIntCol����
                    intDigit = mintNumberDigit
                Case mconIntCol�ɹ���
                   intDigit = mintCostDigit
            End Select
            
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
                If .SelLength = Len(strKey) Then Exit Sub
                If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= intDigit And strKey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            End If
        End If
        
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    Dim intRow As Integer
    
    With mshStructure
        If mshBill.TextMatrix(Row, 0) <> "" And Row <> 0 Then
            .Redraw = False
            For intRow = 1 To .rows - 1
                If Val(.TextMatrix(intRow, mconintColRalation)) = Val(mshBill.TextMatrix(Row, 0)) Then
                    .RowHeight(intRow) = 315
                Else
                    .RowHeight(intRow) = 0
                    .RowHeightMin = 0
                End If
            Next
            .Redraw = True
        End If
    End With
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
'        If .Row <> .LastRow Then
'            SetInputFormat .Row
'            Dim dblCostPrice As Double
'
'            If .TextMatrix(.Row, 0) <> "" Then
'                If SetStructure(.TextMatrix(.Row, 0)) <> False Then
'                    If .TextMatrix(.Row, mconIntCol����) <> "" Then
'                        GetStructureNum .TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol����ϵ��), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), dblCostPrice, False
'                    End If
'                End If
'            Else
'                mshStructure.ClearBill
'            End If
'
'        End If
        
        Select Case .Col
            Case mconIntColҩ��
                .txtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                            
            Case mconIntCol����
                .txtCheck = False
                '.TextMask = "1234567890"
                .MaxLength = mintBatchNoLen
            
            Case mconIntColЧ��
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) And .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0) <> "0" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                                    '����Ϊ��Ч��
                                    .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                                End If
                            End If
                        End If
                    End If
                End If
            Case mconIntCol�ɹ���
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol�ɹ����
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol����
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim rstemp As ADODB.Recordset
    Dim intRow As Integer
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        With mshBill
            .Text = Trim(.Text)
            strKey = Trim(.Text)
            
            If Mid(strKey, 1, 1) = "[" Then
                If InStr(2, strKey, "]") <> 0 Then
                    strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
                Else
                    strKey = Mid(strKey, 2)
                End If
            End If
            
            Select Case .Col
                Case mconIntColҩ��
                    If strKey <> "" Then
                        Dim RecReturn As New Recordset
                        Dim sngLeft As Single
                        Dim sngTop As Single
                        Dim intStockID As Long
                        
                        Select Case mintUnit
                            Case mconint�ۼ۵�λ
                                strUnitQuantity = "d.���㵥λ AS ��λ, TRIM(to_char(s.�������,'99999999999990.00000')) AS ����,'1' as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�,'99999999990.00000')) as �ۼ�,"
                            Case mconint���ﵥλ
                                strUnitQuantity = "d.���ﵥλ AS ��λ, TRIM(to_char(s.������� / d.�����װ,'99999999999990.00000')) AS ����,TRIM(d.�����װ) as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�*d.�����װ,'99999999990.00000')) as �ۼ�, "
                            Case mconintסԺ��λ
                                strUnitQuantity = "d.סԺ��λ AS ��λ, TRIM(to_char(s.������� / d.סԺ��װ,'99999999999990.00000')) AS ����,TRIM(d.סԺ��װ) as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�*d.סԺ��װ,'99999999990.00000')) as �ۼ�,"
                            Case mconintҩ�ⵥλ
                                strUnitQuantity = "d.ҩ�ⵥλ AS ��λ, TRIM(to_char(s.������� / d.ҩ���װ,'99999999999990.00000')) AS ����,TRIM(d.ҩ���װ) as ����ϵ��," _
                                    & "TRIM(to_char(p.�ּ�*d.ҩ���װ,'99999999990.00000')) as �ۼ� , "
                        End Select
                        
                        intStockID = cboStock.ItemData(cboStock.ListIndex)
                        
                        sngLeft = mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                        sngTop = mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight '  50
                        
                        gstrSQL = "" & _
                        " SELECT DECODE(D.���,5,'����ҩ',6,'�г�ҩ','�в�ҩ') AS ���ʷ���,D.����,D.����,D.ͨ������,D.��Ʒ��,D.���,D.����ҩ��,D.����,D.ҩƷID," & _
                        strUnitQuantity & _
                        "      S.�����, D.���Ч��,D.�Ƿ���,D.�ӳ���,D.ҩ����������,E.�ⷿ��λ, D.����ϵ�� " & _
                        " FROM  " & _
                        "     (SELECT DISTINCT C.ҩƷ���� ����,M.���,M.����,M.���� ͨ������,N.���� ��Ʒ��," & _
                        "         M.���,M.����, D.����ҩ��, D.ҩ��ID, D.ҩƷID, M.���㵥λ,NVL (TO_CHAR (D.���Ч��, '9999990'), 0) ���Ч��,D.���ﵥλ," & _
                        "         TO_CHAR (D.�����װ, '999999999990.99999') �����װ,D.סԺ��λ,TO_CHAR (D.סԺ��װ, '999999999990.99999') סԺ��װ," & _
                        "         D.ҩ�ⵥλ,TO_CHAR(D.ҩ���װ, '999999999990.99999') ҩ���װ,M.�Ƿ���,D.�ӳ���,D.ҩ������ AS ҩ����������, D.����ϵ�� " & _
                        "     FROM ����ҩƷ���� F, ҩƷ���� C, ҩƷ��� D,�շ���ĿĿ¼ M," & _
                        "         (Select A.* From �շ���Ŀ���� A,�շ���ĿĿ¼ B" & _
                        "     Where A.�շ�ϸĿID=B.ID And (A.���� Like [2] Or A.���� Like [2] Or B.���� Like [2]) " & _
                        "         And A.����=" & IIf(gint���뷽ʽ = 1, 2, 1) & _
                        "         And (B.վ�� = [3] Or B.վ�� is Null)) A,�շ���Ŀ���� N " & _
                        "     WHERE F.����ҩƷID = D.ҩƷID AND D.ҩƷID=M.ID And D.ҩƷID=A.�շ�ϸĿID AND D.ҩ��ID=C.ҩ��ID " & _
                        "     AND D.ҩƷID = N.�շ�ϸĿID(+) AND N.����(+)=3 AND N.����(+)=1 AND NVL(D.����ҩƷ,0)=1 " & _
                        "     AND (EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID = [1])"
                        gstrSQL = gstrSQL & _
                        "         OR M.��� =(SELECT DISTINCT 5 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID = [1] ) " & _
                        "         OR M.��� =(SELECT DISTINCT 6 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID = [1] ) " & _
                        "         OR M.��� =(SELECT DISTINCT 7 FROM ��������˵�� WHERE �������� LIKE '��ҩ%' AND ����ID = [1] )) " & _
                        "     AND ( EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID =  [1] ) " & _
                        "         OR EXISTS (SELECT 1 FROM ��������˵�� WHERE �������� = '�Ƽ���' AND ����ID = [1] ) " & _
                        "         OR DECODE (�������,1,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID = [1]  AND ������� IN (1, 3)) " & _
                        "         OR DECODE (�������,2,1,3,1,0) =(SELECT DISTINCT '1' FROM ��������˵�� WHERE �������� LIKE '%ҩ��' AND ����ID = [1]  AND ������� IN (2, 3))) " & _
                        "     AND ( M.����ʱ�� IS NULL OR TO_CHAR (M.����ʱ��, 'YYYY-MM-DD') = '3000-01-01') ) D,�շѼ�Ŀ P," & _
                        "     (SELECT ҩƷID,TO_CHAR(SUM(��������), '99999999999990.00000') ��������," & _
                        "         TO_CHAR (SUM (ʵ������), '99999999999990.00000') �������," & _
                        "         TO_CHAR (SUM (ʵ�ʽ��), '99999999999990.00') ����� " & _
                        "     FROM ҩƷ��� " & _
                        "     WHERE �ⷿID = [1]  AND ����=1 " & _
                        "     GROUP BY ҩƷID) S,ҩƷ�����޶� E,(Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1]) F " & _
                        " WHERE D.ҩƷID=P.�շ�ϸĿID AND SYSDATE BETWEEN P.ִ������ AND NVL(P.��ֹ����,SYSDATE)" & _
                        GetPriceClassString("P") & _
                        " AND D.ҩƷID=S.ҩƷID(+) AND D.ҩƷID=E.ҩƷID(+) And D.ҩƷid = F.�շ�ϸĿid AND E.�ⷿID(+)= [1] "
                        
                        Set RecReturn = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, intStockID, IIf(gstrMatchMethod = "0", "%", "") & UCase(strKey) & "%", gstrNodeNo)
                        
                        
                        If RecReturn.EOF Then
                            MsgBox "û��ƥ�������ҩƷ��", vbInformation + vbOKOnly, gstrSysName
                            RecReturn.Close
                            Cancel = True
                            Exit Sub
                        ElseIf RecReturn.RecordCount = 1 Then
                            If SetColValue(.Row, RecReturn!ҩƷID, "[" & RecReturn!���� & "]", RecReturn!ͨ������, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), IIf(IsNull(RecReturn!���), "", RecReturn!���), _
                               RecReturn!��λ, IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), _
                               IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), RecReturn!����ϵ��, RecReturn!�Ƿ���, _
                               RecReturn!�ӳ���, RecReturn!ҩ����������, RecReturn!����ϵ��, "" & RecReturn!����ҩ��) = False Then
                               RecReturn.Close
                               Cancel = True
                               Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            RecReturn.Close
                        Else
                            Set mshDrug.Recordset = RecReturn
                            RecReturn.Close
                            Call SetDrugWidth(sngLeft, sngTop)
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    Call ��ʾ�����
                    'End If
                
                Case mconIntCol����
                    '�޴���
                    If strKey = "" Then
                        If .TxtVisible = True Then
                            .TextMatrix(.Row, mconIntCol����) = ""
                        End If
                        If .ColData(mconIntColЧ��) = 2 Then
                            .Col = mconIntColЧ��
                        Else
                            .Col = mconIntCol����
                        End If
                        
                        
                        Cancel = True
                        Exit Sub
                    End If
                    
                    If Len(strKey) < 8 Then
                        MsgBox "�Բ������ų��Ȳ���������Ϊ8λ,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                Case mconIntColЧ��
                    '�д���
                    If strKey <> "" Then
                        If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                            strKey = TranNumToDate(strKey)
                            If strKey = "" Then
                                MsgBox "�Բ���ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            .Text = strKey
                            Exit Sub
                        End If
                        If Not IsDate(strKey) Then
                            MsgBox "�Բ���ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                    ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntColЧ��) Then
                    
                        If .TxtVisible = True Then
                            .Text = " "
                            Exit Sub
                        End If
                        
                        Exit Sub
                    End If
                    
                Case mconIntCol�ɹ���
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�Բ��𣬲ɹ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If strKey <> "" Then
                        strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    End If
                    .Text = strKey
                    
                    '���ý��
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɹ���) And .TextMatrix(.Row, mconIntCol����) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * strKey, mintMoneyDigit, , True)
                        
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                            If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                                If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                    '���۹���ʱʱ��ҩƷ���ۼ۵��ڳɱ���
                                    .TextMatrix(.Row, mconIntCol�ۼ�) = strKey
                                End If
                            Else
                                .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(strKey * (1 + Split(.TextMatrix(.Row, mconIntColԭ����), "||")(1) / 100), mintPriceDigit, , True)
                            End If
                            
                            .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * .TextMatrix(.Row, mconIntCol����), mintMoneyDigit, , True)
                        End If
                        
                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɹ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                    End If
                    
                    Call ��ʾ�ϼƽ��
                Case mconIntCol�ɹ����
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�Բ��𣬲ɹ�������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                                        
                    If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɹ����) Then
                        If .TextMatrix(.Row, mconIntCol����) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(strKey / .TextMatrix(.Row, mconIntCol����), mintPriceDigit, , True)
                        End If
                        
                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - strKey, mintMoneyDigit, , True)
                        .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                    End If
                    ��ʾ�ϼƽ��
                Case mconIntCol����
                    If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                        MsgBox "�Բ��������������룡", vbOKOnly + vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If Not IsNumeric(strKey) And strKey <> "" Then
                        MsgBox "�Բ�����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If strKey <> "" Then
                        If Val(strKey) = 0 Then
                            MsgBox "�Բ����������������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strKey) < 0 Then
                            If Not zlStr.IsHavePrivs(mstrPrivs, "��������") Then
                                MsgBox "�Բ�����û�и���������Ȩ�ޣ������䣡", vbInformation, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                        
                        Dim dblCostPrice As Double
                        If Val(strKey) >= 10 ^ 11 - 1 Then
                            MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If .TextMatrix(.Row, 0) = "" Then Exit Sub

                        
                        'ȡ���ҩƷ������,����������ҩƷ�Ĳɹ��� ��
                        If GetStructureNum(Val(.TextMatrix(.Row, 0)), strKey * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), dblCostPrice) = False Then
                            Cancel = True
                            Exit Sub
                        Else
                            '�޸���������Ҫ����������ȫ�����
                            .TextMatrix(.Row, mconintColƫ��ɱ����) = ""
                            For intRow = 1 To mshStructure.rows - 1
                                If Val(.TextMatrix(.Row, 0)) = Val(mshStructure.TextMatrix(intRow, mconintColRalation)) Then
                                    mshStructure.TextMatrix(intRow, mconintCol����������) = ""
                                End If
                            Next
                            
                            If mstrWay�ɱ��� = "0" Then '����ԭ��ҩ����
                                .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(dblCostPrice * .TextMatrix(.Row, mconIntCol����ϵ��), mintPriceDigit, , True)
                            Else    '�������һ�����ɱ�����
                                gstrSQL = "select �ɱ��� from ҩƷ��� where ҩƷid=[1]"
                                Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ�ɱ���", .TextMatrix(.Row, 0))
                                If rstemp.RecordCount = 0 Then
                                    Exit Sub
                                Else
                                    .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(IIf(IsNull(rstemp!�ɱ���), 0, rstemp!�ɱ���) * .TextMatrix(.Row, mconIntCol����ϵ��), mintPriceDigit, , True)
                                End If
                            End If
                        End If
                                
                        strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                        .Text = strKey
                        If .TextMatrix(.Row, mconIntCol�ɹ���) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɹ���) * strKey, mintMoneyDigit, , True)
                        End If
                        If Val(.TextMatrix(.Row, mconIntCol�ɹ����)) >= 10 ^ 14 - 1 Then
                            MsgBox "�ɹ�������С��" & (10 ^ 14 - 1) & ",��������������!", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                            If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                                If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                    '���۹���ʱʱ��ҩƷ���ۼ۵��ڳɱ���
                                    .TextMatrix(.Row, mconIntCol�ۼ�) = .TextMatrix(.Row, mconIntCol�ɹ���)
                                End If
                            Else
                                .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɹ���) * (1 + Split(.TextMatrix(.Row, mconIntColԭ����), "||")(1) / 100), mintPriceDigit, , True)
                            End If
                        End If
                        
                        If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                            .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                        End If
                        If Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) >= 10 ^ 14 - 1 Then
                            MsgBox "�ۼ۽�����С��" & (10 ^ 14 - 1) & ",��������������!", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɹ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                    
                        '���۹�������Ƿ���ڲ��������۵�ҩƷ
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                            If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                If Val(.TextMatrix(.Row, mconIntCol�ɹ���)) <> Val(.TextMatrix(.Row, mconIntCol�ۼ�)) Then
                                    MsgBox "��" & .Row & "��ҩƷ���������۹�������ⵥ���ۼۺͳɱ��۲�һ�£����ܱ��棬���飡", vbInformation + vbOKOnly, gstrSysName
                                End If
                            End If
                        End If
                    
                    End If
                    Call ��ʾ�ϼƽ��
                
            End Select
        End With
    ElseIf KeyCode = vbKeyDown And Shift = vbAltMask Then
        mshbill_CommandClick
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��������ҩƷID�ж��Ƿ��й��õ����ҩƷ�����У���������Ӧ������
Private Function GetStructureNum(ByVal lngҩƷid As Long, ByVal dblNum As Double, ByVal dbl����ϵ�� As Double, ByRef dblCostPrice As Double, _
         Optional bln�жϿ�� As Boolean = True) As Boolean
    Dim rsDrug As New Recordset
    Dim intReturn As Integer
    Dim blnContinue As Boolean      '�û���ѡ��0���˳���1����
    Dim dblConstruct As Double      'ʵ��������Ӧ���������
    Dim dblPurchase As Double       '����ҩƷ�ĳɱ��ۣ����У����ҩƷ�Ľ���*���������
    Dim intRow As Integer
    Dim numUseCount As Double
    Dim vardrug As Variant
    Dim dblԭ��д���� As Double
    Dim n As Integer
    Dim lngԭ��ҩID As Long
    Dim intStruCostDigit As Integer        '�ɱ���С��λ��
    Dim intStruNumberDigit As Integer      '����С��λ��
    Dim intStruMoneyDigit As Integer       '���С��λ��
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '           ��ۺͳɱ����ڳ��⴦���еĹ�ʽ
    '   ������=����*�ۼ�
    '   ������=������*��ʵ�ʲ��/ʵ�ʽ�
    '          ���ʵ�ʲ�ۺ�ʵ�ʽ�����ʱ��Ϊ��
    '       ������=������*ָ�������
    '   ���ۣ��ɱ���)=(������-������)/����
    '
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    intStruCostDigit = mintStruCostDigit
    intStruNumberDigit = mintStruNumberDigit
    intStruMoneyDigit = mintStruMoneyDigit
    
    GetStructureNum = False
    blnContinue = False
    With mshStructure
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" And lngҩƷid = Val(.TextMatrix(intRow, mconintColRalation)) Then
                dblConstruct = .TextMatrix(intRow, mconIntCol���������) * dblNum * dbl����ϵ�� / Val(.TextMatrix(intRow, mconIntCol������ϵ��))
                                
                .TextMatrix(intRow, mconIntCol������) = zlStr.FormatEx(dblConstruct, intStruNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol���ۼ۽��) = zlStr.FormatEx(dblConstruct * .TextMatrix(intRow, mconIntCol���ۼ�), intStruMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol���ɹ����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol���ɹ���) * dblConstruct, intStruMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol�����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol���ۼ۽��) - .TextMatrix(intRow, mconIntCol���ɹ����), intStruMoneyDigit, , True)
                dblPurchase = zlStr.FormatEx(dblPurchase + .TextMatrix(intRow, mconIntCol���ɹ����) / dblNum, intStruCostDigit, , True)
                
                'ԭ�Ͽ��ܴ�����ͬ�ģ�Ҫ�ϲ������ܵ�����
                lngԭ��ҩID = Val(.TextMatrix(intRow, mconintcol��ҩƷid))
                For n = 1 To .rows - 1
                    If .TextMatrix(n, 0) <> "" And lngԭ��ҩID = Val(.TextMatrix(n, mconintcol��ҩƷid)) And n <> intRow Then
                        dblConstruct = dblConstruct + Val(.TextMatrix(n, mconIntCol������))
                    End If
                Next
                
                '�Ƚ�ֵ���Ȼ���ټ�飬��Ϊ�˲鿴���ҩƷ��Ҫ�ö�
                If Not CheckUsableNum(cboType.ItemData(cboType.ListIndex), Val(.TextMatrix(intRow, mconintcol��ҩƷid)), 0, dblConstruct, 1, txtNo.Caption, 2, mint�����, IIf(mintStruNumberDigit >= mintNumberDigit, mintStruNumberDigit, mintNumberDigit)) Then
                    GetStructureNum = False
                    Exit Function
                End If
            End If
        Next
        dblCostPrice = dblPurchase
    End With
    
    GetStructureNum = True
End Function
'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal intҩƷid As Long, ByVal strҩƷ���� As String, _
    ByVal strͨ���� As String, ByVal str��Ʒ�� As String, _
    ByVal str��� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, _
    ByVal intԭЧ�� As Integer, ByVal num����ϵ�� As Double, _
    ByVal int�Ƿ��� As Integer, ByVal dbl�ӳ��� As Double, ByVal intҩ���������� As Integer, _
    ByVal dbl����ϵ�� As Double, ByVal str����ҩ�� As String) As Boolean
    
    Dim intCount As Integer
    Dim rsStructure As New Recordset
    Dim intCol As Integer
    Dim strҩ�� As String
    
    SetColValue = False
    With mshBill
        
        If Not SetStructure(intҩƷid) Then Exit Function
        
        For intCol = 0 To .Cols - 1
            '.TextMatrix(intRow, intCol) = ""
            '2010-5-5 ������ʱ������ֵ
            If mconIntCol���� <> intCol Or Trim(.TextMatrix(intRow, mconIntCol����)) = "" Then
                .TextMatrix(intRow, intCol) = ""
            End If
        Next
        
        .TextMatrix(intRow, 0) = intҩƷid
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = strͨ����
        Else
            strҩ�� = IIf(str��Ʒ�� <> "", str��Ʒ��, strͨ����)
        End If
        
        .TextMatrix(intRow, mconIntColҩƷ���������) = strҩƷ���� & strҩ��
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩƷ����
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        Else
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(intRow, mconIntCol��Ʒ��) = str��Ʒ��

        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol����ҩ��) = str����ҩ��
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ�, mintPriceDigit, , True)
        .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(intԭЧ��), "0", intԭЧ��) & "||" & dbl�ӳ��� & "||" & int�Ƿ��� & "||" & intҩ����������
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol����ϵ��) = dbl����ϵ��
        
        SetInputFormat intRow
        
    End With
    SetColValue = True
End Function

Private Function SetStructure(ByVal intҩƷid As Long) As Boolean
    Dim rsStructure As New Recordset
    Dim strҩ�� As String
    Dim rs�ɱ��� As ADODB.Recordset
    Dim intRow As Integer
    Dim blnDouble As Boolean '�Ƿ��Ѿ��ظ�
    
    SetStructure = False
    mshStructure.Redraw = False
    
    On Error GoTo errHandle
    If mint�༭״̬ <> 4 Then
        For intRow = 1 To mshStructure.rows - 1
            If Val(mshStructure.TextMatrix(intRow, mconintColRalation)) = intҩƷid Then
                blnDouble = True
                Exit For
            End If
        Next
             
        If blnDouble = False Then
            gstrSQL = "SELECT DISTINCT B.ҩƷID,'[' || F.���� || ']' As ����,F.���� As ͨ������,E.���� AS ��Ʒ����," & _
                " F.���,C.�ϴβ���,F.���㵥λ AS ��λ,C.ʵ�ʲ��,C.ʵ�ʽ��,D.�ּ� As �ۼ�," & _
                " (A.����/A.��ĸ) AS ���,C.��������,B.�ӳ���,F.�Ƿ���,B.ҩ������ ҩ����������, B.����ϵ��,c.ƽ���ɱ��� " & _
                " FROM ����ҩƷ���� A,ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F,ҩƷ��� C,�շѼ�Ŀ D " & _
                " WHERE A.ԭ��ҩƷID = B.ҩƷID And B.ҩƷID=F.ID AND NVL(F.�Ƿ���,0)=0" & _
                " AND A.ԭ��ҩƷID = D.�շ�ϸĿID AND (SYSDATE BETWEEN ִ������ AND NVL(��ֹ����,SYSDATE))" & _
                GetPriceClassString("D") & _
                " AND B.ҩƷID = E.�շ�ϸĿID(+) AND E.����(+)=1 And E.����(+)=3" & _
                " AND A.ԭ��ҩƷID = C.ҩƷID(+) AND C.�ⷿID(+)=[1] AND C.����(+)=1 " & _
                " AND (F.վ�� = [3] Or F.վ�� is Null) And A.����ҩƷID =[2] "
            gstrSQL = gstrSQL & " UNION " & _
                " SELECT DISTINCT B.ҩƷID,'[' || F.���� || ']' As ����,F.���� As ͨ������,E.���� AS ��Ʒ����," & _
                " F.���,C.�ϴβ���,F.���㵥λ AS ��λ,C.ʵ�ʲ��,C.ʵ�ʽ��,Decode(Nvl(C.����,0),0,C.ʵ�ʽ��/C.ʵ������,Nvl(C.���ۼ�,C.ʵ�ʽ��/C.ʵ������)) AS �ۼ�," & _
                " (A.���� / A.��ĸ) AS ���,C.��������,B.�ӳ���,F.�Ƿ���,B.ҩ������ As ҩ����������, B.����ϵ��,c.ƽ���ɱ��� " & _
                " FROM ����ҩƷ���� A,ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F,ҩƷ��� C" & _
                " WHERE A.ԭ��ҩƷID = B.ҩƷID And B.ҩƷID=F.ID AND NVL(F.�Ƿ���,0)=1 " & _
                " AND B.ҩƷID = E.�շ�ϸĿID(+) And E.����(+)=3 ANd E.����(+)=1 " & _
                " AND A.ԭ��ҩƷID = C.ҩƷID AND C.�ⷿID =[1] AND C.����=1 AND Nvl(C.ʵ������,0)>0" & _
                " AND (F.վ�� = [3] Or F.վ�� is Null) And A.����ҩƷID =[2] "
            
            Set rsStructure = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboType.ItemData(cboType.ListIndex), intҩƷid, gstrNodeNo)
            
            If rsStructure.EOF Then
                mshStructure.Redraw = True
                Exit Function
            End If
            With mshStructure
    '            .ClearBill
                Do While Not rsStructure.EOF
                    If rsStructure!ҩ���������� = 1 Then
                        MsgBox "���ҩƷ��һ��ҩ������ҩƷ������ǰ�汾��֧��ҩ�����������ҩƷ�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshStructure.Redraw = True
                        Exit Function
                    End If
                    
                    .rows = .rows + 1
                    .TextMatrix(.rows - 1, mconintColRalation) = intҩƷid 'ԭ��ҩƷ��Ӧ������ҩƷ
                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                        strҩ�� = rsStructure!ͨ������
                    Else
                        strҩ�� = IIf(IsNull(rsStructure!��Ʒ����), rsStructure!ͨ������, rsStructure!��Ʒ����)
                    End If
                                                    
                    .TextMatrix(.rows - 1, mconIntCol��ҩƷ���������) = rsStructure!���� & strҩ��
                    .TextMatrix(.rows - 1, mconIntCol��ҩƷ����) = rsStructure!����
                    .TextMatrix(.rows - 1, mconIntCol��ҩƷ����) = strҩ��
                    
                    If mintDrugNameShow = 0 Then
                        .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ���������)
                    ElseIf mintDrugNameShow = 1 Then
                        .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ����)
                    Else
                        .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ����)
                    End If
                    
                    .TextMatrix(.rows - 1, mconIntCol����Ʒ��) = IIf(IsNull(rsStructure!��Ʒ����), "", rsStructure!��Ʒ����)
                    
                    .TextMatrix(.rows - 1, mconIntCol�����) = IIf(IsNull(rsStructure!���), "", rsStructure!���)
                    .TextMatrix(.rows - 1, mconIntCol������) = IIf(IsNull(rsStructure!�ϴβ���), "", rsStructure!�ϴβ���)
                    .TextMatrix(.rows - 1, mconIntCol����λ) = rsStructure!��λ
                    .TextMatrix(.rows - 1, mconIntCol���ۼ�) = zlStr.FormatEx(rsStructure!�ۼ�, mintStruPriceDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol����������) = zlStr.FormatEx(IIf(IsNull(rsStructure!��������), "0", rsStructure!��������), mintStruNumberDigit, , True)
                    .TextMatrix(.rows - 1, mconIntCol���������) = rsStructure!���
                    .TextMatrix(.rows - 1, mconIntcol�ӳ���) = rsStructure!�ӳ��� & "||" & IIf(IsNull(rsStructure!�Ƿ���), 0, rsStructure!�Ƿ���) & "||" & IIf(IsNull(rsStructure!ҩ����������), 0, rsStructure!ҩ����������)
                    .TextMatrix(.rows - 1, mconintcol��ʵ�ʲ��) = IIf(IsNull(rsStructure!ʵ�ʲ��), "0", rsStructure!ʵ�ʲ��)
                    .TextMatrix(.rows - 1, mconintcol��ʵ�ʽ��) = IIf(IsNull(rsStructure!ʵ�ʽ��), "0", rsStructure!ʵ�ʽ��)
                    .TextMatrix(.rows - 1, mconintcol��ҩƷid) = rsStructure!ҩƷID
                    .TextMatrix(.rows - 1, mconIntCol������ϵ��) = rsStructure!����ϵ��
                    If IsNull(rsStructure!ƽ���ɱ���) Then
                        gstrSQL = "select �ɱ��� from ҩƷ��� where ҩƷid=[1]"
                        Set rs�ɱ��� = zldatabase.OpenSQLRecord(gstrSQL, "��ѯ�ɱ���", Val(rsStructure!ҩƷID))
                        If rs�ɱ���.RecordCount > 0 Then
                            .TextMatrix(.rows - 1, mconIntCol���ɹ���) = zlStr.FormatEx(rs�ɱ���!�ɱ���, mintStruCostDigit, , True)
                        End If
                    Else
                        .TextMatrix(.rows - 1, mconIntCol���ɹ���) = zlStr.FormatEx(rsStructure!ƽ���ɱ���, mintStruCostDigit, , True)
                    End If
                    
    '                If .Row = .rows - 1 Then
    '                    .rows = .rows + 1
    '                End If
    '                .Row = .Row + 1
                    rsStructure.MoveNext
                Loop
            End With
            
            rsStructure.Close
            SetStructure = True
            mshStructure.Redraw = True
        End If
    Else            '�鿴
        gstrSQL = "SELECT DISTINCT A.ҩƷID,'[' || F.���� || ']' As ����,F.���� As ͨ������,E.���� AS ��Ʒ����,F.���," & _
            " A.����, F.���㵥λ AS ��λ,A.ʵ������,A.�ɱ���,A.�ɱ����,A.���ۼ�,A.���۽��,A.���,nvl(A.����,0) as ������� " & _
            " FROM " & _
            "     (SELECT ҩƷID,����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,���� FROM ҩƷ�շ���¼ " & _
            "     WHERE NO=[1] AND ����=2 AND ��¼״̬=[3] " & _
            "     AND ���ϵ��=-1 AND ����=[4] AND ����ID =[2]) A," & _
            "     ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F " & _
            " WHERE A.ҩƷID = B.ҩƷID And B.ҩƷID=F.ID " & _
            " AND B.ҩƷID = E.�շ�ϸĿID(+) And E.����(+)=3 AND E.����(+)=1"
        
        Set rsStructure = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, txtNo.Tag, intҩƷid, mint��¼״̬, mshBill.RowData(mshBill.Row))
        
        If rsStructure.EOF Then
            mshStructure.Redraw = True
            Exit Function
        End If
        With mshStructure
'            .ClearBill
            Do While Not rsStructure.EOF
                .rows = .rows + 1
                .TextMatrix(.rows - 1, mconintColRalation) = intҩƷid 'ԭ��ҩƷ��Ӧ������ҩƷ
                
                If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                    strҩ�� = rsStructure!ͨ������
                Else
                    strҩ�� = IIf(IsNull(rsStructure!��Ʒ����), rsStructure!ͨ������, rsStructure!��Ʒ����)
                End If
                                                
                .TextMatrix(.rows - 1, mconIntCol��ҩƷ���������) = rsStructure!���� & strҩ��
                .TextMatrix(.rows - 1, mconIntCol��ҩƷ����) = rsStructure!����
                .TextMatrix(.rows - 1, mconIntCol��ҩƷ����) = strҩ��
                
                If mintDrugNameShow = 0 Then
                    .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ���������)
                ElseIf mintDrugNameShow = 1 Then
                    .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ����)
                Else
                    .TextMatrix(.rows - 1, mconIntCol��ҩ��) = .TextMatrix(.rows - 1, mconIntCol��ҩƷ����)
                End If
                
                .TextMatrix(.rows - 1, mconIntCol����Ʒ��) = IIf(IsNull(rsStructure!��Ʒ����), "", rsStructure!��Ʒ����)
                
                .TextMatrix(.rows - 1, mconIntCol�����) = IIf(IsNull(rsStructure!���), "", rsStructure!���)
                .TextMatrix(.rows - 1, mconIntCol������) = IIf(IsNull(rsStructure!����), "", rsStructure!����)
                .TextMatrix(.rows - 1, mconIntCol����λ) = rsStructure!��λ
                .TextMatrix(.rows - 1, mconIntCol������) = zlStr.FormatEx(rsStructure!ʵ������, mintStruNumberDigit, , True)
                .TextMatrix(.rows - 1, mconIntCol���ɹ���) = zlStr.FormatEx(rsStructure!�ɱ���, mintStruCostDigit, , True)
                .TextMatrix(.rows - 1, mconIntCol���ɹ����) = zlStr.FormatEx(IIf(IsNull(rsStructure!�ɱ����), 0, rsStructure!�ɱ����), mintStruMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconIntCol���ۼ�) = zlStr.FormatEx(rsStructure!���ۼ�, mintStruPriceDigit, , True)
                .TextMatrix(.rows - 1, mconIntCol���ۼ۽��) = zlStr.FormatEx(IIf(IsNull(rsStructure!���۽��), 0, rsStructure!���۽��), mintStruMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintCol�����) = zlStr.FormatEx(IIf(IsNull(rsStructure!���), 0, rsStructure!���), mintStruMoneyDigit, , True)
                .TextMatrix(.rows - 1, mconintcol��ҩƷid) = rsStructure!ҩƷID
                .TextMatrix(.rows - 1, mconintCol����������) = zlStr.FormatEx(rsStructure!�������, mintStruNumberDigit, , True)
                
'                If .Row = .rows - 1 Then
'                    .rows = .rows + 1
'                End If
'                .Row = .Row + 1
                rsStructure.MoveNext
            Loop
                
        End With
        rsStructure.Close
        mshStructure.Redraw = True
        Exit Function
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    If mblnEdit = False Then Exit Sub
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then
            .ColData(mconIntColЧ��) = 5
            Exit Sub
        End If
        
        If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
            If Split(.TextMatrix(intRow, mconIntColԭ����), "||")(0) = "0" Then
                .ColData(mconIntColЧ��) = 5
            Else
                .ColData(mconIntColЧ��) = 2                '���������
            End If
        Else
            .ColData(mconIntColЧ��) = 5
        End If
    End With
End Sub


Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntColҩ�� Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub mshDrug_DblClick()
    mshDrug_KeyPress 13
End Sub

Private Sub mshDrug_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshDrug
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For i = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(i)
                If sngWidth > .Width Then
                    .LeftCol = i + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub mshDrug_KeyPress(KeyAscii As Integer)
    With mshDrug
        If KeyAscii = 13 Then
            If Not SetColValue(mshBill.Row, .TextMatrix(.Row, 8), "[" & .TextMatrix(.Row, 2) & "]", .TextMatrix(.Row, 3), .TextMatrix(.Row, 4), _
                 .TextMatrix(.Row, 5), .TextMatrix(.Row, 9), Val(.TextMatrix(.Row, 12)), _
                 IIf(IsNull(.TextMatrix(.Row, 14)), "0", .TextMatrix(.Row, 14)), .TextMatrix(.Row, 11), Val(.TextMatrix(.Row, 15)), _
                 Val(.TextMatrix(.Row, 16)), Val(.TextMatrix(.Row, 17)), Val(.TextMatrix(.Row, 19)), .TextMatrix(.Row, 6)) Then
                mshBill.SetFocus
                mshBill.Col = mconIntColҩ��
                .Visible = False
                Exit Sub
            End If
            .Visible = False
            mshBill.Text = "[" & .TextMatrix(.Row, 2) & "]" & .TextMatrix(.Row, 4)
            
            mshBill.Col = mconIntCol����
            
            mshBill.SetFocus
        End If
    End With
                
            
End Sub

Private Sub mshDrug_LostFocus()
    SaveFlexState mshDrug, MStrCaption
    If mshDrug.Visible Then mshDrug.Visible = False
End Sub

Private Sub mshStructure_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub


Private Sub mshStructure_EnterCell(Row As Long, Col As Long)
    Call ��ʾԭ�Ͽ����
    With mshStructure
        If Row = 0 Then
            If mconintCol���������� = Col Then
                .ColData(mconintCol����������) = 0
            End If
            Exit Sub
        End If
        Select Case Col
            Case mconintCol����������
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
                .ColData(mconintCol����������) = 4
        End Select
    End With
End Sub

Private Sub mshStructure_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With mshStructure
        strKey = Trim(.Text)
        Select Case .Col
            Case mconintCol����������
                If Not IsNumeric(strKey) Then
                    KeyCode = 0
                    Cancel = True
                    .Text = ""
                    Exit Sub
                End If
                strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                
                '�Ƚ�ֵ���Ȼ���ټ�飬��Ϊ�˲鿴���ҩƷ��Ҫ�ö�
                If Not CheckUsableNum(cboType.ItemData(cboType.ListIndex), Val(.TextMatrix(.Row, mconintcol��ҩƷid)), 0, strKey + Val(.TextMatrix(.Row, mconIntCol������)), 1, txtNo.Caption, 2, mint�����, IIf(mintStruNumberDigit >= mintNumberDigit, mintStruNumberDigit, mintNumberDigit)) Then
                    KeyCode = 0
                    Cancel = True
                    .Text = ""
                    Exit Sub
                End If
                
                .Text = strKey
                Call GetCalcPrice(strKey, .Row)
        End Select
    End With
End Sub

Private Sub GetCalcPrice(ByVal dblNum As Double, ByVal intRow As Integer)
    '¼��ƫ���ʱ������ɱ������ۼ۽��ĺ���
    '���� dblNum����������
    '   ��introw���������
    '   :  intClass 0����ɱ����,1�����ۼ۽��
    Dim i As Integer
    Dim dblALLMoney As Double '�ۼ۽��
    Dim dblAllCost As Double '�ɱ����
    Dim dblPianChaCost As Double  'ƫ��ɱ����
    
    With mshStructure
        '�ɱ����
        .TextMatrix(intRow, mconIntCol���ɹ����) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol������)) + dblNum) * Val(.TextMatrix(intRow, mconIntCol���ɹ���)), mintStruMoneyDigit, , True)
        '�ۼ۽��
        .TextMatrix(intRow, mconIntCol���ۼ۽��) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconIntCol������)) + dblNum) * Val(.TextMatrix(intRow, mconIntCol���ۼ�)), mintStruMoneyDigit, , True)
        '���
        .TextMatrix(intRow, mconintCol�����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol���ۼ۽��)) - Val(.TextMatrix(intRow, mconIntCol���ɹ����)), mintStruMoneyDigit, , True)
        '��������ҩ�������
        For i = 1 To .rows - 1
            If .TextMatrix(intRow, mconintColRalation) = .TextMatrix(i, mconintColRalation) Then
                dblAllCost = dblAllCost + Val(.TextMatrix(i, mconIntCol���ɹ����))
                dblALLMoney = dblALLMoney + Val(.TextMatrix(i, mconIntCol���ۼ۽��))
                
                If i = intRow Then
                    dblPianChaCost = dblPianChaCost + (Val(.TextMatrix(i, mconIntCol���ɹ���)) * dblNum)
                Else
                    dblPianChaCost = dblPianChaCost + (Val(.TextMatrix(i, mconIntCol���ɹ���)) * Val(.TextMatrix(i, mconintCol����������)))
                End If
            End If
        Next
        For i = 1 To mshBill.rows - 1
            If Val(mshBill.TextMatrix(i, 0)) = Val(.TextMatrix(intRow, mconintColRalation)) Then
                If Val(mshBill.TextMatrix(i, mconIntCol����)) <> 0 Then
                    mshBill.TextMatrix(i, mconIntCol�ɹ���) = zlStr.FormatEx(dblAllCost / Val(mshBill.TextMatrix(i, mconIntCol����)), mintCostDigit, , True)
                    mshBill.TextMatrix(i, mconIntCol�ɹ����) = zlStr.FormatEx(dblAllCost, mintMoneyDigit, , True)
'                    mshBill.TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(dblALLMoney / Val(mshBill.TextMatrix(i, mconIntCol����)), mintPriceDigit, , True)
'                    mshBill.TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(dblALLMoney, mintMoneyDigit, , True)
                    mshBill.TextMatrix(i, mconintCol���) = zlStr.FormatEx(Val(mshBill.TextMatrix(i, mconIntCol�ۼ۽��)) - Val(mshBill.TextMatrix(i, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                    mshBill.TextMatrix(i, mconintColƫ��ɱ����) = zlStr.FormatEx(dblPianChaCost, mintMoneyDigit, , True)
                Else
                    mshBill.TextMatrix(i, mconIntCol�ɹ���) = zlStr.FormatEx(dblAllCost / 1, mintCostDigit, , True)
                    mshBill.TextMatrix(i, mconIntCol�ɹ����) = zlStr.FormatEx(dblAllCost, mintMoneyDigit, , True)
'                    mshBill.TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(dblALLMoney / 1, mintPriceDigit, , True)
'                    mshBill.TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(dblALLMoney, mintMoneyDigit, , True)
                    mshBill.TextMatrix(i, mconintCol���) = zlStr.FormatEx(Val(mshBill.TextMatrix(i, mconIntCol�ۼ۽��)) - Val(mshBill.TextMatrix(i, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                    mshBill.TextMatrix(i, mconintColƫ��ɱ����) = zlStr.FormatEx(dblPianChaCost, mintMoneyDigit, , True)
                End If
                
                If Split(mshBill.TextMatrix(i, mconIntColԭ����), "||")(2) = 1 Then
                    mshBill.TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(Val(mshBill.TextMatrix(i, mconIntCol�ɹ���)) * (1 + Split(mshBill.TextMatrix(i, mconIntColԭ����), "||")(1) / 100), mintPriceDigit, , True)
                    mshBill.TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(mshBill.TextMatrix(i, mconIntCol�ۼ�) * mshBill.TextMatrix(i, mconIntCol����), mintMoneyDigit, , True)
                    mshBill.TextMatrix(i, mconintCol���) = zlStr.FormatEx(IIf(mshBill.TextMatrix(i, mconIntCol�ۼ۽��) = "", 0, mshBill.TextMatrix(i, mconIntCol�ۼ۽��)) - IIf(mshBill.TextMatrix(i, mconIntCol�ɹ����) = "", 0, mshBill.TextMatrix(i, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                End If
                
                Exit For
            End If
        Next
    End With
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

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ������Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol����))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "��" & intLop & "��ҩƷ�����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol�ɹ���))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ���Ϊ���ˣ����飡", vbInformation, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɹ���
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol�ɹ����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɹ����
                        Exit Function
                    End If
                    
                    If Split(.TextMatrix(intLop, mconIntColԭ����), "||")(0) <> "0" Then
                        If .TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntColЧ��) = "" Then
                            MsgBox "��" & intLop & "�е�ҩƷ��Ч��ҩƷ,����������ż�Ч����Ϣ�������뵥���У�", vbInformation, gstrSysName
                            mshBill.SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If .TextMatrix(intLop, mconIntCol����) = "" Then
                                .Col = mconIntCol����
                            Else
                                .Col = mconIntColЧ��
                            End If
                            Exit Function
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɹ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɹ����
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    '���۹�������Ƿ���ڲ��������۵�ҩƷ������ҩ
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                        If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                            If Val(.TextMatrix(intLop, mconIntCol�ɹ���)) <> Val(.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                                MsgBox "��" & intLop & "������ҩƷ���������۹�������ⵥ���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .MsfObj.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    End If
                    
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    With mshStructure
        For intLop = 1 To .rows - 1
            '���۹�������Ƿ���ڲ��������۵�ҩƷ��ԭ��ҩ
            If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                If IsPriceAdjustMod(Val(.TextMatrix(intLop, mconintcol��ҩƷid))) = True Then
                    If Val(.TextMatrix(intLop, mconIntCol���ɹ���)) <> Val(.TextMatrix(intLop, mconIntCol���ۼ�)) Then
                        MsgBox "��" & intLop & "��ԭ��ҩƷ���������۹������ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshStructure.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    
    ValidData = True
End Function


Private Function SaveCard() As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim lng�Ƽ��� As Long
    Dim str�������� As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim str��������1 As String
    Dim str��������2 As String
    Dim dblƫ��ɱ���� As Double
    
    Dim intRow As Integer
    Dim n As Integer
    
    On Error GoTo errHandle
    
    arrSql = Array()
    SaveCard = False
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = Sys.GetNextNo(22, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lng�Ƽ��� = cboType.ItemData(cboType.ListIndex)
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        str�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        If mint�༭״̬ = 2 Then        '�޸�
            gstrSQL = "zl_�������_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,���"
        recSort.MoveFirst
        
        str��������2 = ""
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                
                strBatchNo = .TextMatrix(intRow, mconIntCol����)
                datTimeLimit = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                    '����ΪʧЧ��������
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                
                
                dblPurchasePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɹ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                
                If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 0 Then
                    '����Ƕ���ҩƷ�����ۼ�ȡԭʼ�۸񱣴�
                    dblSalePrice = Get�ۼ�(Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1, lngDrugID, lngStockid, 0)
                                    
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(lngDrugID) = True Then
                        '�����ʵ�����۹����ҩƷ���ɱ���ҲҪ���ۼ�һ��
                        dblPurchasePrice = dblSalePrice
                    End If
                End If
                
                dblPurchaseMoney = .TextMatrix(intRow, mconIntCol�ɹ����)
                dblSaleMoney = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                dblMistakePrice = .TextMatrix(intRow, mconintCol���)
                
                str��������1 = ""
                For i = 1 To mshStructure.rows - 1
                    If lngDrugID = Val(mshStructure.TextMatrix(i, mconintColRalation)) Then
                        str��������1 = IIf(str��������1 = "", Val(mshStructure.TextMatrix(i, mconintcol��ҩƷid)) & "," & Val(mshStructure.TextMatrix(i, mconintCol����������)), str��������1 & ";" & Val(mshStructure.TextMatrix(i, mconintcol��ҩƷid)) & "," & Val(mshStructure.TextMatrix(i, mconintCol����������)))
                        str��������2 = IIf(str��������2 = "", lngDrugID & "," & Val(mshStructure.TextMatrix(i, mconintcol��ҩƷid)) & "," & Val(mshStructure.TextMatrix(i, mconintCol����������)), str��������2 & ";" & lngDrugID & "," & Val(mshStructure.TextMatrix(i, mconintcol��ҩƷid)) & "," & Val(mshStructure.TextMatrix(i, mconintCol����������)))
                    End If
                Next
                dblƫ��ɱ���� = Val(.TextMatrix(intRow, mconintColƫ��ɱ����))
                
                lngSerial = intRow
                
                gstrSQL = "zl_�������_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '���
                gstrSQL = gstrSQL & "," & lngSerial
                '�ⷿID
                gstrSQL = gstrSQL & "," & lngStockid
                '�Է�����ID
                gstrSQL = gstrSQL & "," & lng�Ƽ���
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngDrugID
                'ʵ������
                gstrSQL = gstrSQL & "," & dblQuantity
                '�ɱ���
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '�ɱ����
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '���ۼ�
                gstrSQL = gstrSQL & "," & dblSalePrice
                '���۽��
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '���
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '������
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-MM-dd") & "','yyyy-mm-dd')")
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS')"
                '��������
                gstrSQL = gstrSQL & ",'" & str��������1 & "'"
                'ƫ��ɱ����
                gstrSQL = gstrSQL & "," & dblƫ��ɱ����
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gstrSQL = "zl_ҩƷ����ԭ�ϳ���_insert('" & chrNo & "'," & lng�Ƽ��� & ",'" & str��������2 & "')"
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zldatabase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & zlStr.FormatEx(curTotal, mintMoneyDigit, , True)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & zlStr.FormatEx(Cur���ʽ��, mintMoneyDigit, , True)
    lblDifference.Caption = "��ۺϼƣ�" & zlStr.FormatEx(Cur���ʲ��, mintMoneyDigit, , True)
End Sub

Private Sub ��ʾ�����()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl���� As Double
    Dim str��λ As String
    Dim intID As Long
    Dim strQuantity As String
    
    On Error GoTo errHandle
    If mshBill.TextMatrix(mshBill.Row, mconIntColҩ��) = "" Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)
    
    If RecTmp.State = 1 Then RecTmp.Close
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strQuantity = "�������� "
        Case mconint���ﵥλ
            strQuantity = "��������/�����װ "
        Case mconintסԺ��λ
            strQuantity = "��������/סԺ��װ "
        Case mconintҩ�ⵥλ
            strQuantity = "��������/ҩ���װ "
    End Select
    
    gstrSQL = "Select b.ҩƷID, Sum(" & strQuantity & ") as ���� From ҩƷ��� a,ҩƷ��� b Where a.����=1 and a.ҩƷid=b.ҩƷid and ��������<>0 And �ⷿID=[1] and b.ҩƷID=[2] Group by b.ҩƷID "
    Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", cboStock.ItemData(cboStock.ListIndex), intID)
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    Dbl���� = IIf(IsNull(RecTmp!����), 0, RecTmp!����)
    
    staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & Dbl���� & "]" & str��λ
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ��ʾԭ�Ͽ����()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl���� As Double
    Dim str��λ As String
    Dim lngID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo errHandle
    
    If mshStructure.Row = 0 Then Exit Sub
    If mshStructure.TextMatrix(mshStructure.Row, mconIntCol��ҩ��) = "" Then
        Exit Sub
    End If
    
    lngID = mshStructure.TextMatrix(mshStructure.Row, mconintcol��ҩƷid)
    
    gstrSQL = "Select b.ҩƷID,Sum(��������) as ����,C.���㵥λ as ��λ " & _
        " From ҩƷ��� a,ҩƷ��� b,�շ���ĿĿ¼ C " & _
        " Where a.����=1 and a.ҩƷid=b.ҩƷid and B.ҩƷID=C.ID and ��������<>0 " & _
        " And �ⷿID=[1] and b.ҩƷID=[2] " & _
        " Group by b.ҩƷID,C.���㵥λ "
    Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾԭ��ҩ�����]", cboType.ItemData(cboType.ListIndex), lngID)
    
    If RecTmp.EOF Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    Dbl���� = RecTmp!����
    
    staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & Dbl���� & "]" & RecTmp!��λ
    Exit Sub
errHandle:
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
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1301", "zl8_bill_1301"), mint��¼״̬, int��λϵ��, 1301, "ҩƷ������ⵥ", strNo
End Sub

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Call zldatabase.OpenRecordset(rsBatchNolen, gstrSQL, "ȡ�ֶγ���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

