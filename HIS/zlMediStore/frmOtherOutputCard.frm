VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmOtherOutputCard 
   Caption         =   "ҩƷ�������ⵥ"
   ClientHeight    =   8295
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14475
   Icon            =   "frmOtherOutputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14475
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   32
      Top             =   5700
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   31
      Top             =   5700
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   13
      Top             =   5370
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   12
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   11
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   9
      Top             =   5280
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   10
      Top             =   5280
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5175
      Left            =   30
      ScaleHeight     =   5115
      ScaleWidth      =   14295
      TabIndex        =   14
      Top             =   0
      Width           =   14355
      Begin VB.ComboBox cbo�����λ 
         Height          =   300
         Left            =   8010
         TabIndex        =   5
         Text            =   "cbo�����λ"
         Top             =   900
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.ComboBox cbo������λ 
         Height          =   300
         Left            =   8010
         TabIndex        =   36
         Text            =   "cbo������λ"
         Top             =   900
         Visible         =   0   'False
         Width           =   3435
      End
      Begin VB.CheckBox chkIn 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   780
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "������ʵ�:F3"
         Top             =   90
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1260
         MaxLength       =   8
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   105
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   1965
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   6
         Top             =   1275
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
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
         TabIndex        =   8
         Top             =   4380
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   510
         Width           =   1965
      End
      Begin VB.Label Txt�޸����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7680
         TabIndex        =   41
         Top             =   4740
         Width           =   1875
      End
      Begin VB.Label Txt�޸��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5640
         TabIndex        =   40
         Top             =   4740
         Width           =   915
      End
      Begin VB.Label lbl�޸��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸���"
         Height          =   180
         Left            =   5040
         TabIndex        =   39
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label lbl�޸����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸�����"
         Height          =   180
         Left            =   6900
         TabIndex        =   38
         Top             =   4800
         Width           =   720
      End
      Begin VB.Label lblOther 
         AutoSize        =   -1  'True
         Caption         =   "���(��)�ϼ�:"
         Height          =   180
         Left            =   6360
         TabIndex        =   37
         Top             =   4140
         Width           =   1170
      End
      Begin VB.Label lbl�����λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����λ(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6960
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl������λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "������λ(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6960
         TabIndex        =   35
         Top             =   960
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   29
         Top             =   4140
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   28
         Top             =   4140
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   27
         Top             =   4140
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10350
         TabIndex        =   25
         Top             =   4740
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12450
         TabIndex        =   24
         Top             =   4740
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   23
         Top             =   4740
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   22
         Top             =   4740
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   7
         Top             =   4455
         Width           =   645
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ�������ⵥ"
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
         TabIndex        =   19
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ(&S)"
         Height          =   180
         Left            =   540
         TabIndex        =   0
         Top             =   570
         Width           =   630
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   18
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   17
         Top             =   4800
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   9765
         TabIndex        =   16
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   11640
         TabIndex        =   15
         Top             =   4800
         Width           =   720
      End
      Begin VB.Label lblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&T)"
         Height          =   180
         Left            =   210
         TabIndex        =   2
         Top             =   960
         Width           =   990
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
            Picture         =   "frmOtherOutputCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1000
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
            Picture         =   "frmOtherOutputCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherOutputCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   30
      Top             =   7935
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOtherOutputCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19182
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherOutputCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherOutputCard.frx":3080
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
      TabIndex        =   26
      Top             =   5400
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
Attribute VB_Name = "frmOtherOutputCard"
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
Private mblnEnterCell As Boolean            '�Ƿ�������ENTERCELL()�¼�
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mbln�¿������� As Boolean           '��Ƿ��¿�������
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���

Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrPrivs As String                 'Ȩ��
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private Const mlng��ɫ As Long = &HC000C0

Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������

Private mlng����ⷿ As Long
Private mintUnit As Integer             '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����

Private Const MStrCaption As String = "ҩƷ�����������"

Dim mstrLike As String

Private mblnLoad As Boolean              '��¼�Ƿ�ִ�����Form_Load�¼�

'�Ӳ�������ȡҩƷ�۸����������С��λ�������㾫�ȣ�
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

'=========================================================================================

Private Const mconIntCol�к� As Integer = 1
Private Const mconIntColҩ�� As Integer = 2
Private Const mconIntCol��Ʒ�� As Integer = 3
Private Const mconIntCol��Դ As Integer = 4
Private Const mconIntCol����ҩ�� As Integer = 5
Private Const mconIntCol��� As Integer = 6
Private Const mconIntCol��� As Integer = 7
Private Const mconIntCol�������� As Integer = 8
Private Const mconIntcol�ӳ��� As Integer = 9
Private Const mconIntColʵ�ʽ�� As Integer = 10
Private Const mconIntColʵ�ʲ�� As Integer = 11
Private Const mconIntCol����ϵ�� As Integer = 12
Private Const mconIntCol���� As Integer = 13
Private Const mconIntCol���� As Integer = 14
Private Const mconIntColԭ���� As Integer = 15
Private Const mconIntCol��λ As Integer = 16
Private Const mconIntCol���� As Integer = 17
Private Const mconIntColЧ�� As Integer = 18
Private Const mconIntCol��׼�ĺ� As Integer = 19
Private Const mconIntCol���� As Integer = 20
Private Const mconIntCol�������� As Integer = 21
Private Const mconIntCol�ɹ��� As Integer = 22
Private Const mconIntCol�ɹ���� As Integer = 23
Private Const mconIntCol�ۼ� As Integer = 24
Private Const mconIntCol�ۼ۽�� As Integer = 25
Private Const mconIntCol����� As Integer = 26
Private Const mconIntCol������ As Integer = 27
Private Const mconIntCol��ֵ˰�� As Integer = 28
Private Const mconIntCol˰�� As Integer = 29
Private Const mconintCol��� As Integer = 30
Private Const mconIntColҩƷ��������� = 31
Private Const mconIntColҩƷ���� = 32
Private Const mconIntColҩƷ���� = 33
Private Const mconintColԭʼ���� As Integer = 34
Private Const mconIntColS  As Integer = 35            '������
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
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(mshBill.TextMatrix(n, mconIntCol���)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol���)))
                !ҩƷID = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mconIntCol����))
                
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
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id " _
        & " FROM ҩƷ�������� A, ҩƷ������ B " _
        & "Where A.���id = B.ID " _
      & "AND A.���� = 11 "
    Call SQLTest(App.Title, "ҩƷ�������ⵥ", gstrSQL)
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "GetDepend")
    Call SQLTest
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ�������������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
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
    mstrPrivs = GetPrivFunc(glngSys, 1306)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
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
    ElseIf mint�༭״̬ = 6 Then
        mblnEdit = False
        CmdSave.Caption = "����(&O)"
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub


Private Sub cboStock_Change()
    mblnChange = True
End Sub


Private Sub cboStock_Click()
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle
    
    str�ⷿ���� = ""
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
            rsDetail.MoveNext
        Loop
        If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
        mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
        
        If mblnLoad = True Then Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
                            
                    mlng����ⷿ = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng����ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboType_click()
    Dim i As Integer
    Dim j As Integer
    Dim intresult As Integer
    
    On Error Resume Next
    
    Me.lbl�����λ.Visible = False
    Me.cbo�����λ.Visible = False
    Me.lbl������λ.Visible = False
    Me.cbo������λ.Visible = False
    
    If cboType.Text = "ҩƷ���" Then
        Me.lbl�����λ.Visible = True
        Me.cbo�����λ.Visible = True
        
        mshBill.TextMatrix(0, mconIntCol�����) = "�����"
        mshBill.TextMatrix(0, mconIntCol������) = "������"
        
        mshBill.ColWidth(mconIntCol�����) = 1000
        mshBill.ColWidth(mconIntCol������) = 1000
        cbo�����λ.Enabled = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
        mshBill.ColData(mconIntCol�����) = IIf(cbo�����λ.Enabled, 4, 5)
        
        mshBill.ColWidth(mconIntCol��ֵ˰��) = 0
        mshBill.ColWidth(mconIntCol˰��) = 0
    ElseIf cboType.Text = "ҩƷ����" Then
        If mshBill.TextMatrix(1, 0) <> "" Then
            intresult = MsgBox("������б����ݣ��Ƿ������", vbYesNo, gstrSysName)
            If intresult = vbYes Then
                Me.lbl������λ.Visible = True
                Me.cbo������λ.Visible = True
                
                mshBill.TextMatrix(0, mconIntCol�����) = "������"
                mshBill.TextMatrix(0, mconIntCol������) = "�������"
                mshBill.ColWidth(mconIntCol�����) = 1000
                mshBill.ColWidth(mconIntCol������) = 1000
                cbo������λ.Enabled = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
                mshBill.ColData(mconIntCol�����) = IIf(cbo������λ.Enabled, 4, 5)
                
                mshBill.ColWidth(mconIntCol��ֵ˰��) = 1000
                mshBill.ColWidth(mconIntCol˰��) = 1000
                
                For i = 1 To mshBill.rows - 1
                  For j = 0 To mshBill.Cols - 1
                      mshBill.TextMatrix(i, j) = ""
                  Next
                Next
                mshBill.rows = 2
                mshBill.SetFocus
            Else
                cboType.Text = "ҩƷ���"
            End If
        Else
            Me.lbl������λ.Visible = True
            Me.cbo������λ.Visible = True
            
            mshBill.TextMatrix(0, mconIntCol�����) = "������"
            mshBill.TextMatrix(0, mconIntCol������) = "�������"
            mshBill.ColWidth(mconIntCol�����) = 1000
            mshBill.ColWidth(mconIntCol������) = 1000
            cbo������λ.Enabled = (mint�༭״̬ = 1 Or mint�༭״̬ = 2)
            mshBill.ColData(mconIntCol�����) = IIf(cbo������λ.Enabled, 4, 5)
            
            mshBill.ColWidth(mconIntCol��ֵ˰��) = 1000
            mshBill.ColWidth(mconIntCol˰��) = 1000
        End If
    Else
        mshBill.ColWidth(mconIntCol�����) = 0
        mshBill.ColWidth(mconIntCol������) = 0
        mshBill.ColData(mconIntCol�����) = 5
        
        mshBill.ColWidth(mconIntCol��ֵ˰��) = 0
        mshBill.ColWidth(mconIntCol˰��) = 0
    End If
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'���ܣ���ItemData��Text����ComboBox������ֵ
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If Rec.IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '�Ⱦ�ȷ����
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '��ģ������
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Private Sub cbo�����λ_GotFocus()
    If cbo�����λ.Style = 0 Then
        Call zlControl.TxtSelAll(cbo�����λ)
    End If
End Sub

Private Sub cbo�����λ_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    Call zlCommFun.PressKey(vbKeyTab)
    If KeyCode = vbKeyDelete Then
        If cbo�����λ.Style = 2 And cbo�����λ.ListIndex <> -1 Then
            cbo�����λ.ListIndex = -1
        End If
    End If
End Sub

Private Sub cbo�����λ_KeyPress(KeyAscii As Integer)
'    Dim IntMatchIdx As Integer
'
'    With cbo�����λ
'        IntMatchIdx = MatchIndex(.hWnd, KeyAscii, 1)
'        If IntMatchIdx = -2 Then Exit Sub
'        .ListIndex = IntMatchIdx
'        If .ListIndex = -1 Then .ListIndex = 0
'    End With

    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cbo�����λ.Locked And cbo�����λ.Style = 2 Then
            lngIdx = Cbo.MatchIndex(cbo�����λ.hWnd, KeyAscii)
            
            If lngIdx = -1 And cbo�����λ.ListCount > 0 Then lngIdx = 0
            cbo�����λ.ListIndex = lngIdx
        End If
    End If
End Sub

Private Sub cbo�����λ_Validate(Cancel As Boolean)
    '���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo�����λ.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cbo�����λ.Text = "" Then cbo�����λ.Tag = "": Exit Sub '������
    
    strInput = UCase(NeedName(cbo�����λ.Text))
    strSQL = "Select Rownum As id,����,����,���� From ҩƷ�����λ Where Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2] Order By ����"
        
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cbo�����λ.hWnd)
    Set rsTmp = zlDataBase.ShowSQLSelect(Me, strSQL, 0, "�����λ", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo�����λ.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cbo�����λ, nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����)
        If intIdx <> -1 Then
            cbo�����λ.ListIndex = intIdx
        Else
            cbo�����λ.AddItem nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����, cbo�����λ.ListCount - 1
            cbo�����λ.ListIndex = cbo�����λ.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�������λ��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cbo������λ_GotFocus()
    If cbo������λ.Style = 0 Then
        Call zlControl.TxtSelAll(cbo������λ)
    End If
End Sub

Private Sub cbo������λ_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    Call zlCommFun.PressKey(vbKeyTab)
    If KeyCode = vbKeyDelete Then
        If cbo������λ.Style = 2 And cbo������λ.ListIndex <> -1 Then
            cbo������λ.ListIndex = -1
        End If
    End If
End Sub


Private Sub cbo������λ_KeyPress(KeyAscii As Integer)
'    Dim IntMatchIdx As Integer
'
'    With cbo������λ
'        IntMatchIdx = MatchIndex(.hWnd, KeyAscii, 1)
'        If IntMatchIdx = -2 Then Exit Sub
'        .ListIndex = IntMatchIdx
'        If .ListIndex = -1 Then .ListIndex = 0
'    End With
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyAscii >= 32 Then
        If Not cbo������λ.Locked And cbo������λ.Style = 2 Then
            lngIdx = Cbo.MatchIndex(cbo������λ.hWnd, KeyAscii)
            If lngIdx = -1 And cbo������λ.ListCount > 0 Then lngIdx = 0
            cbo������λ.ListIndex = lngIdx
        End If
    End If
End Sub


Private Sub cbo������λ_Validate(Cancel As Boolean)
    '���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If cbo������λ.ListIndex <> -1 Then Exit Sub '��ѡ��
    If cbo������λ.Text = "" Then cbo������λ.Tag = "": Exit Sub '������
    
    strInput = UCase(NeedName(cbo������λ.Text))
    strSQL = "Select Rownum As id,����,����,���� From ҩƷ������λ Where Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2] Order By ����"
        
    On Error GoTo errH
    vRect = zlControl.GetControlRect(cbo������λ.hWnd)
    Set rsTmp = zlDataBase.ShowSQLSelect(Me, strSQL, 0, "������λ", False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo������λ.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = SeekCboIndex(cbo������λ, nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����)
        If intIdx <> -1 Then
            cbo������λ.ListIndex = intIdx
        Else
            cbo������λ.AddItem nvl(rsTmp!����) & "-" & Chr(13) & rsTmp!����, cbo������λ.ListCount - 1
            cbo������λ.ListIndex = cbo������λ.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ��������λ��", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkIn_Click()
    txtIn.Enabled = chkIn.Value
    If chkIn.Value Then
        txtIn.SetFocus
    Else
        txtIn.Text = ""
    End If
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(0, mintMoneyDigit, , True)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
    mblnChange = False
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol��������) = .TextMatrix(intRow, mconIntCol����)
                .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ɹ���), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ۼ�), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ۽��) - .TextMatrix(intRow, mconIntCol�ɹ����), mintMoneyDigit, , True)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
    
    mblnChange = True
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
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 6 Then
                MsgBox "�õ�����û�п��Գ�����ҩƷ�����飡", vbOKOnly, gstrSysName
            Else
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
            End If
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
        gint���뷽ʽ = Val(zlDataBase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
    
    'Call cboType_Click
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
    Dim lngҩƷID As Long
    Dim intRow As Integer
    Dim bln�ⷿ As Boolean
    Dim bln���� As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim BlnSuccess As Boolean
    Dim blnTrans As Boolean
    Dim intLop As Integer
    Dim lng�ϴ�ҩƷID As Long
    
    On Error GoTo ErrHand
    
    '�����������ݼ�
    Call SetSortRecord
        
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    '��������ҩƷ����Ԥ���۴���
    For intLop = 1 To Me.mshBill.rows - 1
        If mshBill.TextMatrix(intLop, 0) <> "" Then '��ҩƷ
            Call AutoAdjustPrice_ByID(Val(mshBill.TextMatrix(intLop, 0)))
        End If
    Next
    
    If mint�༭״̬ = 3 Then        '���
        mstrTime_End = GetBillInfo(11, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not ��鵥��(11, txtNo, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    
        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub

        '���۹�������Ƿ���ڲ��������۵�ҩƷ
        For intLop = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(intLop, 0) <> "" And gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                If IsPriceAdjustMod(Val(mshBill.TextMatrix(intLop, 0))) = True Then
                    If CheckPriceAdjust(Val(mshBill.TextMatrix(intLop, 0)), cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(intLop, mconIntCol����))) = False Then
                        MsgBox "��" & intLop & "��ҩƷ���������۹���������¼���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        mshBill.Row = intLop
                        mshBill.MsfObj.TopRow = intLop
                        Exit Sub
                    End If
                End If
            End If
        Next
        
        blnTrans = True
        gcnOracle.BeginTrans
        '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
        End If
        
        If Not SaveCheck Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
                
        gcnOracle.CommitTrans
        
        If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.��������)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.��������)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1306_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                            lng�ϴ�ҩƷID = recSort!ҩƷID
                        End If
                        recSort.MoveNext
                    Loop
                End If

            End If
        End If

        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then '����
        If mblnChange = False Then
            MsgBox "��¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("��ȷʵҪ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 2 Then
        If Not ��鵥��(11, txtNo, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If mint�༭״̬ = 1 Then '��������ʱ���жϼ۸��Ƿ��Ѿ�����
        If ���۸� Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If ValidData = False Then Exit Sub
    
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.��������)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.��������)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1306_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                            lng�ϴ�ҩƷID = recSort!ҩƷID
                        End If
                        recSort.MoveNext
                    Loop
                End If

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
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    
    txtժҪ.Text = ""
    cboType.SetFocus
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lngҩƷID As Long
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsPrice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    Dim intCostDigit As Integer
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
        
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 11 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(b.�ּ�, " & intPriceDigit & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0 " & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 1 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = 11 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�), " & intPriceDigit & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
                  " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 2 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = 11 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & intCostDigit & ")<>round(decode(x.�ּ�,b.ƽ���ɱ���,x.�ּ�)," & intCostDigit & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Order By ����, ҩƷid, ���"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        Dbl���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
        dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ɹ���))
        dbl���ۼ� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ�))
        dbl�ɱ���� = dbl�ɱ��� * Dbl����
        dbl���۽�� = dbl���ۼ� * Dbl����
        dbl��� = dbl���۽�� - dbl�ɱ����
                
        If lngҩƷID <> 0 Then
            rsPrice.Filter = "����='�ۼ�' And ҩƷID=" & lngҩƷID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl���ۼ� = Val(FormatEx(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intPriceDigit))
                dbl���۽�� = Val(zlStr.FormatEx(Val(FormatEx(dbl���ۼ�, intPriceDigit)) * Dbl����, mintMoneyDigit, , True))
                dbl��� = Val(zlStr.FormatEx(dbl���۽�� - dbl�ɱ����, mintMoneyDigit, , True))
            End If
            
            rsPrice.Filter = "����='�ɱ���' And ҩƷID=" & lngҩƷID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(zlStr.FormatEx(Val(FormatEx(dbl���ۼ�, intPriceDigit)) * Dbl����, mintMoneyDigit, , True))
                dbl�ɱ��� = Val(FormatEx(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intCostDigit))
                dbl�ɱ���� = Val(zlStr.FormatEx(dbl�ɱ��� * Dbl����, mintMoneyDigit, , True))
                dbl��� = Val(zlStr.FormatEx(dbl���۽�� - dbl�ɱ����, mintMoneyDigit, , True))
            End If
            
            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, intPriceDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(dbl���۽��, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ���) = zlStr.FormatEx(dbl�ɱ���, intCostDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ����) = zlStr.FormatEx(dbl�ɱ����, mintMoneyDigit, , True)
                mshBill.TextMatrix(lngRow, mconintCol���) = zlStr.FormatEx(dbl���, mintMoneyDigit, , True)
            End If
        End If
    Next
    rsPrice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double, ByVal dbl����ϵ�� As Integer) As Boolean
    '���ܣ��ʱ�����ʵ�������Ƿ��㹻������>0˵���ǰ������γ��⣬����=0˵����������⣬���ַ�ʽ����Ҫ�����
    '����ֵ��true-����㹻��false-��治�㹻
    Dim rsData As ADODB.Recordset
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim lng�ⷿID As Long
    Dim dblʵ������ As Double
    
    With mshBill
        lngҩƷID = Val(.TextMatrix(intRow, 0))
        lng���� = Val(.TextMatrix(intRow, mconIntCol����))
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        
        If lng���� > 0 Then
            gstrSQL = "Select (a.ʵ������ - [1]) As ʣ������,a.ʵ������" & vbNewLine & _
                        "From ҩƷ��� a" & vbNewLine & _
                        "Where a.ҩƷid = [2] And a.�ⷿid = [3] And Nvl(a.����, 0) = [4] and a.���� = 1"
        Else
            gstrSQL = "Select Sum(a.ʵ������) - [1] As ʣ������, Sum(a.ʵ������) As ʵ������" & vbNewLine & _
                        "From ҩƷ��� A" & vbNewLine & _
                        "Where a.ҩƷid = [2] And a.�ⷿid = [3] And a.���� = 1"
        End If
        Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, "�����", dbl��д���� * dbl����ϵ��, lngҩƷID, lng�ⷿID, lng����)
        
        If lng���� > 0 Then
            If rsData.RecordCount > 0 Then
                dblʵ������ = zlStr.FormatEx(nvl(rsData!ʵ������, 0) / dbl����ϵ��, mintNumberDigit, , True)
                
                If rsData!ʣ������ >= 0 Then
                    CheckQuantity = True
                Else
                    CheckQuantity = False
                End If
            Else
                CheckQuantity = False
            End If
        Else
            If rsData.RecordCount > 0 Then
                dblʵ������ = zlStr.FormatEx(nvl(rsData!ʵ������, 0) / dbl����ϵ��, mintNumberDigit, , True)
                
                If IsNull(rsData!ʣ������) Then
                    CheckQuantity = False
                Else
                    If rsData!ʣ������ >= 0 Then
                        CheckQuantity = True
                    Else
                        CheckQuantity = False
                    End If
                End If
            Else
                CheckQuantity = False
            End If
        End If
        
        If CheckQuantity = False Then
            If mint����� = 0 Then
                '0-���㲻���
                CheckQuantity = True
            ElseIf mint����� = 1 Then
                '1-��飬��������
                If MsgBox("����������������˿��ʵ������(" & dblʵ������ & ")���Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    CheckQuantity = True
                End If
            ElseIf mint����� = 2 Then
                '2-��飬�����ֹ
                MsgBox "����������������˿��ʵ������(" & dblʵ������ & ")", vbInformation, gstrSysName
            End If
        End If
    End With
End Function


Private Sub Form_Load()
    Dim rsTemp As New Recordset
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle
    mblnLoad = False
    mblnEnterCell = False
    mstrLike = IIf(Val(zlDataBase.GetPara("����ƥ��")) = 0, "%", "")
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    mblnUpdate = False
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�����������", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GetSysParm
    
    With cboType
        .Clear
        gstrSQL = "SELECT b.Id,b.���� " _
            & " FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID " _
              & "AND A.���� = 11 "
        Call SQLTest(App.Title, "ҩƷ�������ⵥ", gstrSQL)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "Form_Load")
        Call SQLTest
        
        Do While Not rsTemp.EOF
            .AddItem rsTemp.Fields(1)
            .ItemData(.NewIndex) = rsTemp.Fields(0)
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    With cbo�����λ
        .Clear
        gstrSQL = "Select Rownum As Id, ����, ����, ���� From ҩƷ�����λ Order By ����"
        Call zlDataBase.OpenRecordset(rsTemp, gstrSQL, "��ȡ�����λ")
        
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!���� & "-" & rsTemp!����
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    With cbo������λ
        .Clear
        gstrSQL = "Select Rownum As Id, ����, ����, ���� From ҩƷ������λ Order By ����"
        Call zlDataBase.OpenRecordset(rsTemp, gstrSQL, "��ȡ������λ")
        
        .AddItem ""
        Do While Not rsTemp.EOF
            .AddItem rsTemp!���� & "-" & rsTemp!����
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        .ListIndex = 0
    End With
    
    mlng����ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng����ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(11, mstr���ݺ�)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    'ֻ����ҩ��ⷿ����ʾ"ԭ����"��
    str�ⷿ���� = ""
    gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
        rsDetail.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
    mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
    
    mshBill.ColWidth(mconIntCol��������) = IIf(mint�༭״̬ = 6, 1100, 0)
    
    '������ԱȨ�޾����Ƿ���ʾ�ɱ���
    mshBill.ColWidth(mconIntCol�ɹ���) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol�ɹ����) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol���) = IIf(mblnViewCost, 900, 0)
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
    
    mblnEnterCell = True
    
    Call cboType_click
    mblnChange = False
    mblnLoad = True
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
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPriceDigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    On Error GoTo errHandle
    '�ⷿ
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.��������)
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
    intPriceDigit = mintPriceDigit
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
'            Txt�޸��� = UserInfo.�û�����
'            Txt�޸����� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        
        Case 2, 3, 4, 6
            Call initGrid
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "select distinct b.id,b.���� from ҩƷ�շ���¼ a,���ű� b  " _
                    & " where a.�ⷿid=b.id and A.���� =11 and  a.no=[1]"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
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
                    strUnitQuantity = "F.���㵥λ AS ��λ, A.��д���� as ����,a.�ɱ���,a.���ۼ�,nvl(a.����,0) As �����,'1' as ����ϵ��,"
                Case mconint���ﵥλ
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ����,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,nvl(a.����,0)*B.�����װ As �����,B.�����װ as ����ϵ��,"
                Case mconintסԺ��λ
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ����,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,nvl(a.����,0)*B.סԺ��װ As �����,B.סԺ��װ as ����ϵ��,"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ����,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,nvl(a.����,0)*B.ҩ���װ As �����,B.ҩ���װ as ����ϵ��,"
            End Select
            
            If mint�༭״̬ <> 6 Then
                gstrSQL = "SELECT W.*,Z.��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                    " FROM " & _
                    " (SELECT DISTINCT A.ҩƷID,A.���,'[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��, " & _
                    " B.ҩƷ��Դ,B.����ҩ��,F.���,F.���� AS ԭ������,A.����, A.ԭ����,A.����,A.����,B.�ӳ���,A.Ч��," & _
                    strUnitQuantity & _
                    " A.�ɱ����,A.���۽��, A.���,A.ժҪ,������,��������,�޸���,�޸�����,�����,�������,A.�ⷿID,A.������ID,F.�Ƿ���,B.ҩ������ AS ҩ����������," & _
                    " G.���� AS �����λ,A.��׼�ĺ�,H.���� AS ������λ, To_Number(Trim(To_Char(Nvl(A.Ƶ��, '0'), '999999999999.0000'))) As ��ֵ˰�� " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ F,ҩƷ�����λ G,ҩƷ������λ H " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID And A.��ҩ����=G.����(+) And A.��ҩ����=H.����(+) " & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND E.����(+)=1 " & _
                    " AND A.��¼״̬ =[3] " & _
                    " AND A.���� = 11 AND A.NO = [1]) W," & _
                    " (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    " FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1)  Z " & _
                    " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                    " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT W.*,Z.��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                    " FROM " & _
                    " (SELECT DISTINCT A.ҩƷID,A.���,'[' || F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��, " & _
                    " B.ҩƷ��Դ,B.����ҩ��,F.���,F.���� AS ԭ������,A.����, A.ԭ����,A.����,A.����,B.�ӳ���,A.Ч��,G.���� AS �����λ,H.���� AS ������λ,A.��ֵ˰��," & _
                    strUnitQuantity & _
                    " A.�ɱ����,0 ���۽��, 0 ���,A.ժҪ,A.�ⷿID,A.������ID,F.�Ƿ���,B.ҩ������ AS ҩ����������,A.��׼�ĺ�,A.��д���� As ԭʼ���� " & _
                    " FROM " & _
                    "     (SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,SUM(�ɱ����) AS �ɱ����,ҩƷID,���,����, ԭ����,����,Ч��,NVL(����,0) ����," & _
                    " ����,�ɱ���,���ۼ�,ժҪ,�ⷿID,������ID,����,��ҩ����,��׼�ĺ�, To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000'))) As ��ֵ˰��" & _
                    "     FROM ҩƷ�շ���¼ X " & _
                    "     WHERE NO=[1] AND ����=11  " & _
                    "     GROUP BY ҩƷID,���,����,ԭ����,����,Ч��,NVL(����,0),����,�ɱ���,���ۼ�,ժҪ,�ⷿID,������ID,����,��ҩ����,��׼�ĺ�, To_Number(Trim(To_Char(Nvl(Ƶ��, '0'), '999999999999.0000'))) " & _
                    "     HAVING SUM(ʵ������)<>0 ) A," & _
                    "     ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ F,ҩƷ�����λ G,ҩƷ������λ H " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID And A.��ҩ����=G.����(+) And A.��ҩ����=H.����(+) " & _
                    "     AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND E.����(+)=1) W," & _
                    "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "     FROM ҩƷ��� WHERE �ⷿID=[2]  AND ����=1)  Z " & _
                    " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ҩƷ�������ⵥ]", mstr���ݺ�, cboStock.ItemData(cboStock.ListIndex), mint��¼״̬)
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint�༭״̬
            Case 2, 6
                If mint�༭״̬ = 2 Then
                    Txt������ = rsInitCard!������
                    Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                    Txt�޸��� = IIf(IsNull(rsInitCard!�޸���), "", rsInitCard!�޸���)
                    Txt�޸����� = IIf(IsNull(rsInitCard!�޸�����), "", Format(rsInitCard!�޸�����, "yyyy-mm-dd hh:mm:ss"))
                End If
                If mint�༭״̬ = 6 Then
                    Txt������ = UserInfo.�û�����
                    Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'                    Txt�޸��� = UserInfo.�û�����
'                    Txt�޸����� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                    Txt����� = UserInfo.�û�����
                    Txt������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
            Case Else
                Txt������ = rsInitCard!������
                Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                Txt�޸��� = IIf(IsNull(rsInitCard!�޸���), "", rsInitCard!�޸���)
                Txt�޸����� = IIf(IsNull(rsInitCard!�޸�����), "", Format(rsInitCard!�޸�����, "yyyy-mm-dd hh:mm:ss"))
                Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            End Select
            
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboType
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsInitCard!������ID Then
                        .ListIndex = intCount
                        Exit For
                    End If
                Next
                
                If .Text = "ҩƷ���" Then
                    Me.cbo�����λ.Visible = True
                    
                    '��λ�����λ
                    If Not IsNull(rsInitCard!�����λ) Then
                        For i = 1 To cbo�����λ.ListCount - 1
                            If Mid(cbo�����λ.List(i), InStr(1, cbo�����λ.List(i), "-") + 1) = rsInitCard!�����λ Then
                                cbo�����λ.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                End If

                If .Text = "ҩƷ����" Then
                    Me.cbo������λ.Visible = True
                    
                    '��λ������λ
                    If Not IsNull(rsInitCard!������λ) Then
                        For i = 1 To cbo������λ.ListCount - 1
                            If Mid(cbo������λ.List(i), InStr(1, cbo������λ.List(i), "-") + 1) = rsInitCard!������λ Then
                                cbo������λ.ListIndex = i
                                Exit For
                            End If
                        Next
                    End If
                End If
            End With
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
                Do While Not rsInitCard.EOF
                    
                    intRow = intRow + 1
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
                    
                    .TextMatrix(intRow, mconIntCol��Դ) = nvl(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = nvl(rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntCol���) = rsInitCard!���
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!ԭ����), "", rsInitCard!ԭ����)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(rsInitCard!����, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(rsInitCard!�ɱ���, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(IIf(mint�༭״̬ = 6, 0, rsInitCard!�ɱ����), intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx(rsInitCard!�����, intCostDigit, , True)
                    .TextMatrix(intRow, mconIntCol������) = zlStr.FormatEx(rsInitCard!����� * rsInitCard!����, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntcol�ӳ���) = rsInitCard!�ӳ��� / 100 & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������), intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntColʵ�ʲ��) = IIf(IsNull(rsInitCard!ʵ�ʲ��), "0", rsInitCard!ʵ�ʲ��)
                    .TextMatrix(intRow, mconIntColʵ�ʽ��) = IIf(IsNull(rsInitCard!ʵ�ʽ��), "0", rsInitCard!ʵ�ʽ��)
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mconIntCol��ֵ˰��) = zlStr.FormatEx(IIf(IsNull(rsInitCard!��ֵ˰��), "0", rsInitCard!��ֵ˰��), 2, , True)
                    .TextMatrix(intRow, mconIntCol˰��) = zlStr.FormatEx(rsInitCard!����� * rsInitCard!���� * (Val(.TextMatrix(intRow, mconIntCol��ֵ˰��)) / 100 / (1 + Val(.TextMatrix(intRow, mconIntCol��ֵ˰��)) / 100)), intMoneyDigit, , True)
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(0, intNumberDigit, , True)
                        .TextMatrix(intRow, mconIntCol������) = zlStr.FormatEx(0, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconIntCol˰��) = zlStr.FormatEx(0, intMoneyDigit, , True)
                        .TextMatrix(intRow, mconintColԭʼ����) = Val(rsInitCard!ԭʼ����)
                    End If
                    
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!ҩƷID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        mcolUsedCount.Add Array(CStr(rsInitCard!ҩƷID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)), CStr(numUseAbleCount + IIf(IsNull(rsInitCard!����), "0", rsInitCard!����))), CStr(rsInitCard!ҩƷID) & CStr(IIf(IsNull(rsInitCard!����), "0", rsInitCard!����))
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                .rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "������"
        .TextMatrix(0, mconIntColԭ����) = "ԭ����"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconIntCol�����) = "�����"
        .TextMatrix(0, mconIntCol������) = "������"
        .TextMatrix(0, mconIntCol��ֵ˰��) = "��ֵ˰��%"
        .TextMatrix(0, mconIntCol˰��) = "˰��"
        .TextMatrix(0, mconintCol���) = "���"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconintColԭʼ����) = "ԭʼ����"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        .ColWidth(mconIntCol����) = 1100
        .ColWidth(mconIntCol��������) = IIf(mint�༭״̬ = 6, 1100, 0)
        .ColWidth(mconIntCol�ɹ���) = 1000
        .ColWidth(mconIntCol�ɹ����) = 1000
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconIntCol�ۼ۽��) = 1000
        .ColWidth(mconintCol���) = 1000
        .ColWidth(mconIntCol�����) = 0
        .ColWidth(mconIntCol������) = 0
        .ColWidth(mconIntCol��ֵ˰��) = 0
        .ColWidth(mconIntCol˰��) = 0
        
        .ColWidth(mconIntCol��������) = 0
        
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        .ColWidth(mconIntcol�ӳ���) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconintColԭʼ����) = 0
        
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntCol��Դ) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColԭ����) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol��׼�ĺ�) = 5
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntCol��ֵ˰��) = 5
        
        chkIn.Visible = (mint�༭״̬ = 1)
        txtIn.Visible = (mint�༭״̬ = 1)
        
        cbo�����λ.Enabled = False
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            cboType.Enabled = True
            txtժҪ.Enabled = True
            
            cboStock.Enabled = True
            
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol����) = 4
            .ColData(mconIntCol�����) = IIf(Me.cbo�����λ.Visible Or Me.cbo�����λ.Visible, 4, 5)
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 4 Or mint�༭״̬ = 6 Then
            cboStock.Enabled = False
            cboType.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol����) = 5
            .ColData(mconIntCol�����) = 5
        End If
        .ColData(mconIntCol��������) = 4
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
        .ColData(mconIntCol������) = 5
        .ColData(mconIntCol˰��) = 5
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntColʵ�ʲ��) = 5
        .ColData(mconIntColʵ�ʽ��) = 5
        .ColData(mconIntcol�ӳ���) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntCol����) = 5
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColԭ����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
        If InStr(1, "346", mint�༭״̬) <> 0 Then .ColData(mconIntColҩ��) = 0
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
    
    cbo�����λ.Left = mshBill.Left + mshBill.Width - cbo�����λ.Width
    lbl�����λ.Left = cbo�����λ.Left - lbl�����λ.Width - 100
    
    lbl������λ.Left = lbl�����λ.Left
    lbl������λ.Top = lbl�����λ.Top
    cbo������λ.Left = cbo�����λ.Left
    cbo������λ.Top = cbo�����λ.Top
    
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
    
    With lbl�޸���
        .Top = Lbl������.Top
        .Left = Pic����.Width / 2 - (450 + Txt�޸���.Width + lbl�޸���.Width + Txt�޸�����.Width + lbl�޸�����.Width) / 2
    End With
    
    With Txt�޸���
        .Top = Lbl������.Top - 80
        .Left = lbl�޸���.Left + lbl�޸���.Width + 100
    End With
    
    With lbl�޸�����
        .Top = Lbl������.Top
        .Left = Txt�޸���.Left + Txt�޸���.Width + 250
    End With
    
    With Txt�޸�����
        .Top = Lbl������.Top - 80
        .Left = lbl�޸�����.Left + lbl�޸�����.Width + 100
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
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
        lblOther.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    With lblOther
        .Left = lblPurchasePrice.Left + mshBill.Width / 4 * 3
    End With
    
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
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
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
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�����������", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
End Sub

Private Function SaveCheck() As Boolean
    Dim intRow As Integer
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim str����� As String
    Dim dat������� As String
    
    Dim int��� As Integer
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim num���� As Double
    Dim num�ɱ��� As Double
    Dim num�ɱ���� As Double
    Dim num���۽�� As Double
    Dim num��� As Double
    Dim lng������id As Long
    Dim strҩƷ As String
    Dim intNumCol As Integer
    
    Dim arrSql As Variant
    Dim n As Integer
    
    arrSql = Array()
    
    mblnSave = False
    SaveCheck = False
    
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng������id = cboType.ItemData(cboType.ListIndex)
    str����� = UserInfo.�û�����
    strNo = txtNo.Tag
    
    dat������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    
    If mint�༭״̬ = 6 Then
        intNumCol = mconIntCol��������
    Else
        intNumCol = mconIntCol����
    End If
    '�����
    strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, intNumCol, mconIntCol����ϵ��, 1, 1, mintNumberDigit)
    If strҩƷ <> "" Then
        If mint����� = 1 Then '��������
            If MsgBox("ҩƷ��" & strҩƷ & "����治�㣬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        ElseIf mint����� = 2 Then '�����ֹ
            MsgBox "ҩƷ��" & strҩƷ & "����治�㣬������ˣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With mshBill
        On Error GoTo errHandle
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                If Val(.TextMatrix(intRow, mconIntCol����)) = 0 Then
                    .TextMatrix(intRow, mconIntCol�ɹ���) = 0
                Else
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx((.TextMatrix(intRow, mconIntCol�ۼ۽��) - .TextMatrix(intRow, mconintCol���)) / (.TextMatrix(intRow, mconIntCol����)), gtype_UserDrugDigits.Digit_�ɱ���, , True)
                End If
                .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɹ���) * (.TextMatrix(intRow, mconIntCol����)), mintMoneyDigit, , True)
                
                lngҩƷID = .TextMatrix(intRow, 0)
                lng���� = .TextMatrix(intRow, mconIntCol����)
                num���� = .TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol����ϵ��)
                
'                num�ɱ��� = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɹ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                num�ɱ��� = Get�ɱ���(lngҩƷID, lng�ⷿID, lng����)
                
                num�ɱ���� = .TextMatrix(intRow, mconIntCol�ɹ����)
                num���۽�� = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                num��� = .TextMatrix(intRow, mconintCol���)
                int��� = Val(.TextMatrix(intRow, mconIntCol���))

                gstrSQL = "zl_ҩƷ��������_Verify("
                '���
                gstrSQL = gstrSQL & int���
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '�ⷿID
                gstrSQL = gstrSQL & "," & lng�ⷿID
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngҩƷID
                '����
                gstrSQL = gstrSQL & "," & lng����
                'ʵ������
                gstrSQL = gstrSQL & "," & num����
                '�ɱ���
                gstrSQL = gstrSQL & "," & num�ɱ���
                '�ɱ����
                gstrSQL = gstrSQL & "," & num�ɱ����
                '���۽��
                gstrSQL = gstrSQL & "," & num���۽��
                '���
                gstrSQL = gstrSQL & "," & num���
                '�����
                gstrSQL = gstrSQL & ",'" & str����� & "'"
                '�������
                gstrSQL = gstrSQL & ",to_date('" & dat������� & "','yyyy-mm-dd HH24:MI:SS')"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngҩƷID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSql, MStrCaption, False, False) Then Exit Function
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    '���ʳ��� Write by zyb, ##20021016##
    Dim �д�_IN As Integer
    Dim ԭ��¼״̬_IN As Integer
    Dim NO_IN As String
    Dim ���_IN As Integer
    Dim ҩƷID_IN As Long
    Dim ��������_IN As Double
    Dim ������_IN As String
    Dim ��������_IN  As String
    Dim intRow As Integer
    Dim n As Integer
    Dim strҩƷID As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim strҩƷ As String
    Dim intNumCol As Integer
    
    SaveStrike = False
    arrSql = Array()
    With mshBill
        '����������������С����
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntCol��������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mconIntCol����)), Val(.TextMatrix(intRow, mconIntCol��������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Next
        
        If mint�༭״̬ = 6 Then
            intNumCol = mconIntCol��������
        Else
            intNumCol = mconIntCol����
        End If
        '�����
        strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, intNumCol, mconIntCol����ϵ��, 2, 1, mintNumberDigit)
        If strҩƷ <> "" Then
            If mint����� = 1 Then '��������
                If MsgBox("ҩƷ��" & strҩƷ & "����治�㣬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            ElseIf mint����� = 2 Then '�����ֹ
                MsgBox "ҩƷ��" & strҩƷ & "����治�㣬������ˣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        NO_IN = Trim(txtNo.Tag)
        ������_IN = UserInfo.�û�����
        ��������_IN = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ԭ��¼״̬_IN = mint��¼״̬
        
        On Error GoTo errHandle
        
        �д�_IN = 0

        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntCol��������)) <> 0 Then
                �д�_IN = �д�_IN + 1
                
                ҩƷID_IN = .TextMatrix(intRow, 0)
                strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & ҩƷID_IN
                
                If Val(.TextMatrix(intRow, mconIntCol��������)) = Val(.TextMatrix(intRow, mconIntCol����)) Then
                    '�����ȫ�壬������������ԭʼ���������ⵥλ������ֵ����
                    ��������_IN = Val(.TextMatrix(intRow, mconintColԭʼ����))
                Else
                    ��������_IN = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol��������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                End If
                
                ���_IN = .TextMatrix(intRow, mconIntCol���)
                
                gstrSQL = "ZL_ҩƷ��������_STRIKE("
                '�д�
                gstrSQL = gstrSQL & �д�_IN
                'ԭ��¼״̬
                gstrSQL = gstrSQL & "," & ԭ��¼״̬_IN
                'NO
                gstrSQL = gstrSQL & ",'" & NO_IN & "'"
                '���
                gstrSQL = gstrSQL & "," & ���_IN
                'ҩƷID
                gstrSQL = gstrSQL & "," & ҩƷID_IN
                '��������
                gstrSQL = gstrSQL & "," & ��������_IN
                '������
                gstrSQL = gstrSQL & ",'" & ������_IN & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & Format(��������_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        If �д�_IN = 0 Then
            MsgBox "û��ѡ��һ��ҩƷ����������¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '��ʾͣ��ҩƷ
        If strҩƷID <> "" Then
            Call CheckStopMedi(strҩƷID)
        End If
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
    End With
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mconIntColҩ��) = 0 Then
        'Cancel = True    '�ȴ���CANCEL����
        Exit Sub
    End If
        
        
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
    mblnChange = True
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex))
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), , , , , , , , , mstrPrivs)
    If RecReturn.RecordCount > 0 Then
        Set RecReturn = CheckData(RecReturn)
    End If
      
    mshBill.CmdEnable = True
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
        For i = 1 To RecReturn.RecordCount
            intCurRow = mshBill.Row
            With mshBill
                .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                SetColValue .Row, RecReturn!ҩƷID, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                    nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, _
                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                    IIf(IsNull(RecReturn!�ӳ���), "0", RecReturn!�ӳ��� / 100), _
                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                    IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                    IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�), nvl(RecReturn!ԭ����)
                .Col = mconIntCol����
                
                If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                    .rows = .rows + 1
                End If
                .Row = .rows - 1
                RecReturn.MoveNext
            End With
        Next
        mshBill.Row = intOldRow
        RecReturn.Close
    End If
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
        
        If .Col = mconIntCol���� Or .Col = mconIntCol�������� Or .Col = mconIntCol�ɹ��� Or .Col = mconIntCol����� Or .Col = mconIntCol�ۼ� Or .Col = mconIntCol�ɹ���� Or .Col = mconIntCol������ Then
            Select Case .Col
                Case mconIntCol����, mconIntCol��������
                    intDigit = mintNumberDigit
                Case mconIntCol�ɹ���, mconIntCol�����
                   intDigit = mintCostDigit
                Case mconIntCol�ۼ�
                    intDigit = mintPriceDigit
                Case mconIntCol�ɹ����, mconIntCol������
                    intDigit = mintMoneyDigit
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
    If Not mblnEnterCell Then Exit Sub
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntColҩ��
                .txtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
            Case mconIntCol����
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                Call ��ʾ�����
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsTemp As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        .Text = UCase(Trim(.Text))
        strKey = UCase(Trim(.Text))
        
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
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), , , strkey, sngLeft, sngTop)
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
                    End If
                    Set RecReturn = frmSelector.ShowME(Me, 1, 2, strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , , , , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                            If SetColValue(.Row, RecReturn!ҩƷID, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                    nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, _
                                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                    IIf(IsNull(RecReturn!�ӳ���), "0", RecReturn!�ӳ���), _
                                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                                    IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                                    IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�), nvl(RecReturn!ԭ����)) = False Then
                                Cancel = True
                                Exit Sub
                            End If
                            .Text = .TextMatrix(.Row, .Col)
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        If Val(.TextMatrix(.Row, 0)) = 0 Then
                            .Text = .TextMatrix(.Row, .Col)
                            Cancel = True
                        Else
                            .Text = .TextMatrix(.Row, .Col)
                        End If
                    End If
                    Call ��ʾ�����
                End If
            
            Case mconIntCol����
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
                
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
                    
                    '�����
                    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                        If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), strKey, Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��)), txtNo.Caption, 11, mint�����, mintNumberDigit) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                    End If
                    
                    If strKey <> 0 Then
                        .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(Get�ɱ���(Val(.TextMatrix(.Row, 0)), Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, mconIntCol����))) * Val(Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��))), mintCostDigit, , True)
                    End If
                    .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * strKey, mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Val(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��))) - Val(.TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                    
                    '��������ۼ�������:�����=(1+����ѱ���)*����
                    If Val(.TextMatrix(.Row, 0)) <> 0 And cboType.Text = "ҩƷ���" And .TextMatrix(.Row, mconIntCol�����) = "" Then
                        gstrSQL = "Select Nvl(����ѱ���,0) ���� From ҩƷ��� Where ҩƷID=[1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ҩƷ�Ĺ���ѱ���]", Val(.TextMatrix(.Row, 0)))
                        
                        .TextMatrix(.Row, mconIntCol�����) = zlStr.FormatEx((1 + rsTemp!���� / 100) * Val(.TextMatrix(.Row, mconIntCol�ɹ���)), mintCostDigit, , True)
                    End If
                    .TextMatrix(.Row, mconIntCol������) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�����)) * Val(strKey), mintMoneyDigit, , True)
                    
                    '˰��=�������*��ֵ˰��
                    .TextMatrix(.Row, mconIntCol˰��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�����)) * Val(strKey) * (Val(.TextMatrix(.Row, mconIntCol��ֵ˰��)) / 100 / (1 + Val(.TextMatrix(.Row, mconIntCol��ֵ˰��)) / 100)), mintMoneyDigit, , True)
                End If
                ��ʾ�ϼƽ��
            
            Case mconIntCol��������
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
            
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
                    If Val(strKey) < 0 Then
                        If Not zlStr.IsHavePrivs(mstrPrivs, "��������") Then
                            MsgBox "�Բ�����û�и���������Ȩ�ޣ������䣡", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) >= 0 Then
                        If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol����)) Then
                            MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    Else
                        If Val(strKey) < Val(.TextMatrix(.Row, mconIntCol����)) Then
                            MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "������������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If .TextMatrix(.Row, mconIntCol�ɹ���) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɹ���) * Val(strKey), mintMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * Val(strKey), mintMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol�����) <> "" Then
                        .TextMatrix(.Row, mconIntCol������) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�����)) * Val(strKey), mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconIntCol˰��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�����)) * Val(strKey) * (Val(.TextMatrix(.Row, mconIntCol��ֵ˰��)) / 100 / (1 + Val(.TextMatrix(.Row, mconIntCol��ֵ˰��)) / 100)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɹ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                End If
                ��ʾ�ϼƽ��
            Case mconIntCol�����
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "����۱���Ϊ�����ͣ������䣡", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If strKey <> "" Then
                    If Val(strKey) < 0.001 Then
                        MsgBox "�Բ�������۱������0.001,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "����۱���С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    .Text = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .TextMatrix(.Row, .Col) = .Text
                    
                    '����������
                    .TextMatrix(.Row, mconIntCol������) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�����)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
                    
                    '����˰��
                    .TextMatrix(.Row, mconIntCol˰��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�����)) * Val(.TextMatrix(.Row, mconIntCol����)) * (Val(.TextMatrix(.Row, mconIntCol��ֵ˰��)) / 100 / (1 + Val(.TextMatrix(.Row, mconIntCol��ֵ˰��)) / 100)), mintMoneyDigit, , True)
                End If
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lngҩƷID As Long, _
    ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, _
    ByVal strҩƷ��Դ As String, ByVal str����ҩ�� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, _
    ByVal numʵ�ʲ�� As Double, ByVal dbl�ӳ��� As Double, _
    ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal intҩ������ As Integer, ByVal str��׼�ĺ� As String, ByVal strԭ���� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsPrice As New Recordset
    
    Dim dbl����� As Double
    Dim dbl��ֵ˰�� As Double
    Dim dbl˰�� As Double
    Dim strҩ�� As String
    
    SetColValue = False
    On Error GoTo errHandle
    
    With mshBill

        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, 0) = lngҩƷID
        
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
        .TextMatrix(intRow, mconIntCol��Դ) = strҩƷ��Դ
        .TextMatrix(intRow, mconIntCol����ҩ��) = str����ҩ��
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColԭ����) = strԭ����
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ� * num����ϵ��, mintPriceDigit, , True)
        .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintNumberDigit, , True)
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mconIntcol�ӳ���) = dbl�ӳ��� & "||" & int�Ƿ��� & "||" & intҩ������
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol����) = lng����
        .TextMatrix(intRow, mconIntCol��ֵ˰��) = "100.00"
        
        If int�Ƿ��� = 1 Then
            dblPrice = Get���ۼ�(lngҩƷID, Val(cboStock.ItemData(cboStock.ListIndex)), lng����, num����ϵ��)
            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(dblPrice, mintPriceDigit, , True)
        End If
        
        If IsLowerLimit(cboStock.ItemData(cboStock.ListIndex), lngҩƷID) Then Call SetForeColor_ROW(mlng��ɫ)
        Call CheckLapse(strЧ��)
                
        If cboType.Text = "ҩƷ����" Then
            '������Ĭ��Ϊ�ɹ���=�����/����
            gstrSQL = "Select A.ָ��������, A.��ֵ˰��, Nvl(B.�ɹ���,0) As �ɹ��� " & _
                " From ҩƷ��� A, " & _
                " (Select ҩƷid, �ϴβɹ��� / Nvl(�ϴο���, 100) * 100 As �ɹ��� " & _
                " From ҩƷ��� " & _
                " Where ���� = 1 And �ⷿid + 0 = [1] And ҩƷid = [2] And Nvl(����, 0) = [3]) B " & _
                " Where A.ҩƷid = B.ҩƷid(+) And A.ҩƷid = [2]"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "ȡҩƷ������Ϣ", Val(cboStock.ItemData(cboStock.ListIndex)), lngҩƷID, lng����)
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol��ֵ˰��) = zlStr.FormatEx(rsPrice!��ֵ˰��, 2)
                
                If rsPrice!�ɹ��� > 0 Then
                    .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx(rsPrice!�ɹ��� * num����ϵ��, mintPriceDigit, , True)
                Else
                    .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx(rsPrice!ָ�������� * num����ϵ��, mintPriceDigit, , True)
                End If
            End If
        End If
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntColҩ�� Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
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
        If cboType.Text = "ҩƷ���" Then
            If cbo�����λ.ListIndex = 0 Then
                MsgBox "��ѡ��ҩƷ�����λ��", vbInformation, gstrSysName
                cbo�����λ.SetFocus
                Exit Function
            End If
        End If
        
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
                        .Col = mconIntCol����
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
                     
                    If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(intLop, 0)), Val(mshBill.TextMatrix(intLop, mconIntCol����)), _
                                    Val(mshBill.TextMatrix(intLop, mconIntCol����)), Val(.TextMatrix(intLop, mconIntCol����ϵ��)), _
                                    Trim(txtNo.Caption), 11, mint�����, mintNumberDigit) Then
                        mshBill.SetFocus
                        .Row = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
           
                    '���۹�������Ƿ���ڲ��������۵�ҩƷ
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                        If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                            If CheckPriceAdjust(Val(.TextMatrix(intLop, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intLop, mconIntCol����))) = False Then
                                MsgBox "��" & intLop & "��ҩƷ���������۹���������¼���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
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
    
    ValidData = True
End Function

Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim lng������id As Long
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngTypeID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim dblOutPrice As Double   '�����
    Dim strOutUnit As String    '�����λ
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim str��׼�ĺ� As String
    Dim blnTran As Boolean
    Dim dbl��ֵ˰�� As Double
    
    Dim rsTemp As New Recordset
    Dim n As Integer
    
    SaveCard = False
    arrSql = Array()
    
    '����������������ID����Ҫ������ҩƷ��Ҫ����
    On Error GoTo errHandle
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = Sys.GetNextNo(28, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        If cboType.Text = "ҩƷ���" Then
            If cbo�����λ.Text <> "" Then
                strOutUnit = Mid(cbo�����λ.Text, 1, InStr(1, cbo�����λ.Text, "-") - 1)
            Else
                MsgBox "������ҩƷ�����λ��", vbInformation, gstrSysName
                SaveCard = False
                Exit Function
            End If
        ElseIf cboType.Text = "ҩƷ����" Then
            If cbo������λ.Text <> "" Then
                strOutUnit = Mid(cbo������λ.Text, 1, InStr(1, cbo������λ.Text, "-") - 1)
            Else
                MsgBox "������ҩƷ������λ��", vbInformation, gstrSysName
                SaveCard = False
                Exit Function
            End If
        Else
            strOutUnit = ""
        End If
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lng������id = cboType.ItemData(cboType.ListIndex)
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strAssessor = Txt�����
        
        If blnǿ�Ʊ��� Then blnTran = True
        
        If mint�༭״̬ = 2 Or blnǿ�Ʊ��� Then        '�޸�
            gstrSQL = "zl_ҩƷ��������_Delete('" & mstr���ݺ� & "')"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
            
            strBooker = Txt������
            datBookDate = Format(Txt��������, "yyyy-mm-dd hh:mm:ss")
            strModifier = UserInfo.�û�����
            datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol����)
                strOldProducingArea = .TextMatrix(intRow, mconIntColԭ����)
                strBatchNo = .TextMatrix(intRow, mconIntCol����)
                lngBatchID = .TextMatrix(intRow, mconIntCol����)
                datTimeLimit = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                    '����ΪʧЧ��������
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = .TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol����ϵ��)
                
'                dblPurchasePrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblPurchasePrice = Get�ɱ���(lngDrugID, lngStockid, lngBatchID)
                
                dblPurchaseMoney = Val(zlStr.FormatEx(Val(FormatEx(dblPurchasePrice * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintCostDigit)) * Val(.TextMatrix(intRow, mconIntCol����)), mintMoneyDigit, , True)) ' .TextMatrix(intRow, mconIntCol�ɹ����)
                
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dblSalePrice = Get�ۼ�(Split(.TextMatrix(intRow, mconIntcol�ӳ���), "||")(1) = 1, lngDrugID, lngStockid, lngBatchID)
                
                dblSaleMoney = Val(zlStr.FormatEx(Val(FormatEx(dblSalePrice * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintPriceDigit)) * Val(.TextMatrix(intRow, mconIntCol����)), mintMoneyDigit, , True)) ' .TextMatrix(intRow, mconIntCol�ۼ۽��)
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                
                '�����ҩƷ�����������۵����㣬��ֱ����ȡ����ѱ��������������
                If Val(.TextMatrix(intRow, mconIntCol�����)) = 0 And cboType.Text = "ҩƷ���" Then
                    gstrSQL = "Select Nvl(����ѱ���,0) ���� From ҩƷ��� Where ҩƷID=[1]"
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ҩƷ�Ĺ���ѱ���]", lngDrugID)
                    
                    .TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx((1 + rsTemp!���� / 100) * Val(.TextMatrix(intRow, mconIntCol�ɹ���)), gtype_UserDrugDigits.Digit_�ɱ���)
                    .TextMatrix(intRow, mconIntCol������) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�����) * Val(.TextMatrix(intRow, mconIntCol����)), mintMoneyDigit, , True)
                End If
                If cboType.Text = "ҩƷ���" Or cboType.Text = "ҩƷ����" Then
                    dblOutPrice = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�����)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                End If
                dblMistakePrice = Val(zlStr.FormatEx(dblSaleMoney - dblPurchaseMoney, mintMoneyDigit, , True)) '.TextMatrix(intRow, mconintCol���)
                
                dbl��ֵ˰�� = Val(.TextMatrix(intRow, mconIntCol��ֵ˰��))
                
'                If Val(.TextMatrix(intRow, mconIntCol���)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol���))
'                End If
                lngSerial = intRow
                
                gstrSQL = "zl_ҩƷ��������_INSERT("
                '������ID
                gstrSQL = gstrSQL & lng������id
                'NO
                gstrSQL = gstrSQL & ",'" & chrNo & "'"
                '���
                gstrSQL = gstrSQL & "," & lngSerial
                '�ⷿID
                gstrSQL = gstrSQL & "," & lngStockid
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngDrugID
                '����
                gstrSQL = gstrSQL & "," & lngBatchID
                '��д����
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
                '�����(������)
                gstrSQL = gstrSQL & "," & dblOutPrice
                '�����λ(������λ)
                gstrSQL = gstrSQL & ",'" & strOutUnit & "'"
                '������
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '����
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '��ֵ˰��
                gstrSQL = gstrSQL & "," & dbl��ֵ˰��
                'ԭ����
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '�޸���
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '�޸�����
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"
                    
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngDrugID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSql, MStrCaption, False, Not blnǿ�Ʊ���) Then Exit Function
        If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans: blnTran = False
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double, Cur������ As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            Cur������ = Cur������ + Val(.TextMatrix(intLop, mconIntCol������))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & zlStr.FormatEx(curTotal, mintMoneyDigit, , True)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & zlStr.FormatEx(Cur���ʽ��, mintMoneyDigit, , True)
    lblDifference.Caption = "��ۺϼƣ�" & zlStr.FormatEx(Cur���ʲ��, mintMoneyDigit, , True)
    lblOther.Caption = "���(��)�ϼƣ�" & zlStr.FormatEx(Cur������, mintMoneyDigit, , True)
End Sub

Private Sub ��ʾ�����()
    Dim rsUseCount As New Recordset
    
    On Error GoTo errHandle
    With mshBill
        If .TextMatrix(.Row, mconIntColҩ��) = "" Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        gstrSQL = "select ��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as  ��������   from ҩƷ��� where �ⷿid=[1] " _
            & " and ҩƷid=[2] " _
            & " and ����=1 and " _
            & " nvl(����,0)=[3]"
        Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
            
        If rsUseCount.EOF Then
            .TextMatrix(.Row, mconIntCol��������) = 0
        Else
            .TextMatrix(.Row, mconIntCol��������) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
        End If
        rsUseCount.Close
        
        staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol��������), mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol��λ)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtIn_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim IntCheck As Integer
    Dim intRow As Integer
    Dim blnEXIST As Boolean
    Dim intIndex As Integer, intCount As Integer
    Dim rsBill As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    On Error GoTo ErrHand
    
    '��ʼ׼��
    intNO = 28
    lng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtIn.Text) = "" Then Exit Sub
    
    If Len(txtIn.Text) < 8 Then
        txtIn.Text = zlCommFun.GetFullNO(txtIn.Text, intNO, lng�ⷿID)
    End If
    
    '����������ΪҩƷ���
    intCount = cboType.ListCount
    For intIndex = 1 To intCount
        If cboType.List(intIndex - 1) = "ҩƷ���" Then
            cboType.ListIndex = intIndex - 1
            blnEXIST = True
            Exit For
        End If
    Next
'    If Not blnEXIST Then
'        MsgBox "�����⹺��ⵥ�Ĺ���ֻ��Ӧ����������ҩƷ�������", vbInformation, gstrSysName
'        Exit Sub
'    End If
    
    '��ҪҪ������е�������
    For IntCheck = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(IntCheck, 0) <> "" Then
            Exit For
        End If
    Next
    If IntCheck <> mshBill.rows Then
        If MsgBox("��ҪҪ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '����ҩƷ��λ�ı�
        mshBill.ClearBill
    End If
    
    'ȡ����������
    IntCheck = 0
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ����������]", Me.cboStock.ItemData(Me.cboStock.ListIndex))

    If rsTemp.RecordCount <> 0 Then IntCheck = rsTemp!�����
    
    '��ȡ�õ��ݲ���ձ��ֻ������ȡ�������ݣ��ҷ��˻�����
    gstrSQL = "SELECT A.ҩƷID,'['||C.����||']' As ����,'['||C.����||']'|| Nvl(F.����,C.����) As ҩƷ����, C.���� As ͨ����,F.���� As ��Ʒ��,C.���,C.����,A.ԭ����," & _
             "        C.���㵥λ AS ���۵�λ,1 AS ����ϵ��,B.���ﵥλ,B.�����װ,B.סԺ��λ,B.סԺ��װ,B.ҩ�ⵥλ,B.ҩ���װ, " & _
             "        NVL(A.����,0) AS ����,Nvl(C.�Ƿ���,0) AS ʱ��,Nvl(B.ҩ������,0) AS ҩ������,A.����,A.Ч��," & _
             "        B.����ѱ���,B.�ӳ���,A.ʵ������,D.��������,D.ʵ�ʽ��,D.ʵ�ʲ��,E.�ּ�,A.��׼�ĺ�,B.ҩƷ��Դ,B.����ҩ��,d.ƽ���ɱ��� " & _
             " FROM ҩƷ�շ���¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,ҩƷ��� D,�շѼ�Ŀ E,�շ���Ŀ���� F " & _
             " WHERE A.ҩƷID=B.ҩƷID AND B.ҩƷID=C.ID AND B.ҩƷID=D.ҩƷID(+) " & _
             " AND B.ҩƷID=F.�շ�ϸĿID(+) AND F.����(+)=3 AND F.����(+)=1" & _
             " AND B.ҩƷID=E.�շ�ϸĿID(+) AND SYSDATE >=E.ִ������(+)  AND sysdate<=NVL(E.��ֹ����(+),SYSDATE)" & _
             GetPriceClassString("E") & _
             " AND D.�ⷿID(+)=[2] AND D.����(+)=1 AND Nvl(A.����,0)=Nvl(D.����,0)" & _
             " AND A.����=1 AND A.��¼״̬=1 AND NVL(A.��ҩ��ʽ,0)=0 AND A.������� Is Not NULL" & _
             " AND A.NO=[1] And A.�ⷿID+0=[2] " & _
             " ORDER BY A.���"
    Set rsBill = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ�⹺��ⵥ]", txtIn.Text, Me.cboStock.ItemData(Me.cboStock.ListIndex))
             
    If rsBill.RecordCount = 0 Then
        MsgBox "û���ҵ����⹺��ⵥ�ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsBill
        intRow = 1
        Do While Not .EOF
            'װ������ǰ���ȼ����
            If !ʵ������ > !�������� Then
                '���λ�ʱ��ҩƷ�����������
                If !���� <> 0 Or !ʱ�� <> 0 Then
                    MsgBox !ҩƷ���� & "��治�㣬��������⣡��ʱ�ۻ����ҩƷ��", vbInformation, gstrSysName
                    mshBill.ClearBill
                    Exit Sub
                End If
                Select Case IntCheck
                Case 1
                    If MsgBox(!ҩƷ���� & "�Ѿ�û�п�棬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        mshBill.ClearBill
                        Exit Sub
                    End If
                Case 2
                    MsgBox !ҩƷ���� & "�Ѿ�û�п�棬���ܼ���������", vbInformation, gstrSysName
                    mshBill.ClearBill
                    Exit Sub
                End Select
            End If
            
            'װ������(SetColValue)
            If Not SetColValue(intRow, !ҩƷID, !����, !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), _
                nvl(!ҩƷ��Դ), nvl(!����ҩ��), nvl(!���), nvl(!����), _
                Choose(mintUnit, !���۵�λ, !���ﵥλ, !סԺ��λ, !ҩ�ⵥλ), nvl(!�ּ�, 0), _
                nvl(!����), nvl(!Ч��), nvl(!��������, 0), nvl(!ʵ�ʽ��, 0), nvl(!ʵ�ʲ��, 0), _
                nvl(!�ӳ��� / 100, 0), Choose(mintUnit, 1, !�����װ, !סԺ��װ, !ҩ���װ), nvl(!����, 0), !ʱ��, _
                !ҩ������, IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�), nvl(!ԭ����)) Then
                mshBill.ClearBill
                Exit Sub
            End If
            
            '��д�������ɹ��۵���
            mshBill.TextMatrix(intRow, mconIntCol�к�) = intRow
            mshBill.TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(nvl(!ʵ������, 0) / Choose(mintUnit, 1, !�����װ, !סԺ��װ, !ҩ���װ), mintNumberDigit, , True)
            If mshBill.TextMatrix(intRow, mconIntCol�ۼ�) <> "" Then
                mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol�ۼ�)) * Val(mshBill.TextMatrix(intRow, mconIntCol����)), mintMoneyDigit, , True)
            End If
            
'            mshBill.TextMatrix(intRow, mconintCol���) =Str.FormatEx(Get������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(mshBill.TextMatrix(intRow, 0)), Val(mshBill.TextMatrix(intRow, mconIntCol����)), Val(mshBill.TextMatrix(intRow, mconIntColʵ�ʽ��)), Val(mshBill.TextMatrix(intRow, mconIntColʵ�ʲ��)), Val(mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��)), Val(mshBill.TextMatrix(intRow, mconIntCol����)) * Val(mshBill.TextMatrix(intRow, mconIntCol����ϵ��))), mintMoneyDigit)
            
            If nvl(!ʵ������, 0) <> 0 Then
                mshBill.TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(Get�ɱ���(Val(mshBill.TextMatrix(intRow, 0)), Val(cboStock.ItemData(cboStock.ListIndex)), Val(mshBill.TextMatrix(intRow, mconIntCol����))) * Val(mshBill.TextMatrix(intRow, mconIntCol����ϵ��)), mintCostDigit, , True)
'                mshBill.TextMatrix(intRow, mconIntCol�ɹ���) =Str.FormatEx((mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��) - mshBill.TextMatrix(intRow, mconintCol���)) / Val(mshBill.TextMatrix(intRow, mconIntCol����)), mintCostDigit)
            End If
            mshBill.TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol�ɹ���)) * Val(mshBill.TextMatrix(intRow, mconIntCol����)), mintMoneyDigit, , True)
            mshBill.TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��)) - Val(mshBill.TextMatrix(intRow, mconIntCol�ɹ����)), mintMoneyDigit, , True)
            
            '��������ۼ�������:�����=(1+����ѱ���)*����
            mshBill.TextMatrix(intRow, mconIntCol�����) = zlStr.FormatEx((1 + !����ѱ��� / 100) * Val(mshBill.TextMatrix(intRow, mconIntCol�ɹ���)), mintCostDigit, , True)
            mshBill.TextMatrix(intRow, mconIntCol������) = zlStr.FormatEx(Val(mshBill.TextMatrix(intRow, mconIntCol�����)) * Val(mshBill.TextMatrix(intRow, mconIntCol����)), mintMoneyDigit, , True)
            
            intRow = intRow + 1
            mshBill.rows = mshBill.rows + 1
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshBill.ClearBill
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
    Dim strUnit As String
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
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1306", "zl8_bill_1306"), mint��¼״̬, int��λϵ��, 1306, "ҩƷ�������ⵥ", strNo
End Sub

Private Sub SetForeColor_ROW(ByVal lngColor As Long)
    Dim i As Integer, j As Integer
    Dim intCol As Integer
    '����ĳ�е���ɫ
    With mshBill
        intCol = .Col
        mblnEnterCell = False
        For i = mconIntColҩ�� To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            .MsfObj.CellForeColor = lngColor
            .ColData(i) = j
        Next
        .Col = intCol
        mblnEnterCell = True
    End With
End Sub

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ���ʱ��ҩƷ�Ƿ��п��

    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim str��� As String
    Dim strSQL As String
    Dim strDub As String    '�ظ�ҩƷ
    Dim strNotNum As String  '�޿��ҩƷ
    Dim str�ظ�ҩ�� As String   '������¼�ظ�ѡ���˵�ҩƷ����
    Dim strNotҩ�� As String    '������¼��ЩҩƷ��ʱ�۵��޿��
    
    On Error GoTo errHandle

    rsTemp.MoveFirst
    str���� = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        If InStr(1, strTemp, rsTemp!ҩƷID & "," & str����) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷID & "," & str���� & "," & rsTemp!ͨ���� & "|"
        End If
    
        rsTemp.MoveNext
    Loop
        
    With mshBill    '���ظ��Ĳ�ѯ����
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
        '�ж���ʲô��ʽƴ��sql
        If str�ظ�ҩ�� <> "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPrice(ByVal lngҩƷID As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double) As Double
    Dim rsPrice As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(����,0),0,ʵ�ʽ��/ʵ������,Nvl(���ۼ�,ʵ�ʽ��/ʵ������))*" & dbl����ϵ�� & " as  �ۼ� " _
        & "  from ҩƷ��� " _
        & " where �ⷿid=[1] " _
        & " and ҩƷid=[2] " _
        & " and ����=1 and ʵ������>0 and " _
        & " nvl(����,0)=[3]"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lngҩƷID, lng����)

    If rsPrice.EOF Then
        gstrSQL = "Select �ּ� From �շѼ�Ŀ Where �շ�ϸĿid = [1] And Sysdate Between ִ������ And ��ֹ����" & _
                GetPriceClassString("")
        
        Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷID)
        If rsPrice.RecordCount > 0 Then
            GetPrice = rsPrice!�ּ� * dbl����ϵ��
        End If
        Exit Function
    End If
    GetPrice = rsPrice.Fields(0).Value
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Function ���۸�() As Boolean
    '���ܣ�����ʱ���ж�ҩƷ�Ƿ������¼۸񣬲������޸ĺ���ʾ
    Dim strMsg As String '������ʾ��Ϣ
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim bln�Ƿ�ʱ�� As Boolean
    
    On Error GoTo errHandle
    
    ���۸� = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Trim(.TextMatrix(i, mconIntCol����)) <> "" Then
            
                bln�Ƿ�ʱ�� = Val(Split(.TextMatrix(i, mconIntcol�ӳ���), "||")(1)) = 1
                Dbl���� = Val(.TextMatrix(i, mconIntCol����))
                
                '���ɱ���
                dbl�ɱ��� = zlStr.FormatEx(Get�ɱ���(Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol����))) * Val(.TextMatrix(i, mconIntCol����ϵ��)), mintCostDigit)
                If .TextMatrix(i, mconIntCol�ɹ���) <> dbl�ɱ��� Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol�ɹ���) = zlStr.FormatEx(dbl�ɱ���, mintCostDigit, , True)
                    .TextMatrix(i, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(i, mconIntCol�ɹ���) * Dbl����, mintMoneyDigit, , True)
                End If
                
                '����ۼ�
                dbl���ۼ� = zlStr.FormatEx(Get�ۼ�(bln�Ƿ�ʱ��, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol����))) * Val(.TextMatrix(i, mconIntCol����ϵ��)), mintPriceDigit)
                If .TextMatrix(i, mconIntCol�ۼ�) <> dbl���ۼ� Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, mintPriceDigit, , True)
                    .TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(i, mconIntCol�ۼ�) * Dbl����, mintMoneyDigit, , True)
                End If
                
                .TextMatrix(i, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(i, mconIntCol�ۼ۽��)) - Val(.TextMatrix(i, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "�м�¼δʹ�����¼۸񣬳������Զ���ɸ��£��ɱ��ۡ��ɱ����ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
            ���۸� = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

