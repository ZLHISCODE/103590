VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmRequestDrugCard 
   Caption         =   "ҩƷ���쵥"
   ClientHeight    =   7770
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14280
   Icon            =   "frmRequestDrugCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   14280
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȫ������ 
      Caption         =   "ȫ������"
      Height          =   350
      Left            =   9360
      TabIndex        =   31
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdȫ�� 
      Caption         =   "ȫ�����"
      Height          =   350
      Left            =   8040
      TabIndex        =   30
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CheckBox chkExportPlan 
      Caption         =   "����ʱֻͬ�������ǳ���ҩƷ�ļƻ�����"
      Height          =   380
      Left            =   5160
      TabIndex        =   29
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   6480
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   5160
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3240
      TabIndex        =   9
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   1560
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8040
      TabIndex        =   5
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9360
      TabIndex        =   6
      Top             =   5520
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   14175
      TabIndex        =   10
      Top             =   0
      Width           =   14235
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   4
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   557
         Width           =   1515
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   2
         Top             =   950
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
      Begin VB.Label Txt�޸����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7560
         TabIndex        =   35
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�޸��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5520
         TabIndex        =   34
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label lbl�޸��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸���"
         Height          =   180
         Left            =   4920
         TabIndex        =   33
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label lbl�޸����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸�����"
         Height          =   180
         Left            =   6780
         TabIndex        =   32
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   24
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10110
         TabIndex        =   21
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12210
         TabIndex        =   20
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   17
         Top             =   550
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
         TabIndex        =   16
         Top             =   587
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ���쵥"
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
         TabIndex        =   15
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ�ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   617
         Width           =   990
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   9525
         TabIndex        =   12
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   11400
         TabIndex        =   11
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
            Picture         =   "frmRequestDrugCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1000
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
            Picture         =   "frmRequestDrugCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   7410
      Width           =   14280
      _ExtentX        =   25188
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestDrugCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18838
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestDrugCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestDrugCard.frx":3080
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
      Left            =   2760
      TabIndex        =   22
      Top             =   5160
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
Attribute VB_Name = "frmRequestDrugCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5��ͨ����������6�����ܣ����պ��¼���յǼ��ˣ�����ȡ������Ľ��գ���7������
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mbln����״̬ As Boolean
Private mstr�ⷿ As String                  '��¼�Ѿ�����˵Ŀⷿ

Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mint��������ⷿ As Integer     '�����ڳ���ʱ��ԭ���ⷿ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrPrivs As String                     'Ȩ��
Private mlngStockID As Long                 '��ǰ�û���ѡ��ҩ��ID
Private mintApplyType As Integer            '���췽ʽ��0-�ֹ�����;1-����������;2-��������;3-��������;4-����������;5-�������쵥δ����;6-������������;7-������������
Private mstrEndTime As String               '���Զ����췽ʽΪ7ʱ������ʱ�䷶Χ�еĽ���ʱ��
Private rsDepend As New ADODB.Recordset

Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mstrTime_Start As String                        '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
Private mblnUpdate As Boolean               '������¼������˺��Ƿ�������¼۸�
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private Const MStrCaption As String = "ҩƷ�������"
Private mint��ʾ��ǰ��淽ʽ As Integer     '0-��ʾ���ʵ������,1-��ʾ����������
Private mint��ʾ�Է���淽ʽ As Integer     '0-��ʾ���ʵ������,1-��ʾ����������
Private mint��ǰ��水������ʾ As Integer   '0-������������ʾ,1-����ǰ�ⷿ��ҩƷ��������ʾ
Private mint�����γ��� As Integer           '0-�������γ���,1-�����γ���
Private mblnLoad As Boolean              '��¼�Ƿ�ִ�����Form_Load�¼�

'�Ӳ�������ȡҩƷ�۸����������С��λ�������㾫�ȣ�
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mint����ʽ As Integer             '����ʱ��0������������1�������������뵥��

Private mbln����� As Boolean

'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntColҩ�� As Integer = 2
Private Const mconIntCol��Ʒ�� As Integer = 3
Private Const mconIntCol��Դ As Integer = 4
Private Const mconIntCol����ҩ�� As Integer = 5
Private Const mconIntCol��� As Integer = 6
Private Const mconIntCol��� As Integer = 7
Private Const mconIntCol�������� As Integer = 8
Private Const mconIntCol���Ч�� As Integer = 9
Private Const mconIntCol�������� As Integer = 10
Private Const mconIntcol�ӳ��� As Integer = 11
Private Const mconIntColʵ�ʽ�� As Integer = 12
Private Const mconIntColʵ�ʲ�� As Integer = 13
Private Const mconIntCol����ϵ�� As Integer = 14
Private Const mconIntCol���� As Integer = 15
Private Const mconIntCol���� As Integer = 16
Private Const mconIntColԭ���� As Integer = 17
Private Const mconIntCol��λ As Integer = 18
Private Const mconIntCol���� As Integer = 19
Private Const mconIntColЧ�� As Integer = 20
Private Const mconIntCol��׼�ĺ� As Integer = 21
Private Const mconintcol��ǰ��� As Integer = 22
Private Const mconintcol�Է���� As Integer = 23
Private Const mconIntCol�������� As Integer = 24
Private Const mconIntCol��д���� As Integer = 25
Private Const mconIntColʵ������ As Integer = 26
Private Const mconIntCol�ɹ��� As Integer = 27
Private Const mconIntCol�ɹ���� As Integer = 28
Private Const mconIntCol�ۼ� As Integer = 29
Private Const mconIntCol�ۼ۽�� As Integer = 30
Private Const mconintCol��� As Integer = 31
Private Const mconIntCol�ϴι�Ӧ��ID As Integer = 32
Private Const mconintCol��ʵ���� As Integer = 33
Private Const mconIntColҩƷ��������� As Integer = 34
Private Const mconIntColҩƷ���� As Integer = 35
Private Const mconIntColҩƷ���� As Integer = 36
Private Const mconIntCol����ҩƷ As Integer = 37
Private Const mconIntColԭʼ���� As Integer = 38
Private Const mconIntColS  As Integer = 39            '������
'=========================================================================================

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
Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    CheckBill = ""
    On Error GoTo errHandle
    gstrSQL = " Select �������,��ҩ����,��ҩ�� From ҩƷ�շ���¼ " & _
            " Where ����=6 And NO=[1] And ��¼״̬=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��鵥��]", strNo)
    
    With rs
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then
            CheckBill = "�õ����Ѿ�����������Աɾ����"
        ElseIf Not IsNull(!�������) Then
            CheckBill = "�õ����Ѿ�����������Ա��ˣ�"
        ElseIf Not IsNull(!��ҩ����) Then
            CheckBill = "�õ����Ѿ�����������Ա���ͣ�"
        ElseIf Not IsNull(!��ҩ��) Then
            CheckBill = "�õ����Ѿ�����������Ա��ҩ��"
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'�������������
Private Function GetDepend() As Boolean
    Dim strMsg As String

    GetDepend = False
    On Error GoTo ErrHand

    '���ҩƷ�������Ƿ�����
    strMsg = "û������ҩƷ�ƿ����⼰�����������ҩƷ������࣡"
    gstrSQL = "SELECT B.Id,B.ϵ�� " _
            & "FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID AND A.���� = 6"
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�ƿ����")

    With rsDepend
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "ϵ��=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û������ҩƷ�ƿ������������ҩƷ������࣡"
            GoTo ErrHand
        End If
        .Filter = "ϵ��=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û������ҩƷ�ƿ�ĳ����������ҩƷ������࣡"
            GoTo ErrHand
        End If
        .Filter = 0
        
        'gstrSQL = ReturnSQL(mlngStockID, False)
    End With
    'Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "ҩƷ�������", mlngStockID)
    Set rsDepend = ReturnSQL(mlngStockID, "ҩƷ�������", False, 1343)

    strMsg = "û���κοⷿ�������죬����[������������]��ҩƷ���������ã�"
    If rsDepend.RecordCount = 0 Then
        MsgBox strMsg, vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    GetDepend = True
    Exit Function
ErrHand:
    
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False, Optional lngStockid As Long = 0, Optional int����ʽ As Integer = 0, Optional intApplyType As Integer = 0)
    Dim strsql As String
    Dim rsPara As New ADODB.Recordset
    
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mint����ʽ = int����ʽ
    mintApplyType = intApplyType
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1343)
    mlngStockID = IIf(lngStockid = 0, glngDeptId, lngStockid)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    mint��������ⷿ = MediWork_GetCheckStockRule(mlngStockID)
    
    If mint�༭״̬ <> 5 Then
        Me.cmdȫ������.Visible = False
        Me.cmdȫ��.Visible = False
    End If
    
    mblnEdit = False
         
    If mint�༭״̬ = 5 Then
        Me.Height = Me.Height + Me.cmdȫ��.Height
    End If
         
    If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then
        mblnEdit = True
        mblnFirst = True
        
        chkExportPlan.Visible = True
    
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
        mblnFirst = True
    ElseIf mint�༭״̬ = 4 Then
        mblnFirst = True
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If Not IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint�༭״̬ = 7 Then
        mblnEdit = False
        mblnFirst = True
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        If mint����ʽ = 1 Then
            CmdSave.Caption = "�������(&O)"
            CmdSave.Width = CmdSave.Width + 200
        Else
            CmdSave.Caption = "����(&O)"
            CmdSave.Width = CmdCancel.Width
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub


Private Sub cboStock_Click()
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        str�ⷿ���� = ""
        gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
            rsDetail.MoveNext
        Loop
        If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
        mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
    
        If mblnLoad = True Then Call SetSelectorRS(2, "ҩƷ�������", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln�����)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(0, mintMoneyDigit, , True)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
    If mint�༭״̬ <> 7 Then Call CheckNumber
    mblnChange = False
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntColʵ������) = .TextMatrix(intRow, mconIntCol��д����)
                .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol�ɹ���), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol�ۼ�), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ۽��) - .TextMatrix(intRow, mconIntCol�ɹ����), mintMoneyDigit, , True)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
    If mint�༭״̬ <> 7 Then Call CheckNumber
    mblnChange = True
End Sub

Private Sub cboStock_Change()
    mblnChange = True
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
                    
                    If Me.mshBill.ColWidth(mconIntCol��������) > 0 Then
                        Me.mshBill.ColWidth(mconIntCol��������) = 0
                        Me.cmdȫ������.Visible = False
                        Me.cmdȫ��.Visible = False
                        Call Form_Resize
                    End If
                    
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    End With
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
        FindRow mshBill, mconIntColҩ��, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdȫ������_Click()
    Dim Row As Integer
    Dim count As Integer
    
    For Row = 1 To Me.mshBill.rows - 1
        If Val(Me.mshBill.TextMatrix(Row, 0)) <> 0 Then
            count = count + 1
            Exit For
        End If
    Next
    
    If count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫ������������ֵ��Ϊ��д������ʵ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        For Row = 1 To Me.mshBill.rows - 2
            Me.mshBill.TextMatrix(Row, mconIntCol��д����) = Me.mshBill.TextMatrix(Row, mconIntCol��������)
            Me.mshBill.TextMatrix(Row, mconIntColʵ������) = Me.mshBill.TextMatrix(Row, mconIntCol��������)
            If Val(Me.mshBill.TextMatrix(Row, mconIntCol��д����)) <> 0 Then
                Call GetPrice(Row)
            Else
                With Me.mshBill
                    .TextMatrix(Row, mconIntCol�ۼ۽��) = 0
                    .TextMatrix(Row, mconintCol���) = 0
                    .TextMatrix(Row, mconIntCol�ɹ���) = 0
                    .TextMatrix(Row, mconIntCol�ɹ����) = 0
                End With
            End If
        Next
        Call ��ʾ�ϼƽ��
        If mint�༭״̬ <> 7 Then Call CheckNumber
    End If
End Sub

Private Sub cmdȫ��_Click()
    Dim Row As Integer
    Dim count As Integer
    
    For Row = 1 To Me.mshBill.rows - 1
        If Val(Me.mshBill.TextMatrix(Row, 0)) <> 0 Then
            count = count + 1
            Exit For
        End If
    Next
    
    If count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫ����д������ʵ��������Ϊ0��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        For Row = 1 To Me.mshBill.rows - 2
            Me.mshBill.TextMatrix(Row, mconIntCol��д����) = 0
            Me.mshBill.TextMatrix(Row, mconIntColʵ������) = 0
            With Me.mshBill
                .TextMatrix(Row, mconIntCol�ۼ۽��) = 0
                .TextMatrix(Row, mconintCol���) = 0
                .TextMatrix(Row, mconIntCol�ɹ���) = 0
                .TextMatrix(Row, mconIntCol�ɹ����) = 0
            End With
        Next
        Call ��ʾ�ϼƽ��
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then
        If mshBill.rows > 50 Then
            Call AviShow(Me) '��ʾ�û����ڲ�ѯ����
        End If
        Call get�������    'Ϊ��ǰ��������ͶԷ���������и�ֵ
        If mshBill.rows > 50 Then
            Call AviShow(Me, False)
        End If
        Exit Sub
    End If
    
    mblnFirst = False
    If mint�༭״̬ = 5 Then
        If Not frmRequestNavigation.ShowNavigation(Me, mlngStockID, mintApplyType, mstrEndTime, mbln����״̬) = True Then
            Unload Me
            Exit Sub
        End If
        If mint�༭״̬ <> 7 Then Call CheckNumber
        mshBill.SetFocus
        If mintApplyType = 7 And Not IsHavePrivs(mstrPrivs, "�Զ�����ʱ�޸�ҩƷ����") Then
            mshBill.Active = False
        End If
    End If
    If mbln����״̬ = True Then
        Call Form_Resize
    End If
    
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 7 Then
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


Private Sub cmdSave_Click()
    Dim BlnSuccess As Boolean
    Dim Row As Integer
    Dim count As Integer
    Dim intRows As Integer
    Dim lng�ϴ�ҩƷID As Long
    
    '�����������ݼ�
    Call SetSortRecord
        
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If Me.mshBill.TextMatrix(Me.mshBill.rows - 1, 0) <> "" Then
        intRows = Me.mshBill.rows - 1
    Else
        intRows = Me.mshBill.rows - 2
    End If
    
    For Row = 1 To intRows
        If Val(Me.mshBill.TextMatrix(Row, mconIntCol��д����)) = 0 Then
            count = count + 1
            If count = intRows Then
                MsgBox "�����쵥�ϵ�����ҩƷ����д������Ϊ0�����ܼ���������", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
    Next

    For Row = 1 To Me.mshBill.rows - 2
        If zlStr.Nvl(Me.mshBill.TextMatrix(Row, mconIntCol��д����), 0) = 0 Then
            If MsgBox("�����쵥������д����Ϊ0��ҩƷ��" & vbCrLf & "��д����Ϊ0��ҩƷ�����ܱ���Ϊ���쵥��" & vbCrLf & "�Ƿ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next
    
    '��������ҩƷ����Ԥ���۴���
    For Row = 1 To Me.mshBill.rows - 1
        If mshBill.TextMatrix(Row, 0) <> "" Then '��ҩƷ
            Call AutoAdjustPrice_ByID(Val(mshBill.TextMatrix(Row, 0)))
        End If
    Next
    
    
    
    '10.35.30�޸ģ��Զ����첻�����
'    If mint�༭״̬ = 5 Then '�Զ�������Ҫ�������
'        mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
'        For Row = 1 To intRows
'            If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(Row, 0)), Val(mshBill.TextMatrix(Row, mconIntCol����)), Val(Me.mshBill.TextMatrix(Row, mconIntCol��д����)), Val(mshBill.TextMatrix(Row, mconIntCol����ϵ��)), txtNo.Caption, 6, mint�����) Then
'                mshBill.Row = Row
'                mshBill.Col = mconIntCol��д����
'                mshBill.SetFocus
'                Exit Sub
'            End If
'        Next
'    End If
    
    If mint�༭״̬ = 6 Then       '���
        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub
        If SaveCheck() = True Then
            If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, 1343)) = 1 Then
                '��ӡ
                If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                    printbill
                    
                    If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, 1343)) = 1 And IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                        '��ҩƷID˳���������
                        recSort.Sort = "ҩƷid"
                        recSort.MoveFirst
                        '��ӡҩƷ����
                        Do While Not recSort.EOF
                            If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                                ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1343_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
                                lng�ϴ�ҩƷID = recSort!ҩƷID
                            End If
                            recSort.MoveNext
                        Loop
                    End If
                    
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 7 Then '����
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
'        If Not ��鵥��(6, txtNO.Tag, False, True) Then
'            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
'            Call RefreshBill
'            mblnUpdate = True
'            Exit Sub
'        End If
        
        If ���۸� Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    
    End If
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then '��������ʱ���жϼ۸��Ƿ��Ѿ�����
        If ���۸� Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If ValidData = False Then Exit Sub
    
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, 1343)) = 1 Then
            '��ӡ
            If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, 1343)) = 1 And IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1343_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
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
    cboStock.SetFocus
    mblnChange = False

    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Load()
    Dim strStock As String
    Dim rsStock As New Recordset
    Dim intStock As Integer
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    
    On Error GoTo errHandle
    
    mblnLoad = False
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    mblnViewCost = IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    chkExportPlan.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "ͬ�����ɼƻ���", 0))
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    mbln����� = (Val(zlDataBase.GetPara("��ʾ�޿��ҩƷ", glngSys, 1343, 0)) = 0)
    mint��ʾ��ǰ��淽ʽ = Val(zlDataBase.GetPara("�ʱ��ǰ�ⷿ�����ʾ��ʽ", glngSys, 1343, 0))
    mint��ʾ�Է���淽ʽ = Val(zlDataBase.GetPara("�ʱ�Է��ⷿ�����ʾ��ʽ", glngSys, 1343, 0))
    mint��ǰ��水������ʾ = Val(zlDataBase.GetPara("��ǰ�ⷿҩƷ�����Ƿ�������ʾ", glngSys, 1343, 0))
    mint�����γ��� = Val(zlDataBase.GetPara("ҩƷ�����γ���", glngSys, 1343, 0))
    
    intStock = -1
    With cboStock
        .Clear
        mstr�ⷿ = ""
        Do While Not rsDepend.EOF
            If InStr(1, mstr�ⷿ, "|" & rsDepend!Id & "|") = 0 Then
                .AddItem rsDepend!����
                .ItemData(.NewIndex) = rsDepend!Id
                mstr�ⷿ = mstr�ⷿ & "|" & rsDepend!Id & "|"
                
                If rsDepend!ҩ������ = 1 And intStock = -1 Then
                    intStock = .NewIndex
                End If
            End If
            
            rsDepend.MoveNext
        Loop
        .ListIndex = IIf(intStock = -1, 0, intStock)
    End With
    
    If mlngStockID = 0 Then
        mlngStockID = mfrmMain.cboStock.ItemData(Me.cboStock.ListIndex)
    End If
    Call GetDrugDigit(mlngStockID, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(6, mstr���ݺ�)
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
        
    '����ϵͳ��������ҩ����Ա�鿴����ʱ���Ƿ���ʾ�ɱ���
    mshBill.ColWidth(mconIntCol�ɹ���) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol�ɹ����) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol���) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconIntCol�ϴι�Ӧ��ID) = 0
    mshBill.ColWidth(mconintCol��ʵ����) = 0
    mshBill.ColWidth(mconIntCol��������) = 0
    
'    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
'        mshBill.ColWidth(mconintcol��ǰ���) = 1100
'        mshBill.ColWidth(mconintcol�Է����) = 1100
'    Else
'        mshBill.ColWidth(mconintcol��ǰ���) = 0
'        mshBill.ColWidth(mconintcol�Է����) = 0
'    End If
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
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
    Dim lngStockid As Long
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim strUnitQuantity_Stock As String
    Dim intRow As Integer
    Dim vardrug As Variant
    Dim numUseAbleCount As Double
    Dim dateCurDate As Date
    Dim strOrder As String, strCompare As String
    Dim intCount As Integer
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPricedigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("����", glngSys, 1343)
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
        
    If mint�༭״̬ = 4 Then
        With cboStock
            'ȡָ�����ݵĳ���ⷿ�����ⷿ
            gstrSQL = " Select �ⷿID,�Է�����ID From ҩƷ�շ���¼" & _
                      " Where NO=[1] And ����=6 And ���ϵ��=-1 And Rownum<2"
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡָ�����ݵĳ���ⷿ�����ⷿ]", mstr���ݺ�)
            
            If rsInitCard.RecordCount <> 0 Then
                lngStockid = rsInitCard!�ⷿid
            End If
            
            For intCount = 0 To .ListCount - 1
                If .ItemData(intCount) = lngStockid Then
                    .ListIndex = intCount: Exit For
                End If
            Next
        End With
    Else
        With cboStock
            If Not (mint�༭״̬ = 1 Or mint�༭״̬ = 5) Then
                'ȡָ�����ݵĳ���ⷿ�����ⷿ
                gstrSQL = " Select �ⷿID,�Է�����ID From ҩƷ�շ���¼" & _
                          " Where NO=[1] And ����=6 And ���ϵ��=-1 And Rownum<2"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡָ�����ݵĳ���ⷿ�����ⷿ]", mstr���ݺ�)
                
                If rsInitCard.RecordCount <> 0 Then
                    lngStockid = rsInitCard!�ⷿid
                End If
            End If
            For intCount = 0 To .ListCount - 1
                If .ItemData(intCount) = lngStockid Then
                    .ListIndex = intCount: Exit For
                End If
            Next
            mintcboIndex = .ListIndex
        End With
    End If
    
    If mint�༭״̬ = 7 Then
       lngStockid = mlngStockID
    End If
    
    dateCurDate = Sys.Currentdate()
    
    Select Case mint�༭״̬
        Case 1, 5
            Txt������ = gstrUserName
            Txt�������� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
'            Txt�޸��� = gstrUserName
'            Txt�޸����� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 4, 6, 7 '2���޸ģ�4���鿴6�����ܣ����պ��¼���յǼ��ˣ�����ȡ������Ľ��գ���7������
            initGrid
            
            Select Case mintUnit
                Case mconint�ۼ۵�λ
                    strUnitQuantity = "D.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
                Case mconint���ﵥλ
                    strUnitQuantity = "B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,(A.ʵ������ / B.�����װ) AS ʵ������,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,B.�����װ as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������/B.�����װ As ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
                Case mconintסԺ��λ
                    strUnitQuantity = "B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,(A.ʵ������ / B.סԺ��װ) AS ʵ������,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,B.סԺ��װ as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������/B.סԺ��װ As ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,(A.ʵ������ / B.ҩ���װ) AS ʵ������,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,B.ҩ���װ as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������/B.ҩ���װ As ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
            End Select
            
            If mint�༭״̬ = 7 Then
                gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                    "     B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ������,A.����, A.ԭ����, A.����,A.����,B.�ӳ���,B.ҩ����� AS ��������," & _
                    "     B.���Ч��,A.Ч��," & strUnitQuantity & _
                    "     A.�ɱ����,0 ���۽��, 0 ���,D.ժҪ,A.�ⷿID,A.�Է�����ID,D.�Ƿ���,B.ҩ������ AS ҩ����������,A.�ϴι�Ӧ��ID,A.��׼�ĺ�,A.��д���� ��ʵ���� " & _
                    "     FROM " & _
                    "         (SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,0 ʵ������,SUM(�ɱ����) AS �ɱ����," & _
                    "          ҩƷID,���,����, ԭ����, ����,Ч��,NVL(����,0) ����,����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0) �ϴι�Ӧ��ID,��׼�ĺ� " & _
                    "          FROM ҩƷ�շ���¼ X " & _
                    "          WHERE NO=[1] AND ����=6 AND ���ϵ��=-1 " & _
                    "          GROUP BY ҩƷID,���,����,ԭ����, ����,Ч��,NVL(����,0),����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0),��׼�ĺ�" & _
                    "          HAVING SUM(ʵ������)<>0 ) A," & _
                    "     ҩƷ��� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E, " & _
                    " (Select ���, ժҪ From ҩƷ�շ���¼ " & _
                    "  Where ���� = 6 And NO = [1] And ���ϵ�� = -1 And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) D " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND B.ҩƷID=D.ID And A.��� = D.���) W," & _
                    "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                    " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Z.����(+) " & _
                     " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                    " B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ������,A.����,A.ԭ����,A.����,A.����,B.�ӳ���,B.ҩ����� AS ��������,A.��д���� as ԭʼ����, " & _
                    " B.���Ч��,A.Ч��," & strUnitQuantity & _
                    " A.�ɱ����,A.���۽��, A.���, " & strUnitQuantity_Stock & _
                    " ,A.ժҪ,������,��������,�޸���,�޸�����,�����,�������,A.�ⷿID,A.�Է�����ID,D.�Ƿ���,B.ҩ������ AS ҩ����������,NVL(A.��ҩ��λID,0) �ϴι�Ӧ��ID,A.��׼�ĺ�,nvl(A.����,0) As ���췽ʽ  " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ D, " & _
                    "     (SELECT ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND b.ҩƷID=D.ID " & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND A.���� = 6 AND A.���ϵ��=-1 AND A.NO = [1] AND A.��¼״̬ =[3] " & _
                    " AND A.ҩƷID=Z.ҩƷID(+) AND NVL(A.����,0)=Z.����(+) " & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, lngStockid, mint��¼״̬)
        
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 4 Or mint�༭״̬ = 6 Then
                mintApplyType = rsInitCard!���췽ʽ
            End If
            mshBill.Active = IIf(mintApplyType = 0, True, IsHavePrivs(mstrPrivs, "�Զ�����ʱ�޸�ҩƷ����"))
            
            If mint�༭״̬ = 7 Then '7������
                Txt������ = gstrUserName
                Txt�������� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
                Txt�޸��� = gstrUserName
                Txt�޸����� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
                Txt����� = gstrUserName
                Txt������� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
            Else '2���޸ģ�4���鿴6�����ܣ����պ��¼���յǼ��ˣ�����ȡ������Ľ��գ�
                Txt������ = rsInitCard!������
                Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                
                Txt�޸��� = IIf(IsNull(rsInitCard!�޸���), "", rsInitCard!�޸���)
                Txt�޸����� = IIf(IsNull(rsInitCard!�޸�����), "", Format(rsInitCard!�޸�����, "yyyy-mm-dd hh:mm:ss"))
                
                Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            End If
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint�༭״̬ = 2 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    'IntRow = rsInitCard!���
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
                    .TextMatrix(intRow, mconIntCol���) = Nvl(rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!ԭ����), "", rsInitCard!ԭ����)
                    .TextMatrix(intRow, mconIntCol��λ) = Nvl(rsInitCard!��λ)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                                
                    .TextMatrix(intRow, mconIntCol��д����) = zlStr.FormatEx(rsInitCard!��д����, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntColʵ������) = zlStr.FormatEx(rsInitCard!ʵ������, intNumberDigit, , True)
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconIntColԭʼ����) = zlStr.FormatEx(rsInitCard!ԭʼ����, intNumberDigit, , True)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(rsInitCard!�ɱ���, intCostDigit, , True)
                    
                    .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(IIf(mint�༭״̬ = 7, 0, rsInitCard!�ɱ����), intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPricedigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                    
                    .TextMatrix(intRow, mconIntCol���Ч��) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntcol�ӳ���) = Nvl(rsInitCard!�ӳ���, 0) / 100
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntColʵ�ʲ��) = IIf(IsNull(rsInitCard!ʵ�ʲ��), "0", rsInitCard!ʵ�ʲ��)
                    .TextMatrix(intRow, mconIntColʵ�ʽ��) = IIf(IsNull(rsInitCard!ʵ�ʽ��), "0", rsInitCard!ʵ�ʽ��)
                    .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID) = Nvl(rsInitCard!�ϴι�Ӧ��ID)
                                        
                    If mint�༭״̬ = 7 Then
                        .TextMatrix(intRow, mconintCol��ʵ����) = rsInitCard!��ʵ����
                    End If
                        
                    
                    If mint�༭״̬ = 2 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!ҩƷID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        mcolUsedCount.Add Array(CStr(rsInitCard!ҩƷID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)), CStr(numUseAbleCount + IIf(IsNull(rsInitCard!��д����), "0", rsInitCard!��д����))), CStr(rsInitCard!ҩƷID) & CStr(IIf(IsNull(rsInitCard!����), "0", rsInitCard!����))
                        
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    
    Call get�������
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    If mint�༭״̬ <> 7 Then Call CheckNumber
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
        .TextMatrix(0, mconintcol��ǰ���) = "��ǰ���"
        .TextMatrix(0, mconintcol�Է����) = "�Է����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol��д����) = IIf(mint�༭״̬ = 7, "����", "��д����")
        .TextMatrix(0, mconIntColʵ������) = IIf(mint�༭״̬ = 7, "��������", "ʵ������")
        
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol���Ч��) = "���Ч��"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconIntcol�ӳ���) = "�ӳ���"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ϴι�Ӧ��ID) = "�ϴι�Ӧ��ID"
        .TextMatrix(0, mconintCol��ʵ����) = "��ʵ����"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol����ҩƷ) = "����ҩƷ"
        .TextMatrix(0, mconIntColԭʼ����) = "ԭʼ����"
         
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntColҩ��) = 2200
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntCol��λ) = 400
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol��д����) = 1100
        .ColWidth(mconIntColʵ������) = 1100
        .ColWidth(mconIntCol�ɹ���) = 1000
        .ColWidth(mconIntCol�ɹ����) = 900
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = 800
        
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol���Ч��) = 0
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        .ColWidth(mconIntcol�ӳ���) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntCol�ϴι�Ӧ��ID) = 0
        .ColWidth(mconintCol��ʵ����) = 0
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntCol����ҩƷ) = 0
        .ColWidth(mconIntColԭʼ����) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol��������) = 0
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
        .ColData(mconIntCol����ҩƷ) = 5
        .ColData(mconIntColԭʼ����) = 5
        
        '��״̬Ϊ���ܱ༭
        .ColData(mconintcol��ǰ���) = 5
        .ColData(mconintcol�Է����) = 5
        
'        '��������Ϊ�༭״̬���������޸ģ�ʱ�ɼ�
'        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
            .ColWidth(mconintcol��ǰ���) = 1100
            '��û����ʾ�Է����Ȩ�޵�ʱ������ʾ�Է����
            If IsHavePrivs(mstrPrivs, "��ʾ�Է����") Then
                .ColWidth(mconintcol�Է����) = 1100
            Else
                .ColWidth(mconintcol�Է����) = 0
            End If
'        Else
'            .ColWidth(mconintcol��ǰ���) = 0
'            .ColWidth(mconintcol�Է����) = 0
'        End If
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
            
            cboStock.Enabled = True
            txtժҪ.Enabled = True
            
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol��д����) = 4
            .ColData(mconIntColʵ������) = 5
        ElseIf mint�༭״̬ = 4 Or mint�༭״̬ = 6 Then
            cboStock.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = IIf(mint�༭״̬ <> 6, 4, 5)
            .ColData(mconIntColҩ��) = 0
        ElseIf mint�༭״̬ = 7 Then
            cboStock.Enabled = False
            txtժҪ.Enabled = True
            
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 4
            .ColData(mconIntColҩ��) = 0
        End If
        
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
        
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntCol���Ч��) = 5
        .ColData(mconIntColʵ�ʲ��) = 5
        .ColData(mconIntColʵ�ʽ��) = 5
        .ColData(mconIntcol�ӳ���) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconintCol��ʵ����) = 5
        .ColData(mconIntCol�ϴι�Ӧ��ID) = 5
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColԭ����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconintcol��ǰ���) = flexAlignRightCenter
        .ColAlignment(mconintcol�Է����) = flexAlignRightCenter
        .ColAlignment(mconIntCol��д����) = flexAlignRightCenter
        .ColAlignment(mconIntColʵ������) = flexAlignRightCenter
        .ColAlignment(mconintCol��ʵ����) = flexAlignRightCenter
        
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mconIntColҩ��) = 0
    End With
    txtժҪ.MaxLength = GetLength("ҩƷ�շ���¼", "ժҪ")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - IIf(Me.cmdȫ������.Visible, 350, 0) - .Top - 100 - CmdCancel.Height - 200
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
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
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
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
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
    
    With chkExportPlan
        .Top = lblCode.Top
    End With
    
    With cmdȫ��
        If .Visible = True Then
            .Left = Me.CmdSave.Left
            .Top = Me.CmdSave.Top
        End If
    End With
    
    With cmdȫ������
        If .Visible = True Then
            .Left = Me.CmdCancel.Left
            .Top = Me.CmdCancel.Top
        End If
    End With
    
    If mint�༭״̬ = 5 And Me.cmdȫ��.Visible = True Then
        With Me.CmdSave
            .Left = Me.CmdSave.Left
            .Top = Me.CmdSave.Height + Me.CmdSave.Top + 100
        End With
    
        With Me.CmdCancel
            .Left = Me.CmdCancel.Left
            .Top = Me.CmdCancel.Height + Me.CmdCancel.Top + 100
        End With
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mintApplyType = 0
    mstrEndTime = ""
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�������", "ͬ�����ɼƻ���", Me.chkExportPlan.Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�������", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If mblnChange = False Or mint�༭״̬ = 4 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        mblnUpdate = False
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
    mblnUpdate = False
End Sub

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
Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "3467", mint�༭״̬) <> 0 Then
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
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "ҩƷ�������", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln�����)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , mbln�����, False, IsHavePrivs(mstrPrivs, "��ʾ�Է����"), 0, , mstrPrivs)
    If RecReturn.RecordCount > 0 Then
        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
    End If
    mshBill.CmdEnable = True
    
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
        For i = 1 To RecReturn.RecordCount
            intCurRow = mshBill.Row
            With mshBill
                .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                SetColValue .Row, RecReturn!ҩƷID, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                    zlStr.Nvl(RecReturn!ҩƷ��Դ), zlStr.Nvl(RecReturn!����ҩ��), _
                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
                    IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                    RecReturn!ҩ�����, _
                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                    IIf(IsNull(RecReturn!�ӳ���), "0", RecReturn!�ӳ��� / 100), _
                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                    IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                    RecReturn!�ϴι�Ӧ��ID, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�), Nvl(RecReturn!ԭ����)
                .Col = mconIntCol��д����
'                .TextMatrix(.Row, mconIntCol����ҩƷ) = True
                
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

Private Sub mshBill_DblClick(Cancel As Boolean)
    If Me.mshBill.Row <> Me.mshBill.rows - 1 Then
        If Me.mshBill.Col = mconIntCol�������� And Me.mshBill.Row <> 0 Then
            If Val(Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����)) = 0 Then
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����) = Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��������)
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntColʵ������) = Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��������)
            Else
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����) = 0
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntColʵ������) = 0
            End If
        End If
        
        If Val(Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����)) <> 0 Then
            Call GetPrice(Me.mshBill.Row)
        Else
             With Me.mshBill
                .TextMatrix(Me.mshBill.Row, mconIntCol�ۼ۽��) = 0
                .TextMatrix(Me.mshBill.Row, mconintCol���) = 0
                .TextMatrix(Me.mshBill.Row, mconIntCol�ɹ���) = 0
                .TextMatrix(Me.mshBill.Row, mconIntCol�ɹ����) = 0
            End With
        End If
        
        Call ��ʾ�ϼƽ��
        If mint�༭״̬ <> 7 Then Call CheckNumber
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    With mshBill
        If .Col <> mconIntCol���� Then
            mshBill.Text = UCase(curText)
            mshBill.SelStart = Len(mshBill.Text)
        End If
    End With
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strKey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol��д���� Or .Col = mconIntColʵ������ Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mconIntCol��д����, mconIntColʵ������
                    intDigit = mintNumberDigit
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
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntColҩ��
                .TxtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
'                Call ��ʾ�����
                
            Case mconIntCol����
                .TxtCheck = True
                .TextMask = "1234567890"
                .MaxLength = 8
            
            Case mconIntColЧ��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol����) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) And .TextMatrix(.Row, mconIntCol���Ч��) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol���Ч��), "||")(0) <> 0 Then
                            strxq = .TextMatrix(.Row, mconIntCol����)
                            strxq = TranNumToDate(strxq)
                            If strxq = "" Then Exit Sub
                            
                            .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol���Ч��), "||")(0), strxq), "yyyy-mm-dd")
                            If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                                '����Ϊ��Ч��
                                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mconIntCol��д����, mconIntColʵ������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
'                Call ��ʾ�����
                
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
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
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
                    End If

                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, "ҩƷ�������", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln�����)
                    End If
                    Set RecReturn = frmSelector.ShowMe(Me, 1, 2, strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , mbln�����, False, IsHavePrivs(mstrPrivs, "��ʾ�Է����"), 0, , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                    End If
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                            If SetColValue(.Row, RecReturn!ҩƷID, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                    zlStr.Nvl(RecReturn!ҩƷ��Դ), zlStr.Nvl(RecReturn!����ҩ��), _
                                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
                                    IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                                    RecReturn!ҩ�����, _
                                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                    IIf(IsNull(RecReturn!�ӳ���), "0", RecReturn!�ӳ��� / 100), _
                                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                                    IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                                    RecReturn!�ϴι�Ӧ��ID, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�), Nvl(RecReturn!ԭ����)) = False Then
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
                        .Col = mconIntCol��д����
                    Else
                        .TextMatrix(.Row, mconIntCol����ҩƷ) = True
                        If Val(.TextMatrix(.Row, 0)) = 0 Then
                            .Text = .TextMatrix(.Row, .Col)
                            Cancel = True
                        Else
                            .Text = .TextMatrix(.Row, .Col)
                        End If
                    End If
'                    Call ��ʾ�����
                End If
            Case mconIntCol����
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol����) = ""
                    End If
                    If .ColData(mconIntColЧ��) = 2 Then
                        .Col = mconIntColЧ��
                    Else
                        .Col = mconIntCol��д����
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
                            MsgBox "�Բ���Ч�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�Բ���Ч�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
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
            
            Case mconIntCol��д����, mconIntColʵ������
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
                    If Val(strKey) >= 10 ^ 11 - 1 Then
                        MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mint�༭״̬ = 7 Then
                       If Val(strKey) >= 0 Then
                            If Val(strKey) > Val(.TextMatrix(.Row, mconIntCol��д����)) Then
                                MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        Else
                            If Val(strKey) < Val(.TextMatrix(.Row, mconIntCol��д����)) Then
                                MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    '10.35.40,�������γ���ʱ��������������򲻼��(�ƿ�����ʱ�ټ��)
                    If ((mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5) And mint�����γ��� = 1) Or mint�༭״̬ = 7 Then
                        If Not CheckUsableNum(IIf(mint�༭״̬ = 7, mlngStockID, cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), _
                            strKey, Val(.TextMatrix(.Row, mconIntCol����ϵ��)), Trim(txtNo.Caption), _
                            6, IIf(mint�༭״̬ = 7, mint��������ⷿ, mint�����), mintNumberDigit, IIf(mint�༭״̬ = 7, Val(.TextMatrix(.Row, mconIntCol���)), 0), _
                            IIf(mint�༭״̬ = 7, Get����д����(.Row, strKey), 0)) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '�ɱ��۵Ĺ�ʽ��     ������=����*�ۼ�
                    '                  ������=������*��ʵ�ʲ��/ʵ�ʽ�
                    '                  if ʵ�ʽ��=0 then  ������=������*ָ�������
                    '                  ���ۣ��ɱ��ۣ�=ֱ�Ӵӿ�����ȡƽ���ɱ���
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'ʵ�ʽ��=0������£����ο��Ǵӡ�����¼���ϴβɹ��ۡ�����ҩƷ���ĳɱ��ۡ�����ָ������ʡ�ȡֵ
                    
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = strKey
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                    End If
                    
                    If strKey <> 0 Then
'                        .TextMatrix(.Row, mconIntCol�ɹ���) = FormatEx((Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - .TextMatrix(.Row, mconintCol���)) / strkey, mintCostDigit)
                         If mint�༭״̬ <> 7 Then .TextMatrix(.Row, mconIntCol�ɹ���) = zlStr.FormatEx(Get�ɱ���(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mconIntCol����))) * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), mintCostDigit, , True)
                    End If
                    
                    .TextMatrix(.Row, mconIntCol�ɹ����) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * strKey, mintMoneyDigit, , True)
                    
'                    If mint�༭״̬ = 7 Then
                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit, , True)
'                    Else
'                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Get������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), Val(.TextMatrix(.Row, mconIntColʵ�ʽ��)), Val(.TextMatrix(.Row, mconIntColʵ�ʲ��)), Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)), Val(strkey) * Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��))), mintMoneyDigit)
'                    End If
                    
                    If .Col = mconIntCol��д���� Then
                        .TextMatrix(.Row, mconIntColʵ������) = strKey
                    End If
                    
                    
                End If
                
                ��ʾ�ϼƽ��
                If mint�༭״̬ <> 7 Then Call CheckNumber(1)
        End Select
    End With
End Sub

Private Function Get����д����(ByVal intRow As Integer, ByVal dbl��д���� As Double)
    'ȡ��������ѡ����ҩƷ������д��������������Ƿ����ģ���������������ε�����д������
    '���㷽���������еĸ�ҩƷ��д����+�����ѡ���е���д����
    Dim n As Integer
    Dim dblSum As Double
    
    With mshBill
    
        For n = 1 To .rows - 1
            If n <> intRow And Val(.TextMatrix(n, 0)) = Val(.TextMatrix(intRow, 0)) Then
                '����ѡ���е�ͬһ��ҩƷ
                dblSum = dblSum + Val(.TextMatrix(n, mconIntColʵ������))
            End If
        Next
        
        dblSum = dblSum + dbl��д����
    End With
    
    Get����д���� = dblSum
End Function
Private Sub GetPrice(ByVal intRow As Integer)
    With Me.mshBill
        .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) * Me.mshBill.TextMatrix(intRow, mconIntCol��д����), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconIntCol�ɹ���) = zlStr.FormatEx(Get�ɱ���(Val(.TextMatrix(intRow, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol����))) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintCostDigit, , True)
        .TextMatrix(intRow, mconIntCol�ɹ����) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) * Val(Me.mshBill.TextMatrix(intRow, mconIntCol��д����)), mintMoneyDigit, , True)
        .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) - .TextMatrix(intRow, mconIntCol�ɹ����), mintMoneyDigit, , True)
    End With
End Sub

'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lngҩƷid As Long, _
    ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, _
    ByVal str����ҩ�� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal int���Ч�� As Integer, ByVal int�������� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal dbl�ӳ��� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal intҩ������ As Integer, ByVal lng�ϴι�Ӧ��ID As Long, ByVal str��׼�ĺ� As String, ByVal strԭ���� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsPrice As New Recordset
    Dim strҩ�� As String
    
    On Error GoTo errHandle
    SetColValue = False
    
    With mshBill
        
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, 0) = lngҩƷid
        
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
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ� * num����ϵ��, mintPriceDigit, , True)
        If int�Ƿ��� = 1 Then
            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Getʱ�����ۼ�(lngҩƷid, cboStock.ItemData(cboStock.ListIndex), lng����, num����ϵ��), mintPriceDigit, , True)
        End If
        .TextMatrix(intRow, mconIntCol��������) = int��������
        .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(num��������, mintNumberDigit, , True)
        .TextMatrix(intRow, mconIntCol���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & intҩ������
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mconIntcol�ӳ���) = dbl�ӳ���
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID) = lng�ϴι�Ӧ��ID
        
        If lng���� > 0 Then
            .TextMatrix(intRow, mconIntCol����) = lng����
        Else
            .TextMatrix(intRow, mconIntCol����) = 0
        End If
        
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        
        Call CheckLapse(strЧ��)
        
        '�Ƿ񳣱�ҩƷ
        Dim rsTmp As ADODB.Recordset
        gstrSQL = "select nvl(�Ƿ񳣱�,0) �Ƿ񳣱� from ҩƷ��� where ҩƷid=[1]"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷid)
        .TextMatrix(intRow, mconIntCol����ҩƷ) = IIf(rsTmp!�Ƿ񳣱� = 1, False, True)
        
        Call get�������(intRow)
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
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol��д����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ������Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol��д����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol��д����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ����д�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol��д����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntColʵ������)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ��ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntColʵ������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɹ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol��д����) = 4, mconIntCol��д����, mconIntColʵ������)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol��д����) = 4, mconIntCol��д����, mconIntColʵ������)
                        Exit Function
                    End If
                            
                    If mint�����γ��� = 1 Then
                        If Not CheckUsableNum(cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(intLop, 0)), Val(mshBill.TextMatrix(intLop, mconIntCol����)), _
                                        Val(mshBill.TextMatrix(intLop, mconIntColʵ������)), Val(.TextMatrix(intLop, mconIntCol����ϵ��)), _
                                        Trim(txtNo.Caption), 6, mint�����, mintNumberDigit) Then
                            mshBill.SetFocus
                            .Row = intLop
                            .Col = mconIntCol��д����
                            Exit Function
                        End If
                    End If
                            
                    '���۹�������Ƿ���ڲ��������۵�ҩƷ
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                        If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                            If CheckPriceAdjust(Val(.TextMatrix(intLop, 0)), cboStock.ItemData(cboStock.ListIndex), IIf(mint�����γ��� = 0, -1, Val(.TextMatrix(intLop, mconIntCol����)))) = False Then
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

Private Function SaveCard() As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngEnterStockID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblRealQuantity As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strModifier As String
    Dim datModifyDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim lng�ϴι�Ӧ��ID As Long
    Dim str��׼�ĺ� As String
    Dim int��� As Integer
    
    Dim intRow As Integer
    Dim arrSql As Variant
    'ҩƷ�ɹ��ƻ�
    Dim strSQLDrugPlan As String
    Dim arrSQLDrugPlanDetail As Variant
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim arrSum As Variant
    
    '�Զ��ֽ������¼ʱʹ��
    Dim blnAuto As Boolean              '�Ƿ���Ҫ�Զ��ֽ�
    Dim rsStock As New ADODB.Recordset
    
    Dim strCheckString As String
    Dim n As Integer, intPlanSN As Integer
    Dim rsSpec As ADODB.Recordset   '������ݼ�
    Dim dbl�ͻ����� As Double
    
    SaveCard = False
    arrSql = Array()
    arrSQLDrugPlanDetail = Array()
    arrSum = Array()
    
    On Error GoTo errHandle
    
    With mshBill
        chrNo = Trim(txtNo)
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        If chrNo = "" Then chrNo = Sys.GetNextNo(26, lngStockid)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        lngEnterStockID = mlngStockID
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strAssessor = Txt�����
        
        ID_IN = Sys.NextId("ҩƷ�ɹ��ƻ�")
        NO_IN = Sys.GetNextNo(32, mlngStockID)
        
        If mint�༭״̬ = 2 Then        '�޸�
            strCheckString = CheckBill(chrNo)
            If strCheckString <> "" Then
                MsgBox strCheckString, vbInformation, gstrSysName
                Exit Function
            End If
        
            gstrSQL = "zl_ҩƷ�ƿ�_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
            
            strBooker = Txt������
            datBookDate = Format(Txt��������, "yyyy-mm-dd hh:mm:ss")
            strModifier = gstrUserName
            datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If zlStr.Nvl(.TextMatrix(intRow, mconIntCol��д����), 0) <> 0 Then
                int��� = intRow 'int��� + 1
                If .TextMatrix(intRow, 0) <> "" Then
                    '�����ǰ����ҩƷ�������Զ�ȡ�������ε�ҩƷ��������������¼
                    lngDrugID = .TextMatrix(intRow, 0)
                    strProducingArea = .TextMatrix(intRow, mconIntCol����)
                    strOldProducingArea = .TextMatrix(intRow, mconIntColԭ����)
                    strBatchNo = .TextMatrix(intRow, mconIntCol����)
                    lngBatchID = Val(.TextMatrix(intRow, mconIntCol����))
                    datTimeLimit = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                        '����ΪʧЧ��������
                        datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                    End If
                    
                    dblQuantity = .TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol����ϵ��)
                    dblRealQuantity = .TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��)
                    
'                    dblPurchasePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɹ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                    
                    dblPurchasePrice = Get�ɱ���(lngDrugID, lngStockid, lngBatchID)
                    
                    dblPurchaseMoney = Val(zlStr.FormatEx(Val(FormatEx(dblPurchasePrice * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintCostDigit)) * Val(.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit, , True)) ' .TextMatrix(intRow, mconIntCol�ɹ����)
                    
'                    dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                    
                    dblSalePrice = Get���ۼ�(Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1, lngDrugID, lngStockid, lngBatchID)
                    
                    dblSaleMoney = Val(zlStr.FormatEx(Val(FormatEx(dblSalePrice * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintPriceDigit)) * Val(.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit, , True))  ' .TextMatrix(intRow, mconIntCol�ۼ۽��)
                    dblMistakePrice = Val(zlStr.FormatEx(dblSaleMoney - dblPurchaseMoney, mintMoneyDigit, , True)) ' Val(.TextMatrix(intRow, mconintCol���))
                    
                    lng�ϴι�Ӧ��ID = .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID)
                    str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                    
'                    If Val(.TextMatrix(intRow, mconIntCol���)) = 0 Then
'                        lngSerial = 2 * int��� - 1  '����������ʽΪ��2n-1;�������Ϊż��
'                    Else
'                        lngSerial = Val(.TextMatrix(intRow, mconIntCol���))
'                    End If
                    lngSerial = 2 * int��� - 1
                    
                    gstrSQL = "zl_ҩƷ����_INSERT("
                    'NO
                    gstrSQL = gstrSQL & "'" & chrNo & "'"
                    '���
                    gstrSQL = gstrSQL & "," & lngSerial
                    '�ⷿID
                    gstrSQL = gstrSQL & "," & lngStockid
                    '�Է�����ID
                    gstrSQL = gstrSQL & "," & lngEnterStockID
                    'ҩƷID
                    gstrSQL = gstrSQL & "," & lngDrugID
                    '����
                    gstrSQL = gstrSQL & "," & lngBatchID
                    '��д����
                    gstrSQL = gstrSQL & "," & dblQuantity
                    'ʵ������
                    gstrSQL = gstrSQL & "," & dblRealQuantity
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
                    gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                    '����
                    gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                    'Ч��
                    gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & datTimeLimit & "','yyyy-mm-dd')")
                    'ժҪ
                    gstrSQL = gstrSQL & ",'" & strBrief & "'"
                    '��������
                    gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                    '��Ӧ��ID
                    gstrSQL = gstrSQL & "," & IIf(lng�ϴι�Ӧ��ID = 0, "NULL", lng�ϴι�Ӧ��ID)
                    '��׼�ĺ�
                    gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                    '���췽ʽ
                    gstrSQL = gstrSQL & "," & mintApplyType
                    '����ʱ��
                    gstrSQL = gstrSQL & ",'" & mstrEndTime & "'"
                    'ԭ����
                    gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                    '�޸���
                    gstrSQL = gstrSQL & ",'" & strModifier & "'"
                    '�޸�����
                    gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                    gstrSQL = gstrSQL & ")"
    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = CStr(lngDrugID) & ";" & gstrSQL
                    
                    'ҩƷ�ɹ��ƻ�����
                    If chkExportPlan.Value = 1 And chkExportPlan.Visible Then
                        If .TextMatrix(intRow, mconIntCol����ҩƷ) = "" Then .TextMatrix(intRow, mconIntCol����ҩƷ) = True
                        If .TextMatrix(intRow, mconIntCol����ҩƷ) = False Then
                            gstrSQL = "Select �ͻ���λ,�ͻ���װ From ҩƷ��� Where ҩƷid = [1]"
                            Set rsSpec = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�ͻ���λ", lngDrugID)
                            If IsNull(rsSpec!�ͻ���λ) = False Then
                                dbl�ͻ����� = zlStr.FormatEx(dblRealQuantity / rsSpec!�ͻ���װ, 1, , True)
                            End If
                            '��������ͬҩƷID���ϲ�����
                            If CheckRepeatDrugID(recSort, n, lngDrugID) Then
                                '�ϲ�����
                                SumQuantity arrSum, lngDrugID, dblQuantity
                            Else
                                intPlanSN = intPlanSN + 1
                                gstrSQL = "zl_ҩƷ�ƻ�����α�_INSERT(" & _
                                          ID_IN & "," & _
                                          lngDrugID & "," & _
                                          intPlanSN & "," & _
                                          GetQuantity(arrSum, lngDrugID, dblQuantity) & "," & _
                                          dblPurchasePrice & "," & _
                                          dblPurchaseMoney & "," & _
                                          "null,null,0," & _
                                          IIf(lng�ϴι�Ӧ��ID <= 0, "null", "'" & GetProvider(lng�ϴι�Ӧ��ID) & "'") & "," & _
                                          IIf(strProducingArea = "", "null", "'" & strProducingArea & "'") & "," & _
                                          "null," & _
                                          dblSalePrice & "," & _
                                          dblSaleMoney & "," & _
                                          "null,null," & _
                                          dbl�ͻ����� & ")"
                                
                                ReDim Preserve arrSQLDrugPlanDetail(UBound(arrSQLDrugPlanDetail) + 1)
                                arrSQLDrugPlanDetail(UBound(arrSQLDrugPlanDetail)) = gstrSQL & ";"
                            End If
                        End If
                    End If
                End If
            End If
            recSort.MoveNext
        Next
        
        'ҩƷ�ɹ��ƻ�
        If chkExportPlan.Value = 1 And chkExportPlan.Visible Then
            strSQLDrugPlan = "zl_ҩƷ�ƻ���������_INSERT(" & _
                             ID_IN & ",'" & _
                             NO_IN & "'," & _
                             "0," & _
                             "null," & _
                             lngStockid & "," & _
                             lngEnterStockID & "," & _
                             "0,'" & _
                             strBooker & "'," & _
                             "to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                             "��ҩƷ�깺�����Զ����ɡ�')"
        End If
         
        If Not ExecuteSql(arrSql, strSQLDrugPlan, arrSQLDrugPlanDetail, MStrCaption) Then Exit Function
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function SaveCheck() As Boolean
    Dim rstemp As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim lng�Է�����id As Long
    Dim str����� As String
    
    Dim lngҩƷid As Long
    Dim str���� As String
    Dim lng������ As Long
    Dim num��д���� As Double
    Dim numʵ������ As Double
    Dim num�ɱ��� As Double
    Dim num�ɱ���� As Double
    Dim dbl�ۼ� As Double
    Dim num���۽�� As Double
    Dim num��� As Double
    Dim lng�����id As Long
    Dim lng�����id As Long
    Dim str���� As String
    Dim datЧ�� As String
    Dim dat������� As String
    Dim int���к� As Integer
    Dim lng�ϴι�Ӧ��ID As Long
    Dim str��׼�ĺ� As String
    Dim strҩƷ As String

    Dim arrSql As Variant
    Dim n As Integer
    
    arrSql = Array()
    mblnSave = False
    SaveCheck = False
    On Error GoTo errHandle
    '���õ����Ƿ��ڽ���༭����󣬱���������Ա�޸�
    mstrTime_End = GetBillInfo(6, mstr���ݺ�)
    If mstrTime_End = "" Then
        MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    If mstrTime_End > mstrTime_Start Then
        MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���õ����Ƿ���������
    gstrSQL = " Select ��ҩ���� From ҩƷ�շ���¼ " & _
            " Where ����=6 And NO=[1] And Rownum<2"
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���õ����Ƿ���������]", Me.txtNo.Tag)
    
    If IsNull(rstemp!��ҩ����) Then
        MsgBox "�õ��ݱ���������Աȡ�����ͣ���������գ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng�Է�����id = mlngStockID
    str����� = gstrUserName
    strNo = txtNo.Tag
    
    gstrSQL = "SELECT b.ϵ��,b.id AS ���id " _
            & " FROM ҩƷ�������� a, ҩƷ������ b " _
            & "Where a.���id = b.ID " _
            & "  AND a.���� = 6 "
    
    Call SQLTest(App.Title, "ҩƷ�ƿ����", gstrSQL)
    If rstemp.State = 1 Then rstemp.Close
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "SaveCheck")
    Call SQLTest
    
    If rstemp.EOF Then
        MsgBox "�Բ���ҩƷ������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rstemp.RecordCount < 2 Then
        MsgBox "�Բ���ҩƷ������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rstemp.MoveFirst
    Do While Not rstemp.EOF
        If rstemp!ϵ�� = 1 Then
            lng�����id = rstemp!���id
        Else
            lng�����id = rstemp!���id
        End If
        rstemp.MoveNext
    Loop
    rstemp.Close
    
'    If mblnUpdate = False Then
'        If Not ��鵥��(6, txtNO.Tag, False, True) Then
'            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
'            Call RefreshBill
'            mblnUpdate = True
'            Exit Function
'        End If
'    End If

    If ���۸� Then
        mblnUpdate = True
        mblnChange = True
        Exit Function
    End If
    
    '�����
    strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, mconIntColʵ������, mconIntCol����ϵ��, 1, 1, mconIntColԭʼ����)
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
    
    '���۹�������Ƿ���ڲ��������۵�ҩƷ
    For n = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(n, 0) <> "" And gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
            If IsPriceAdjustMod(Val(mshBill.TextMatrix(n, 0))) = True Then
                If CheckPriceAdjust(Val(mshBill.TextMatrix(n, 0)), cboStock.ItemData(cboStock.ListIndex), Val(mshBill.TextMatrix(n, mconIntCol����))) = False Then
                    MsgBox "��" & n & "��ҩƷ���������۹���������¼���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
                    mshBill.SetFocus
                    mshBill.Row = n
                    mshBill.MsfObj.TopRow = n
                    Exit Function
                End If
            End If
        End If
    Next
    
    dat������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        On Error GoTo errHandle
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngҩƷid = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng������ = .TextMatrix(intRow, mconIntCol����)
                
                If Val(.TextMatrix(intRow, mconIntCol��д����)) = Val(.TextMatrix(intRow, mconIntColʵ������)) Then
                    num��д���� = Val(.TextMatrix(intRow, mconIntColԭʼ����))
                    numʵ������ = Val(.TextMatrix(intRow, mconIntColԭʼ����))
                Else
                    num��д���� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol��д����)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                    numʵ������ = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntColʵ������)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                End If
                
'                num�ɱ��� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���, , True)
                num�ɱ��� = Get�ɱ���(lngҩƷid, lng�ⷿID, lng������)
                num�ɱ���� = Val(.TextMatrix(intRow, mconIntCol�ɹ����))
'                dbl�ۼ� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�, , True)
                dbl�ۼ� = Get���ۼ�(Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1, lngҩƷid, lng�ⷿID, lng������)
                num���۽�� = Val(.TextMatrix(intRow, mconIntCol�ۼ۽��))
                num��� = Val(.TextMatrix(intRow, mconintCol���))
                str���� = .TextMatrix(intRow, mconIntCol����)
                datЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datЧ�� <> "" Then
                    '����ΪʧЧ��������
                    datЧ�� = Format(DateAdd("D", 1, datЧ��), "yyyy-mm-dd")
                End If
                                
                int���к� = Val(.TextMatrix(intRow, mconIntCol���))
                lng�ϴι�Ӧ��ID = .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID)
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                        
                gstrSQL = "zl_ҩƷ�ƿ�_Verify("
                '���
                gstrSQL = gstrSQL & int���к�
                '�ⷿID
                gstrSQL = gstrSQL & "," & lng�ⷿID
                '�Է�����ID
                gstrSQL = gstrSQL & "," & lng�Է�����id
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngҩƷid
                '����
                gstrSQL = gstrSQL & ",'" & str���� & "'"
                '������
                gstrSQL = gstrSQL & "," & lng������
                'ʵ������
                gstrSQL = gstrSQL & "," & numʵ������
                '�ɱ���
                gstrSQL = gstrSQL & "," & num�ɱ���
                '�ɱ����
                gstrSQL = gstrSQL & "," & num�ɱ����
                '���۽��
                gstrSQL = gstrSQL & "," & num���۽��
                '���
                gstrSQL = gstrSQL & "," & num���
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '�����
                gstrSQL = gstrSQL & ",'" & str����� & "'"
                '����
                gstrSQL = gstrSQL & ",'" & str���� & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '�������
                gstrSQL = gstrSQL & ",to_date('" & dat������� & "','yyyy-mm-dd HH24:MI:SS')"
                '��Ӧ��ID
                gstrSQL = gstrSQL & "," & IIf(lng�ϴι�Ӧ��ID = 0, "NULL", lng�ϴι�Ӧ��ID)
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '���ۼ�
                gstrSQL = gstrSQL & "," & dbl�ۼ�
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngҩƷid) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
'    gcnOracle.BeginTrans
    If Not ExecuteSql(arrSql, "", "", MStrCaption) Then
'        gcnOracle.RollbackTrans
        Exit Function
    End If
'    gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
'    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lngҩƷid As Long
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsPrice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    Dim intCostDigit As Integer
    Dim intPricedigit As Integer
            
    On Error GoTo errHandle
    intPricedigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
        
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 6 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & intPricedigit & ") <> Round(b.�ּ�, " & intPricedigit & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0 " & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 1 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = 6 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & intPricedigit & ") <> Round(decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�), " & intPricedigit & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 2 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where a.���� = 6 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & intCostDigit & ")<>round(decode(x.�ּ�,null,b.ƽ���ɱ���,x.�ּ�)," & intCostDigit & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " AND a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) " & _
            " Order By ����, ҩƷid, ���"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lngҩƷid = Val(mshBill.TextMatrix(lngRow, 0))
        Dbl���� = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))
        dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ɹ���))
        dbl���ۼ� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ�))
        dbl�ɱ���� = dbl�ɱ��� * Dbl����
        dbl���۽�� = dbl���ۼ� * Dbl����
        dbl��� = dbl���۽�� - dbl�ɱ����
                
        If lngҩƷid <> 0 Then
            rsPrice.Filter = "����='�ۼ�' And ҩƷID=" & lngҩƷid & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl���ۼ� = Val(zlStr.FormatEx(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intPricedigit, , True))
                dbl���۽�� = Val(zlStr.FormatEx(dbl���ۼ� * Dbl����, mintMoneyDigit, , True))
                dbl��� = Val(zlStr.FormatEx(dbl���۽�� - dbl�ɱ����, mintMoneyDigit, , True))
            End If
            
            rsPrice.Filter = "����='�ɱ���' And ҩƷID=" & lngҩƷid & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(zlStr.FormatEx(dbl���ۼ� * Dbl����, mintMoneyDigit, , True))
                dbl�ɱ��� = Val(zlStr.FormatEx(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intCostDigit, , True))
                dbl�ɱ���� = Val(zlStr.FormatEx(dbl�ɱ��� * Dbl����, mintMoneyDigit, , True))
                dbl��� = Val(zlStr.FormatEx(dbl���۽�� - dbl�ɱ����, mintMoneyDigit, , True))
            End If
            
            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, intPricedigit, , True)
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
    Dim rstemp As New ADODB.Recordset
    Dim n As Integer
    Dim ժҪ_IN As String
    Dim strҩƷid As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim strҩƷ As String
    
    arrSql = Array()
    SaveStrike = False
    
    With mshBill
        '����������������С����
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntColʵ������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mconIntCol��д����)), Val(.TextMatrix(intRow, mconIntColʵ������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '������������������������
            If mint�༭״̬ = 7 And mint����ʽ = 1 Then
                If Not CheckUsableNum(mlngStockID, Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol����)), _
                    Val(.TextMatrix(intRow, mconIntColʵ������)), Val(.TextMatrix(intRow, mconIntCol����ϵ��)), Trim(txtNo.Caption), _
                    6, mint��������ⷿ, mintNumberDigit, Val(.TextMatrix(intRow, mconIntCol���)), _
                    Get����д����(intRow, Val(.TextMatrix(intRow, mconIntColʵ������)))) Then
                    Exit Function
                End If
            End If
        Next
        
        '��ͨ�������ʵ������
        If mint�༭״̬ = 7 And mint����ʽ = 0 Then
            strҩƷ = CheckNumStock(mshBill, mlngStockID, 0, mconIntCol����, mconIntColʵ������, mconIntCol����ϵ��, 2, 0, mconintCol��ʵ����)
            If strҩƷ <> "" Then
                If mint��������ⷿ = 1 Then '��������
                    If MsgBox("ҩƷ��" & strҩƷ & "����治�㣬�Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf mint��������ⷿ = 2 Then '�����ֹ
                    MsgBox "ҩƷ��" & strҩƷ & "����治�㣬������ˣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        NO_IN = Trim(txtNo.Tag)
        ������_IN = gstrUserName
        ��������_IN = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ԭ��¼״̬_IN = mint��¼״̬
        ժҪ_IN = Trim(txtժҪ.Text)
        
        On Error GoTo errHandle
        
        �д�_IN = 0
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntColʵ������)) <> 0 Then
                �д�_IN = �д�_IN + 1
                
                ҩƷID_IN = .TextMatrix(intRow, 0)
                strҩƷid = IIf(strҩƷid = "", "", strҩƷid & ",") & ҩƷID_IN
                If .TextMatrix(intRow, mconIntColʵ������) = .TextMatrix(intRow, mconIntCol��д����) Then
                    ��������_IN = .TextMatrix(intRow, mconintCol��ʵ����)
                Else
                    ��������_IN = zlStr.FormatEx(.TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                End If
                
                ���_IN = .TextMatrix(intRow, mconIntCol���)
                
                gstrSQL = "ZL_ҩƷ�ƿ�_STRIKE("
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
                'ժҪ
                gstrSQL = gstrSQL & "," & IIf(ժҪ_IN = "", "Null", "'" & ժҪ_IN & "'")
                '������ʽ
                gstrSQL = gstrSQL & "," & mint����ʽ
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
        If strҩƷid <> "" And mint����ʽ = 0 Then
            Call CheckStopMedi(strҩƷid)
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
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim dbl����ⷿ���� As Double
    Dim lng���� As Long
    
    On Error GoTo errHandle
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
        Exit Sub
    Else
        With mshBill
            lng���� = Val(.TextMatrix(.Row, mconIntCol����))
            
            '���տⷿ
            If .TextMatrix(.Row, mconIntColҩ��) = "" Then
                staThis.Panels(2).Text = ""
                Exit Sub
            End If
            If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
            If lng���� > 0 Then
                gstrSQL = " Select ��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as �������� from ҩƷ��� " & _
                          " Where �ⷿid=[1] " & _
                          " And ҩƷid=[2] And ����=1 " & _
                          " And Nvl(����,0)=[3]"
            Else
                gstrSQL = " Select Sum(��������)/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as �������� from ҩƷ��� " & _
                          " Where �ⷿid=[1] " & _
                          " And ҩƷid=[2] And ����=1 "
            End If
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�����ⷿ��������]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
            
            If rsUseCount.EOF Then
                dbl����ⷿ���� = 0
            Else
                dbl����ⷿ���� = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            rsUseCount.Close
            
            If lng���� > 0 And Get��������(mlngStockID, Val(.TextMatrix(.Row, 0))) = 1 Then
                gstrSQL = " Select Sum(��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & ") as �������� from ҩƷ��� where �ⷿid=[1] " & _
                      " And ҩƷid=[2] And ����=1 And Nvl(����,0)=[3] "
            Else
                gstrSQL = " Select Sum(��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & ") as �������� from ҩƷ��� where �ⷿid=[1] " & _
                      " And ҩƷid=[2] And ����=1 "
            End If
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ǰҩ����������]", mlngStockID, Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                dblStock = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            
            Dim blnIs��ʾ�Է���� As Boolean
            Dim str�Է������ As String
            
            blnIs��ʾ�Է���� = IsHavePrivs(mstrPrivs, "��ʾ�Է����")
            str�Է������ = "��" & Me.cboStock.Text & "�����Ϊ[" & zlStr.FormatEx(dbl����ⷿ����, mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol��λ)
            
            staThis.Panels(2).Text = "��ҩƷ" & frmRequestDrugList.cboStock.Text & "�����Ϊ[" & zlStr.FormatEx(dblStock, mintNumberDigit, , True) & "]" & .TextMatrix(.Row, mconIntCol��λ) _
                & IIf(blnIs��ʾ�Է����, str�Է������, "")
        End With
    End If
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

'ת����ֵΪ����
Private Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim StrDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 2000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    StrDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(StrDate) Then Exit Function
    
    StrDate = Format(StrDate, "yyyy-mm-dd")
    TranNumToDate = StrDate
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function ExecuteSql(ByRef arrSql As Variant, ByVal strSQLDrugPlan As String _
    , ByRef arrSQLDrugPlanDetail As Variant, strTitle As String, Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer

    ExecuteSql = False
    If UBound(arrSql) >= 0 Then
        '��SQL���а�ҩƷID��������
        For i = 0 To UBound(arrSql) - 1
            For j = i + 1 To UBound(arrSql)
                If CLng(Split(arrSql(j), ";")(0)) < CLng(Split(arrSql(i), ";")(0)) Then
                    strTmp = CStr(arrSql(j))
                    arrSql(j) = arrSql(i)
                    arrSql(i) = strTmp
                End If
            Next
        Next
        
        'ִ��SQL���
        On Error GoTo errH
        If Not blnǿ�Ʊ��� Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(Split(arrSql(i), ";")(1)), strTitle)
        Next
        'ҩƷ�ɹ��ƻ�
        If Trim(strSQLDrugPlan) <> "" Then
            If UBound(arrSQLDrugPlanDetail) >= 0 Then
                Call zlDataBase.ExecuteProcedure(strSQLDrugPlan, strTitle & "-�ɹ��ƻ�")
                For i = 0 To UBound(arrSQLDrugPlanDetail)
                    Call zlDataBase.ExecuteProcedure(CStr(Split(arrSQLDrugPlanDetail(i), ";")(0)), strTitle & "-�ɹ��ƻ�����")
                Next
            End If
        End If
        
        If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
errH:
    If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'��ӡ����
Private Sub printbill()
    Dim int��λϵ�� As Integer
    
    With mshBill
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
        FrmBillPrint.ShowMe Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), mint��¼״̬, int��λϵ��, 1304, "ҩƷ���쵥", txtNo.Tag
    End With
End Sub


Private Sub get�������(Optional ByVal intRow As Integer = 0)
'''''''''''''''''''''''''''''''''''''
'��ȡ��������ķ���
'''''''''''''''''''''''''''''''''''''
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim dbl����ⷿ���� As Double
    Dim int�����ⷿ�������� As Integer
    Dim int���տⷿ�������� As Integer
    Dim int�������� As Integer '��ȡ�ⷿ�Ĺ������ʣ���ҩ�⻹��ҩ��
    Dim blnIs��ʾ�Է���� As Boolean
    Dim str�Է������ As String
    Dim intStart As Integer, intEnd As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    
    blnIs��ʾ�Է���� = IsHavePrivs(mstrPrivs, "��ʾ�Է����")
    
    If intRow > 0 Then
        intStart = intRow
        intEnd = intRow
    Else
        intStart = 1
        intEnd = mshBill.rows - 1
    End If
    
    With mshBill
        For i = intStart To intEnd
            If .TextMatrix(i, 0) = "" Then Exit Sub

            If blnIs��ʾ�Է���� Then
                If Val(.TextMatrix(i, mconIntCol����)) > 0 Then
                    gstrSQL = " Select Nvl(��������,0)/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ��������, Nvl(ʵ������,0)/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ʵ������ from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 " & _
                              " And Nvl(����,0)=[3] "
                Else
                    If Get��������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(i, 0))) = 1 Then
                        '�������ⷿ�Ƿ�������ͳ������Ϊ0��δ�ֽ�����
                        gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ʵ������ from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 And Nvl(����,0)>0 "
                    Else
                        '�������ⷿ�ǲ������ģ���ͳ���ܵ�����
                        gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ʵ������ from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 "
                    End If
                End If
                Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�����ⷿʵ������]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, mconIntCol����)))
                
                If rsUseCount.EOF Then
                    dbl����ⷿ���� = 0
                Else
                    If mint�༭״̬ = 6 Then
                        '����(���)ʱ��ʾʵ������
                        dbl����ⷿ���� = IIf(IsNull(rsUseCount.Fields(1)), 0, rsUseCount.Fields(1))
                    Else
                        '����״̬ʱ������
                        dbl����ⷿ���� = IIf(mint��ʾ�Է���淽ʽ = 0, IIf(IsNull(rsUseCount.Fields(1)), 0, rsUseCount.Fields(1)), IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0)))
                    End If
                End If
                .TextMatrix(i, mconintcol�Է����) = zlStr.FormatEx(dbl����ⷿ����, mintNumberDigit, , True)
                rsUseCount.Close
            End If
                
            '��ָ��������ʾ
            If Val(.TextMatrix(i, mconIntCol����)) > 0 And Get��������(mlngStockID, Val(.TextMatrix(i, 0))) = 1 Then
                If mint��ǰ��水������ʾ = 0 Then
                    gstrSQL = " Select Nvl(��������,0)/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ��������, Nvl(ʵ������,0)/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ʵ������ from ҩƷ��� where �ⷿid=[1] " & _
                          " And ҩƷid=[2] And ����=1 And Nvl(����,0)=[3] "
                Else
                    gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ʵ������ from ҩƷ��� where �ⷿid=[1] " & _
                          " And ҩƷid=[2] And ����=1 "
                End If
            Else
                If Get��������(mlngStockID, Val(.TextMatrix(i, 0))) = 1 Then
                    gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ʵ������ from ҩƷ��� where �ⷿid=[1] " & _
                        " And ҩƷid=[2] And ����=1 And Nvl(����,0)>0 "
                Else
                    gstrSQL = " Select Sum(Nvl(��������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ��������, Sum(Nvl(ʵ������,0))/" & .TextMatrix(i, mconIntCol����ϵ��) & " as ʵ������ from ҩƷ��� where �ⷿid=[1] " & _
                          " And ҩƷid=[2] And ����=1 "
                End If
            End If
            
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ǰҩ��ʵ������]", mlngStockID, Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, mconIntCol����)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                If mint�༭״̬ = 6 Then
                    '����(���)ʱ��ʾʵ������
                    dblStock = IIf(IsNull(rsUseCount.Fields(1)), 0, rsUseCount.Fields(1))
                Else
                    '����״̬ʱ������
                    dblStock = IIf(mint��ʾ��ǰ��淽ʽ = 0, IIf(IsNull(rsUseCount.Fields(1)), 0, rsUseCount.Fields(1)), IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0)))
                End If
            End If
            .TextMatrix(i, mconintcol��ǰ���) = zlStr.FormatEx(dblStock, mintNumberDigit, , True)
       Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function GetProvider(ByVal lngProviderID As Long) As String
    Dim rstemp As ADODB.Recordset
    
    If lngProviderID <= 0 Then Exit Function
    On Error GoTo errHandle
    gstrSQL = "select ���� from ��Ӧ�� where id=[1]"
    Set rstemp = zlDataBase.OpenSQLRecord(gstrSQL, "��Ӧ������", lngProviderID)
    If Not rstemp.EOF Then
        GetProvider = zlStr.Nvl(rstemp!����)
    End If
    rstemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData(ByVal rstemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ���ʱ��ҩƷ�Ƿ��п��
    'ͬһҩƷ����ͬʱ���ڲ�����(����Ϊ0���ͷ����ļ�¼
    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim strInfo As String
    Dim strInfo������� As String
    Dim rsPrice As ADODB.Recordset
    Dim str��� As String
    Dim strDub As String    '�ظ�ҩƷ
    Dim strNotNum As String  '�޿��ҩƷ
    Dim str�ظ�ҩ�� As String   '������¼�ظ�ѡ���˵�ҩƷ����
    Dim strNotҩ�� As String    '������¼��ЩҩƷ��ʱ�۵��޿��
    Dim rsRe As ADODB.Recordset
    Dim str�������Լ�� As String
        
    On Error GoTo errHandle
    
    rstemp.MoveFirst
    
    Do While Not rstemp.EOF
        str���� = IIf(IsNull(rstemp!����), "0", rstemp!����)
        If InStr(1, strTemp, rstemp!ҩƷID & "," & str����) = 0 Then
            strTemp = strTemp & rstemp!ҩƷID & "," & str���� & "," & rstemp!ͨ���� & "|"
        End If
        rstemp.MoveNext
    Loop
        
    With mshBill    '���ظ��Ĳ�ѯ����
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol����)) > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
            End If
        Next
        
        '����Ƿ�ͬʱ��������Ϊ0�����β�Ϊ0������
        rstemp.MoveFirst
        Do While Not rstemp.EOF
            For i = 1 To .rows - 2
                '���صļ�¼���ķ������Ժͽ������еķ������Բ�һ��ʱ�������������ȡ���ݵ�����
                If rstemp!ҩƷID = Val(.TextMatrix(i, 0)) And _
                    ((Nvl(rstemp!����, 0) = 0 And Val(.TextMatrix(i, mconIntCol����)) > 0) Or _
                    (Nvl(rstemp!����, 0) > 0 And Val(.TextMatrix(i, mconIntCol����)) = 0)) Then
                    
                    '���뵽��Ҫ�ų����嵥��
                    If InStr(1, strInfo�������, rstemp!ҩƷID & "," & Nvl(rstemp!����, 0)) = 0 Then
                         strInfo������� = strInfo������� & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
                    End If
                    
                    '���뵽�������ѵ��嵥��
                    If InStr(1, "," & str�������Լ�� & ",", "," & .TextMatrix(i, mconIntColҩ��) & ",") = 0 Then
                        str�������Լ�� = IIf(str�������Լ�� = "", "", str�������Լ�� & ",") & .TextMatrix(i, mconIntColҩ��)
                    End If
                End If
            Next
            rstemp.MoveNext
        Loop
        
        'ͬһҩƷ��ͬ���ε�
        If strInfo <> "" Then   'Ϊ��������ƴ��sql
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
        
        'ͬһҩƷ��ǰѡ������κ��б����������Բ�һ�µ�
        If strInfo������� <> "" Then   'Ϊ��������ƴ��sql
            For i = 0 To UBound(Split(strInfo�������, "|")) - 1
                strDub = strDub & "ҩƷid<>" & Split(Split(strInfo�������, "|")(i), ",")(0) & " and "
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
                
        '�ж���ʲô��ʽƴ��sql
        If str�ظ�ҩ�� <> "" Then MsgBox str�ظ�ҩ�� & "�б����Ѿ��и�ҩƷ����ͬ���Σ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
        If str�������Լ�� <> "" Then MsgBox str�������Լ�� & vbCrLf & "������ѡҩƷ���б��д����ҷ������Բ�һ�£�������ӣ�", vbInformation, gstrSysName
        
        If strDub <> "" Then
            rstemp.Filter = strDub
        End If
        
        Set CheckData = rstemp
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Get�۸�(ByVal lngҩƷid As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double) As Double
    Dim rsPrice As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(����,0),0,ʵ�ʽ��/ʵ������,Nvl(���ۼ�,ʵ�ʽ��/ʵ������))*" & dbl����ϵ�� & " as  �ۼ� " _
        & "  from ҩƷ��� " _
        & " where �ⷿid=[1] " _
        & " and ҩƷid=[2] " _
        & " and ����=1 and ʵ������>0 and " _
        & " nvl(����,0)=[3]"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lngҩƷid, lng����)

    If Not rsPrice.EOF Then
        Get�۸� = rsPrice.Fields(0).Value
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRepeatDrugID(ByVal rstemp As ADODB.Recordset, ByVal intRecEnd As Integer, ByVal lngDrugID As Long) As Boolean
'----------------------
'���ܣ����¼���ظ�ҩƷ
'----------------------
    Dim i As Integer
    Dim rsClone As ADODB.Recordset
    
    CheckRepeatDrugID = False
    Set rsClone = rstemp.Clone
    With rsClone
        .Sort = "ҩƷid,����,���"
        .MoveFirst
        For i = 1 To .RecordCount
            If i > intRecEnd Then
                If lngDrugID = !ҩƷID Then
                    CheckRepeatDrugID = True
                    Exit Function
                End If
            End If
            .MoveNext
        Next
    End With

End Function

Private Sub SumQuantity(ByRef arrVal As Variant, ByVal lngDrugID As Long, ByVal dblQTY As Double)
'------------------------
'���ܣ�����ͬҩƷID������
'------------------------
    Dim i As Integer
    Dim blnFind As Boolean
    
    If UBound(arrVal) > 0 Then
        For i = 0 To UBound(arrVal, 2) - 1
            If arrVal(0, i) = lngDrugID Then
                arrVal(1, i) = arrVal(1, i) + dblQTY
                blnFind = True
                Exit For
            End If
        Next
    Else
        ReDim arrVal(2, 1)
        arrVal(0, 0) = lngDrugID
        arrVal(1, 0) = dblQTY
        blnFind = True
    End If
    If blnFind = False Then
        ReDim Preserve arrVal(2, UBound(arrVal) + 1)
        arrVal(0, UBound(arrVal)) = lngDrugID
        arrVal(1, UBound(arrVal)) = dblQTY
    End If
End Sub

Private Function GetQuantity(ByVal arrVal As Variant, ByVal lngDrugID As Long, ByVal dblQTY As Double) As Double
'----------------------------
'���ܣ���ȡ������ҩƷID������
'----------------------------
    If UBound(arrVal) > 0 Then
        Dim i As Integer
        For i = 0 To UBound(arrVal, 2) - 1
            If arrVal(0, i) = lngDrugID Then
                GetQuantity = arrVal(1, i) + dblQTY
                Exit Function
            End If
        Next
    End If
    GetQuantity = dblQTY
End Function

Private Function ���۸�() As Boolean
    '���ܣ�����ʱ���ж�ҩƷ�Ƿ������¼۸񣬲������޸ĺ���ʾ
    Dim strMsg As String '������ʾ��Ϣ
    Dim i As Integer, intSum As Integer, intPricedigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim bln�Ƿ�ʱ�� As Boolean
    Dim lngStockid As Long
    Dim bln���� As Boolean
    
    On Error GoTo errHandle
    
    ���۸� = False
    lngStockid = cboStock.ItemData(cboStock.ListIndex)
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Trim(.TextMatrix(i, mconIntCol��д����)) <> "" Then
                bln���� = Get��������(lngStockid, Val(.TextMatrix(i, 0))) '�Ƿ����
                bln�Ƿ�ʱ�� = Val(Split(.TextMatrix(i, mconIntCol���Ч��), "||")(1)) = 1
                Dbl���� = Val(.TextMatrix(i, mconIntColʵ������))
                    
                If (bln���� And Val(.TextMatrix(i, mconIntCol����)) <> 0) Or Not bln���� Then '���������β�Ϊ0�򲻷����ĲŽ��м۸��飨�������������п��ܲ���飩
                    
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
                
                '���۷�����û��ȷ����Ҳ����ۼ�
                If bln�Ƿ�ʱ�� = False And (bln���� And Val(.TextMatrix(i, mconIntCol����)) = 0) Then
                    '����ۼ�
                    dbl���ۼ� = zlStr.FormatEx(Get�ۼ�(bln�Ƿ�ʱ��, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol����))) * Val(.TextMatrix(i, mconIntCol����ϵ��)), mintPriceDigit)
                    If .TextMatrix(i, mconIntCol�ۼ�) <> dbl���ۼ� Then
                        intSum = intSum + 1
                        .TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, mintPriceDigit, , True)
                        .TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(i, mconIntCol�ۼ�) * Dbl����, mintMoneyDigit, , True)
                    End If
                    
                    .TextMatrix(i, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(i, mconIntCol�ۼ۽��)) - Val(.TextMatrix(i, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                End If
                
            End If
        Next
        
        If intSum > 0 Then '����0��ʾ�м۸����
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

Private Sub CheckNumber(Optional int����״̬ As Integer = 0)
    '�����д������ʵ��������һ�£����ú�ɫ�����עʵ��������������
    Dim intRow As Integer, j As Integer
    Dim blnColor As Boolean
    With mshBill
        If int����״̬ = 1 Then
            blnColor = False
            If .TextMatrix(.Row, 0) = "" Then Exit Sub
            If Val(.TextMatrix(.Row, mconIntCol��д����)) <> Val(.TextMatrix(.Row, mconIntColʵ������)) Then blnColor = True
            j = .ColData(mconIntColʵ������)
            If j = 5 Then mshBill.ColData(mconIntColʵ������) = 0
            .Col = mconIntColʵ������
            .MsfObj.CellForeColor = IIf(blnColor, &HFF&, &H0&)
            .ColData(mconIntColʵ������) = j
        Else
            For intRow = 1 To .rows - 1
                blnColor = False
                If .TextMatrix(intRow, 0) = "" Then Exit Sub
                If Val(.TextMatrix(intRow, mconIntCol��д����)) <> Val(.TextMatrix(intRow, mconIntColʵ������)) Then blnColor = True
                j = .ColData(mconIntColʵ������)
                If j = 5 Then .ColData(mconIntColʵ������) = 0
                .Row = intRow
                .Col = mconIntColʵ������
                .MsfObj.CellForeColor = IIf(blnColor, &HFF&, &H0&)
                .ColData(mconIntColʵ������) = j
            Next
        End If
    End With
End Sub
