VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmOtherInputCard 
   Caption         =   "ҩƷ������ⵥ"
   ClientHeight    =   7905
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14715
   Icon            =   "frmOtherInputCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   14715
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2175
      Left            =   2760
      TabIndex        =   37
      Top             =   1380
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox PicInput 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   240
      ScaleHeight     =   1635
      ScaleWidth      =   2775
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1620
      Visible         =   0   'False
      Width           =   2805
      Begin VB.CommandButton CmdYes 
         Caption         =   "ȷ��"
         Height          =   345
         Left            =   810
         TabIndex        =   35
         Top             =   1140
         Width           =   855
      End
      Begin VB.CommandButton CmdNO 
         Caption         =   "ȡ��"
         Height          =   345
         Left            =   1800
         TabIndex        =   36
         Top             =   1140
         Width           =   855
      End
      Begin VB.TextBox Txt�Ӽ��� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         MaxLength       =   8
         TabIndex        =   34
         Text            =   "15.0000"
         Top             =   690
         Width           =   1725
      End
      Begin VB.Label Lbl�Ӽ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ӳ���(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   90
         TabIndex        =   33
         Top             =   750
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "    ������ӳ��ʣ����ۼ۵ļ��㹫ʽ�����ۼ�=�ɱ���*(1+�ӳ���%)"
         ForeColor       =   &H00400000&
         Height          =   585
         Left            =   0
         TabIndex        =   32
         Top             =   150
         Width           =   2805
      End
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6240
      TabIndex        =   30
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7560
      TabIndex        =   29
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   10
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6240
      TabIndex        =   6
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7560
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   14655
      TabIndex        =   12
      Top             =   0
      Width           =   14715
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   3
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
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   5
         Top             =   4080
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
      Begin VB.Label lbl�޸����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸�����"
         Height          =   180
         Left            =   7020
         TabIndex        =   41
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label lbl�޸��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�޸���"
         Height          =   180
         Left            =   5160
         TabIndex        =   40
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Txt�޸��� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5760
         TabIndex        =   39
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt�޸����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7800
         TabIndex        =   38
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   27
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   25
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10590
         TabIndex        =   23
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   12690
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   21
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   20
         Top             =   4440
         Width           =   1005
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
         TabIndex        =   4
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ������ⵥ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
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
         TabIndex        =   15
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   10005
         TabIndex        =   14
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   11880
         TabIndex        =   13
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������(&T)"
         Height          =   180
         Left            =   8040
         TabIndex        =   2
         Top             =   660
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
            Picture         =   "frmOtherInputCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1000
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
            Picture         =   "frmOtherInputCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOtherInputCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   7545
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmOtherInputCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19606
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherInputCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmOtherInputCard.frx":3080
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
Attribute VB_Name = "frmOtherInputCard"
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
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mbln�¿������� As Boolean           '��Ƿ��¿�������
Private mrs�ֶμӳ� As ADODB.Recordset      '�ֶμӳɼ���
Private mintʱ�۷ֶμӳɷ�ʽ As Integer     ' 0-�����ֶμӳɣ�Ĭ�ϣ� 1-���ֶμӳ�
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private mintȡ�ϴγɱ��۷�ʽ As Integer     '0-���ȴ�ҩƷ���ȡ;1-���ȴ�ҩƷ���ȡ
Private marrFrom As Variant                   '��¼�û��ָ�������и���
Private marrInitGrid As Variant                '��¼��ʼ��������и���

Private mintLastCol As Integer              '�û����������е����ɼ��е��к�

Private mrsInOutType As Recordset           '������
Private mbln�Ӽ��� As Boolean               'ʱ��ҩƷ�Ƿ��������Ӽ���
Private mdbl�Ӽ��� As Double
Private mstrPrivs As String                 'Ȩ��

'Private mintʱ���ۼ�λ�� As Integer         '��¼ʱ��ҩƷ�û��Զ���С��λ��

Private mcolUsedCount As Collection         '��ʹ�õ���������
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������

Private mlng���ⷿ As Long
Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
Private Const MStrCaption As String = "ҩƷ����������"

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

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

Private mstrѡ���� As String
Private mstr������ As String

Private mlng�����̳��� As Long                 '�������ֶγ���
Private mlngԭ���س��� As Long                 'ԭ�����ֶγ���

'=========================================================================================
Private mconIntCol�к� As Integer
Private mconIntColҩ�� As Integer
Private mconIntCol��Ʒ�� As Integer
Private mconIntCol��Դ As Integer
Private mconIntCol����ҩ�� As Integer
Private mconIntCol��� As Integer
Private mconIntCol��� As Integer
Private mconIntColԭ������ As Integer
Private mconIntColԭ���� As Integer
Private mconIntCol����ϵ�� As Integer
Private mconIntCol���� As Integer
Private mconIntColԭ���� As Integer
Private mconIntCol��λ As Integer
Private mconIntCol���� As Integer
Private mconIntCol�������� As Integer
Private mconIntColЧ�� As Integer
Private mconIntCol��׼�ĺ� As Integer
Private mconIntCol��� As Integer
Private mconIntCol���� As Integer
Private mconIntCol�������� As Integer
Private mconIntCol�ɱ��� As Integer
Private mconIntCol�ɱ���� As Integer
Private mconIntCol�ۼ� As Integer
Private mconIntCol�ۼ۽�� As Integer
Private mconintCol��� As Integer

Private mconintCol���ۼ� As Integer
Private mconintCol���۵�λ As Integer
Private mconintCol���۽�� As Integer
Private mconintCol���۲�� As Integer

Private mconintCol��ʵ���� As Integer
Private mconIntCol�������� As Integer
Private mconIntCol�Ƿ����� As Integer
Private mconIntColҩƷ��������� As Integer
Private mconIntColҩƷ���� As Integer
Private mconIntColҩƷ���� As Integer
Private mconIntCol���� As Integer
Private Const mconIntColS = 37
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

'�������������
Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    
    On Error GoTo errHandle
    GetDepend = False
    gstrSQL = "SELECT B.Id,b.���� " _
        & " FROM ҩƷ�������� A, ҩƷ������ B " _
        & "Where A.���id = B.ID " _
      & "AND A.���� = 4 "
    Call SQLTest(App.Title, "ҩƷ����������", gstrSQL)
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "GetDepend")
    Call SQLTest
    If rsDepend.EOF Then
        MsgBox "û������ҩƷ������������������ҩƷ������࣡", vbInformation + vbOKOnly, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    Set mrsInOutType = rsDepend
       
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
    mstrPrivs = GetPrivFunc(glngSys, 1302)
    mintʱ�۷ֶμӳɷ�ʽ = gtype_UserSysParms.P181_ҩƷ��ⰴ�ֶμӳ�
        
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
    Dim str������ As String
    
    On Error GoTo errHandle
    
    str�ⷿ���� = ""
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
        Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
        Do While Not rsDetail.EOF
            str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
            rsDetail.MoveNext
        Loop
        If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
    
        str������ = zlDataBase.GetPara("������", glngSys, ģ���.�������)
        
        If InStr(1, "|" & str������ & "|", "|ԭ����|") = 0 Then mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
        
        If mblnLoad = True Then Call SetSelectorRS(1, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
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
                    
                    mlng���ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
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

Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(0, mintNumberDigit, , True)
                .TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(0, mintMoneyDigit, , True)
                
                Call Setʱ�۷���ҩƷ���ۼ�(intRow, Val(.TextMatrix(intRow, mconintCol���ۼ�)))
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
                .TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ɱ���), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol�ۼ�), mintMoneyDigit, , True)
                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ۽��) - .TextMatrix(intRow, mconIntCol�ɱ����), mintMoneyDigit, , True)
                
                Call Setʱ�۷���ҩƷ���ۼ�(intRow, Val(.TextMatrix(intRow, mconintCol���ۼ�)))
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
        mstrTime_End = GetBillInfo(4, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not ��鵥��(4, txtNo.Tag, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�������ۼۣ������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
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
                    If Val(mshBill.TextMatrix(intLop, mconIntCol�ɱ���)) <> Val(mshBill.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                        MsgBox "��" & intLop & "��ҩƷ���������۹�������ⵥ���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
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
            Exit Sub
        End If
        
        gcnOracle.CommitTrans
        
        If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.�������)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.�������)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1302_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
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
        
        If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
            MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
            txtժҪ.SetFocus
            Exit Sub
        End If
        
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
        If Not ��鵥��(4, txtNo.Tag, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�������ۼۣ������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If mint�༭״̬ = 1 Then '��������ʱ���ж��ۼ��Ƿ��Ѿ�����
        If ����ۼ� Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
        
    If ValidData = False Then Exit Sub
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.�������)) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
                
                If Val(zlDataBase.GetPara("��ӡҩƷ����", glngSys, ģ���.�������)) = 1 And zlStr.IsHavePrivs(mstrPrivs, "ҩƷ�����ӡ") Then
                    '��ҩƷID˳���������
                    recSort.Sort = "ҩƷid"
                    recSort.MoveFirst
                    '��ӡҩƷ����
                    Do While Not recSort.EOF
                        If lng�ϴ�ҩƷID <> Val(recSort!ҩƷID) Then
                            ReportOpen gcnOracle, glngSys, "ZL1_INSIDE_1302_1", Me, "ҩƷ=" & Val(recSort!ҩƷID), 2
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
    SetEdit
    
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
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
        
    gstrSQL = " Select �շ�ϸĿID,nvl(�ּ�,0) �ּ� From �շѼ�Ŀ " & _
            " Where (��ֹ���� Is NULL Or sysdate Between ִ������ And nvl(��ֹ����,to_date('3000-01-01','yyyy-MM-dd')))" & _
            GetPriceClassString("")
    gstrSQL = "Select A.���,A.ҩƷID,B.�ּ� From ҩƷ�շ���¼ A,(" & gstrSQL & ") B,�շ���ĿĿ¼ C" & _
            " Where A.����=4 And A.NO=[1] And A.ҩƷID=B.�շ�ϸĿID And C.ID=B.�շ�ϸĿID And Round(A.���ۼ�," & intPriceDigit & ")<>Round(B.�ּ�," & intPriceDigit & ") And Nvl(C.�Ƿ���,0)=0" & _
            " Union All " & _
            " Select A.���, A.ҩƷid, decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) �ּ� " & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C, ҩƷ��� D , " & _
            "      (Select x.ҩƷid,x.�ⷿid,x.����,x.�ּ� from ҩƷ�۸��¼ x where x.�۸����� = 1 and (x.��ֹ���� Is Null Or Sysdate Between x.ִ������ And Nvl(x.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) X" & _
            " Where A.���� = 4 And A.NO = [1] And C.ID = A.ҩƷid And Round(A.���ۼ�, " & intPriceDigit & ") <> Round(decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�), " & intPriceDigit & ") And " & _
            " Nvl(C.�Ƿ���, 0) = 1 And D.ҩƷid = A.ҩƷid And B.���� = 1 And B.�ⷿid = A.�ⷿid And B.ҩƷid = A.ҩƷid And " & _
            " a.ҩƷid = x.ҩƷid(+) And a.�ⷿid = x.�ⷿid(+) And Nvl(a.����, 0) = Nvl(x.����(+), 0) AND " & _
            " Nvl(B.����, 0) = Nvl(A.����, 0) And NVL(b.ʵ������, 0) <> 0 And decode(x.�ּ�,null,decode(nvl(b.���ۼ�,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�),x.�ּ�) > 0 " & _
            " Order by ҩƷid,���"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        If lngҩƷID <> 0 Then
            rsPrice.Filter = "ҩƷID=" & lngҩƷID
            If rsPrice.RecordCount <> 0 Then
                '�Ե�ǰ���¼۸����µ���������ݣ����ۡ����۽���ۣ�
                dbl���ۼ� = rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��))
                dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ɱ���))
                Dbl���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
                dbl�ɱ���� = dbl�ɱ��� * Dbl����
                dbl���۽�� = dbl���ۼ� * Dbl����
                dbl��� = dbl���۽�� - dbl�ɱ����
                
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, intPriceDigit, , True)
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(dbl���۽��, mintMoneyDigit, , True)
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


Private Sub Form_Load()
    Dim rsDetail As New Recordset
    Dim bln��ҩ�ⷿ As Boolean
    Dim str�ⷿ���� As String
    Dim i As Integer, j As Integer
    Dim str������ As String
    
    mblnLoad = False
    marrFrom = Array()
    marrInitGrid = Array()
    mintBatchNoLen = GetBatchNoLen()
    mbln�Ӽ��� = Get�Ӽ���
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mintȡ�ϴγɱ��۷�ʽ = Val(zlDataBase.GetPara("ȡ�ϴγɱ��۷�ʽ", glngSys, ģ���.�⹺���))
    
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    mblnUpdate = False
    
    On Error GoTo errHandle
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ����������", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    Call GetDefineSize
    Call GetSysParm
    
    Set mrs�ֶμӳ� = Nothing
    If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
        gstrSQL = "select ���, ��ͼ�, ��߼�, �ӳ���, ��۶�, ˵��, ���� from ҩƷ�ӳɷ��� order by ���"
        Set mrs�ֶμӳ� = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�ֶμӳ�")
    End If
    mshBill.Value = Format(Sys.Currentdate, "YYYY-MM-DD")
    
    With cboType
        .Clear
        Do While Not mrsInOutType.EOF
            .AddItem mrsInOutType.Fields(1)
            .ItemData(.NewIndex) = mrsInOutType.Fields(0)
            mrsInOutType.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    mlng���ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    Call GetDrugDigit(mlng���ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)

    Call initCard
    
    mstrTime_Start = GetBillInfo(4, mstr���ݺ�)
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    
    'ֻ����ҩ��ⷿ����ʾ"ԭ����"��
    str�ⷿ���� = ""
    gstrSQL = "Select �������� From ��������˵�� Where ����id =[1]"
    Set rsDetail = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ����", cboStock.ItemData(cboStock.ListIndex))
    Do While Not rsDetail.EOF
        str�ⷿ���� = str�ⷿ���� & "," & rsDetail!��������
        rsDetail.MoveNext
    Loop
    If str�ⷿ���� Like "*��ҩ*" Or str�ⷿ���� Like "*�Ƽ���*" Then bln��ҩ�ⷿ = True
    str������ = zlDataBase.GetPara("������", glngSys, ģ���.�������)
    If InStr(1, "|" & str������ & "|", "|ԭ����|") = 0 Then mshBill.ColWidth(mconIntColԭ����) = IIf(bln��ҩ�ⷿ, 800, 0)
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrInitGrid(UBound(marrInitGrid) + 1)
        marrInitGrid(UBound(marrInitGrid)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    RestoreWinState Me, App.ProductName, MStrCaption
    
    For i = 1 To mconIntColS - 1
        ReDim Preserve marrFrom(UBound(marrFrom) + 1)
        marrFrom(UBound(marrFrom)) = mshBill.TextMatrix(0, i) & "|" & mshBill.ColWidth(i)
    Next
    
    For i = 0 To UBound(marrInitGrid)
        For j = 0 To UBound(marrFrom)
            If Split(marrInitGrid(i), "|")(0) = Split(marrFrom(j), "|")(0) And Split(marrInitGrid(i), "|")(1) * Split(marrFrom(j), "|")(1) = 0 Then
                mshBill.ColWidth(i + 1) = Split(marrInitGrid(i), "|")(1)
            End If
        Next
    Next
  
    mshBill.ColWidth(mconIntCol��������) = IIf(mint�༭״̬ = 6, 1100, 0)
    
    If mintUnit = mconint�ۼ۵�λ Then
        mshBill.ColWidth(mconintCol���ۼ�) = 0
        mshBill.ColWidth(mconintCol���۵�λ) = 0
        mshBill.ColWidth(mconintCol���۽��) = 0
        mshBill.ColWidth(mconintCol���۲��) = 0
    Else
        mshBill.ColWidth(mconintCol���ۼ�) = 0
        mshBill.ColWidth(mconintCol���۵�λ) = 0
        mshBill.ColWidth(mconintCol���۽��) = 0
        mshBill.ColWidth(mconintCol���۲��) = 0
        
        If InStr(1, "|" & mstr������ & "|", "|���ۼ�|") = 0 Then mshBill.ColWidth(mconintCol���ۼ�) = 1000
        If InStr(1, "|" & mstr������ & "|", "|���۵�λ|") = 0 Then mshBill.ColWidth(mconintCol���۵�λ) = 1000
        If InStr(1, "|" & mstr������ & "|", "|���۽��|") = 0 Then mshBill.ColWidth(mconintCol���۽��) = 1000
        If InStr(1, "|" & mstr������ & "|", "|���۲��|") = 0 Then mshBill.ColWidth(mconintCol���۲��) = 1000
    End If
    
    '������ԱȨ���жϣ��Ƿ���ʾ�ɱ���
    If InStr(1, "|" & mstr������ & "|", "|�ɱ���|") = 0 Then mshBill.ColWidth(mconIntCol�ɱ���) = IIf(mblnViewCost, 1000, 0)
    If InStr(1, "|" & mstr������ & "|", "|�ɱ����|") = 0 Then mshBill.ColWidth(mconIntCol�ɱ����) = IIf(mblnViewCost, 900, 0)
    If InStr(1, "|" & mstr������ & "|", "|���|") = 0 Then mshBill.ColWidth(mconintCol���) = IIf(mblnViewCost, 900, 0)
    If InStr(1, "|" & mstr������ & "|", "|���۲��|") = 0 Then mshBill.ColWidth(mconintCol���۲��) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconintCol��ʵ����) = 0
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
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
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strOrder As String, strCompare As String
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim str���� As String
    Dim strArray As String
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPriceDigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    '�ⷿ
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.�������)
    strCompare = Mid(strOrder, 1, 1)
    
    On Error GoTo errHandle
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
                gstrSQL = "select b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id and A.���� = 4 and a.no=[1]"
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
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,F.���㵥λ AS ��λ, A.��д���� AS ����,b.ָ�������� as ָ��������, a.�ɱ���,A.���ۼ�,1 as ����ϵ��,"
                Case mconint���ﵥλ
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ����,b.ָ��������*B.�����װ as ָ�������� , a.�ɱ���*B.�����װ as �ɱ���,A.���ۼ�*B.�����װ as ���ۼ� ,B.�����װ as ����ϵ��,"
                Case mconintסԺ��λ
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ����,b.ָ��������*B.סԺ��װ as ָ�������� , a.�ɱ���*B.סԺ��װ as �ɱ���,A.���ۼ�*B.סԺ��װ as ���ۼ� ,  B.סԺ��װ as ����ϵ��,"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "F.���㵥λ AS �ۼ۵�λ,B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ����,b.ָ��������*B.ҩ���װ as ָ�������� , a.�ɱ���*B.ҩ���װ as �ɱ���,A.���ۼ�*B.ҩ���װ as ���ۼ� ,B.ҩ���װ as ����ϵ��,"
            End Select
            
            If mint�༭״̬ <> 6 Then
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.ҩƷID,A.���,'[' ||F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��, " & _
                    " B.ҩƷ��Դ,B.����ҩ��,F.���,F.���� AS ԭ������,A.����, A.ԭ����,A.����,A.����," & _
                    " B.���Ч��,A.Ч��," & strUnitQuantity & " A.�ɱ����, " & _
                    " A.���۽��, A.���,B.�ӳ���/100 AS �ӳ���,F.�Ƿ���,B.ҩ������ AS ҩ����������, " & _
                    " A.ժҪ,������,��������,�޸���,�޸�����,�����,�������,A.�ⷿID,G.���� AS ����,A.������ID,A.��������,A.��׼�ĺ�,A.���, Nvl(A.�÷�, 0) As ���� " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F,���ű� G " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID AND A.�ⷿID=G.ID" & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND E.����(+)=1 " & _
                    " AND A.��¼״̬ =[2] " & _
                    " AND A.���� = 4 AND A.NO = [1])" & _
                    " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT * FROM (SELECT DISTINCT A.ҩƷID,A.���,'[' ||F.���� || ']' As ҩƷ����, F.���� As ͨ����, E.���� As ��Ʒ��, " & _
                    " B.ҩƷ��Դ,B.����ҩ��,F.���,F.���� AS ԭ������,A.����, A.ԭ����,A.����,A.����," & _
                    " B.���Ч��,A.Ч��," & strUnitQuantity & " A.�ɱ����, " & _
                    " 0 ���۽��,0 ���,B.�ӳ���/100 AS �ӳ���,F.�Ƿ���,B.ҩ������ AS ҩ����������, " & _
                    " A.�ⷿID,G.���� AS ����,A.������ID, A.��������,A.��׼�ĺ�,A.���,A.��д���� ��ʵ����,A.���� " & _
                    " FROM " & _
                    "     (SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,SUM(�ɱ����) AS �ɱ����,Sum(To_Number(Nvl(�÷�, 0))) As ����," & _
                    "     ҩƷID,���,����, ԭ����,����,nvl(����,0) as ����,Ч��,����,�ɱ���,���ۼ�,�ⷿID,������ID,X.��������,X.��׼�ĺ�,X.���" & _
                    "     FROM ҩƷ�շ���¼ X " & _
                    "     WHERE NO=[1] AND ����=4  " & _
                    "     GROUP BY ҩƷID,���,����, ԭ����,����,nvl(����,0),Ч��,����,�ɱ���,���ۼ�,�ⷿID,������ID,X.��������,X.��׼�ĺ�,X.���" & _
                    "     HAVING SUM(ʵ������)<>0 ) A," & _
                    "     ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ F,���ű� G " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=F.ID AND A.�ⷿID=G.ID" & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND E.����(+)=1 )" & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, mint��¼״̬)
            
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
                        Txt����� = ""
                        Txt������� = ""
                    Else
                        Txt������ = UserInfo.�û�����
                        Txt�������� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
'                        Txt�޸��� = UserInfo.�û�����
'                        Txt�޸����� = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
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
            
            If mint�༭״̬ <> 6 Then
                txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            Else
                txtժҪ.Text = GetժҪ(mstr���ݺ�)
            End If
            
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
            End With
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            intRow = 0
            With mshBill
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
                    
                    .TextMatrix(intRow, mconIntCol��Դ) = nvl(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = nvl(rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntCol���) = rsInitCard!���
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!ԭ����), "", rsInitCard!ԭ����)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol����) = zlStr.FormatEx(rsInitCard!����, intNumberDigit, , True)
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "", rsInitCard!��������)
                    If rsInitCard!���� <> 0 Then
                        .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(rsInitCard!�ɱ���, intCostDigit, , True)
                    Else
                        .TextMatrix(intRow, mconIntCol�ɱ���) = IIf(mintUnit = mconintҩ�ⵥλ, "0.00000", "0.0000000")
                    End If
                    .TextMatrix(intRow, mconIntCol�ɱ����) = zlStr.FormatEx(IIf(mint�༭״̬ = 6, 0, rsInitCard!�ɱ����), intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPriceDigit, , True)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                    .TextMatrix(intRow, mconIntColԭ������) = IIf(IsNull(rsInitCard!ԭ������), "!", rsInitCard!ԭ������)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol�Ƿ�����) = "��"
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconIntCol��������) = zlStr.FormatEx(0, intNumberDigit, , True)
                        .TextMatrix(intRow, mconintCol��ʵ����) = rsInitCard!��ʵ����
                    End If
                        
                    .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�ӳ��� & "||" & IIf(IsNull(rsInitCard!�Ƿ���), 0, rsInitCard!�Ƿ���) & "||" & IIf(IsNull(rsInitCard!ҩ����������), 0, rsInitCard!ҩ����������)
                        
                    '��������
                    Call GetҩƷ��������(intRow)
                    
                    'ʱ�۷���ҩƷ������Ҫ���������ۼۡ��ۼ۽����
                    If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
                        If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                            .TextMatrix(intRow, mconintCol���۵�λ) = rsInitCard!�ۼ۵�λ
                            .TextMatrix(intRow, mconintCol���ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ� / Val(rsInitCard!����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�, , True)
                            .TextMatrix(intRow, mconintCol���۽��) = zlStr.FormatEx(rsInitCard!���۽��, intMoneyDigit, , True)
                            .TextMatrix(intRow, mconintCol���۲��) = zlStr.FormatEx(rsInitCard!���, intMoneyDigit, , True)
                            
                            If mint�༭״̬ <> 6 Then
                                '���ǳ���ʱ
                                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(rsInitCard!����), intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۲��)) - Val(rsInitCard!����), intMoneyDigit, , True)
                                
                                If Val(.TextMatrix(intRow, mconIntCol����)) <> 0 And rsInitCard!���� <> 0 Then
                                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) / Val(rsInitCard!����), intPriceDigit, , True)
                                End If
                            Else
                                '����ʱ
                                .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(0, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(0, intMoneyDigit, , True)
                                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx((Val(.TextMatrix(intRow, mconintCol���ۼ�)) * Val(rsInitCard!����ϵ��) * Val(rsInitCard!����) - Val(rsInitCard!����)) / Val(rsInitCard!����), intPriceDigit, , True)
                            End If
                        End If
                    End If
                    
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!ҩƷID & "0") Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str���� = rsInitCard!ҩƷID & "0"
                        strArray = numUseAbleCount + IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                        mcolUsedCount.Add Array(str����, strArray), str����
                    End If
                    
                    rsInitCard.MoveNext
                Loop
                .rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    
    SetEdit         '���ñ༭����
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetժҪ(ByVal strNo As String) As String
    '��ȡ�µ�ժҪ
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
         '����(ȡ���һ�γ�����ժҪ)
    gstrSQL = "Select ժҪ From ҩƷ�շ���¼ Where ����=4 And No=[1] and (��¼״̬ =1 or mod(��¼״̬,3)=0) Order By ������� Desc "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡժҪ��Ϣ", strNo)
    
    If Not rsTemp.EOF Then
        GetժҪ = nvl(rsTemp!ժҪ)
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = IIf(mint�༭״̬ = 6, 5, 0)
            Next
            If mint�༭״̬ = 6 Then
                .ColData(mconIntColҩ��) = 0
                .ColData(mconIntCol��������) = 4
                txtժҪ.Enabled = True
            End If
            
            cboStock.Enabled = False
            cboType.Enabled = False
            
            If mint�༭״̬ <> 6 Then
                txtժҪ.Enabled = False
            End If
        Else
            .ColData(0) = 5
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol���) = 5
            .ColData(mconIntCol���) = 5
            If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                .ColData(mconIntCol����) = 1
                .ColData(mconIntColԭ����) = 1
            Else
                .ColData(mconIntCol����) = 5
                .ColData(mconIntColԭ����) = 5
            End If
            .ColData(mconIntCol��λ) = 5
            .ColData(mconIntCol����) = 4
            .ColData(mconIntCol����) = 5
            .ColData(mconIntCol��������) = 2
            .ColData(mconIntColЧ��) = 5
            .ColData(mconIntCol����) = 4
            .ColData(mconIntCol�ɱ���) = 4
            .ColData(mconIntCol�ɱ����) = 4
            .ColData(mconIntCol�ۼ�) = 5
            .ColData(mconIntCol�ۼ۽��) = 5
            .ColData(mconintCol���) = 5
            
            .ColData(mconIntColԭ������) = 5
            .ColData(mconIntColԭ����) = 5
            .ColData(mconIntCol����ϵ��) = 5
            .ColData(mconIntCol��׼�ĺ�) = 4
            .ColData(mconIntCol���) = 1
            
            .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol���) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntColԭ����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol��������) = flexAlignLeftCenter
            .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɱ���) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɱ����) = flexAlignRightCenter
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
        Call SetColumnByUserDefine
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
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol�ɱ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɱ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        .TextMatrix(0, mconintCol���ۼ�) = "���ۼ�"
        .TextMatrix(0, mconintCol���۵�λ) = "���۵�λ"
        .TextMatrix(0, mconintCol���۽��) = "���۽��"
        .TextMatrix(0, mconintCol���۲��) = "���۲��"
        .TextMatrix(0, mconIntColԭ������) = "ԭ������"
        .TextMatrix(0, mconIntColԭ����) = "ԭЧ��"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconintCol��ʵ����) = "��ʵ����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol�Ƿ�����) = "�Ƿ�����"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        
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
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntCol��������) = 1000
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        .ColWidth(mconIntCol���) = 1000
        .ColWidth(mconIntCol����) = 1100
        .ColWidth(mconIntCol��������) = IIf(mint�༭״̬ = 6, 1100, 0)
        .ColWidth(mconIntCol�ɱ���) = 1000
        .ColWidth(mconIntCol�ɱ����) = 900
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = 800
        .ColWidth(mconintCol���ۼ�) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconintCol���۵�λ) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconintCol���۽��) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconintCol���۲��) = IIf(mintUnit = mconint�ۼ۵�λ, 0, 1000)
        .ColWidth(mconIntColԭ������) = 0
        .ColWidth(mconIntColԭ����) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconintCol��ʵ����) = 0
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol�Ƿ�����) = 0
        
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0

        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntColҩ��) = 1
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol��Դ) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol���) = 5
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            .ColData(mconIntCol����) = 1
            .ColData(mconIntColԭ����) = 1
        Else
            .ColData(mconIntCol����) = 5
            .ColData(mconIntColԭ����) = 5
        End If
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol����) = 4
        .ColData(mconIntCol��������) = 2
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol��׼�ĺ�) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����) = 4
        .ColData(mconIntCol��������) = 4
        .ColData(mconIntCol�ɱ���) = 4
        .ColData(mconIntCol�ɱ����) = 4
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
        .ColData(mconintCol���ۼ�) = 5
        .ColData(mconintCol���۵�λ) = 5
        .ColData(mconintCol���۽��) = 5
        .ColData(mconintCol���۲��) = 5
        .ColData(mconIntColԭ������) = 5
        .ColData(mconIntColԭ����) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconintCol��ʵ����) = 5
        .ColData(mconIntCol�Ƿ�����) = 5
        
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Դ) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����ҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColԭ����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��������) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɱ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɱ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        .ColAlignment(mconintCol���ۼ�) = flexAlignRightCenter
        .ColAlignment(mconintCol���۵�λ) = flexAlignRightCenter
        .ColAlignment(mconintCol���۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���۲��) = flexAlignRightCenter
        .ColAlignment(mconintCol��ʵ����) = flexAlignRightCenter
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
    End With
    txtժҪ.MaxLength = Sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
    Call SetColumnByUserDefine
End Sub

Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer, Optional ByVal blnShow As Boolean = True)
    Select Case str����
    Case "ҩ��"
        mconIntColҩ�� = intValue
    Case "ҩƷ��Դ"
        mconIntCol��Դ = intValue
    Case "����ҩ��"
        mconIntCol����ҩ�� = intValue
    Case "���"
        mconIntCol��� = intValue
    Case "������"
        mconIntCol���� = intValue
    Case "ԭ����"
        mconIntColԭ���� = intValue
    Case "��λ"
        mconIntCol��λ = intValue
    Case "����"
        mconIntCol���� = intValue
    Case "��������"
        mconIntCol�������� = intValue
    Case "Ч��"
        mconIntColЧ�� = intValue
    Case "��׼�ĺ�"
        mconIntCol��׼�ĺ� = intValue
    Case "���"
        mconIntCol��� = intValue
    Case "����"
        mconIntCol���� = intValue
    Case "��������"
        mconIntCol�������� = intValue
    Case "�ɱ���"
        mconIntCol�ɱ��� = intValue
    Case "�ɱ����"
        mconIntCol�ɱ���� = intValue
    Case "�ۼ�"
        mconIntCol�ۼ� = intValue
    Case "�ۼ۽��"
        mconIntCol�ۼ۽�� = intValue
    Case "���"
        mconintCol��� = intValue
    Case "���ۼ�"
        mconintCol���ۼ� = intValue
    Case "���۵�λ"
        mconintCol���۵�λ = intValue
    Case "���۽��"
        mconintCol���۽�� = intValue
    Case "���۲��"
        mconintCol���۲�� = intValue
    End Select
    
    If Not blnShow Then
        mshBill.ColWidth(intValue) = 0
        mshBill.ColData(intValue) = 5
    Else
        mintLastCol = intValue
    End If
End Sub

Private Sub SetColumnByUserDefine()
    Dim intCol As Integer
    Dim arr����, arr��������
    Dim str���� As String, str�������� As String
    Dim intColumns As Integer
    Dim intCols As Integer
    Dim strColumn_Selected As String
    Dim strColumn_All As String
    Dim arrColumn_All, arrColumn_Selected, arrColumn_UnSelected
    On Error GoTo ErrHand
    mstrѡ���� = zlDataBase.GetPara("ѡ����", glngSys, ģ���.�������)
    mstr������ = zlDataBase.GetPara("������", glngSys, ģ���.�������)
    
    str���� = "ҩ��|ҩƷ��Դ|����ҩ��|���|������|ԭ����|����|��������|Ч��|��λ|����|��������|�ɱ���|�ɱ����|" & _
                        "�ۼ�|�ۼ۽��|���|��׼�ĺ�|���|���ۼ�|���۵�λ|���۽��|���۲��"

    '��δѡ����е��п�����Ϊ�㣬��������Ϊ5��������ѡ��
    If mstrѡ���� <> "" Then
        If InStr(1, "|" & mstr������ & "|", "|����|") <> 0 Then
            mstr������ = Replace("|" & mstr������ & "|", "|����|", "|������|")
            mstr������ = Left(mstr������, Len(mstr������) - 1)
            mstr������ = Mid(mstr������, 2)
        End If
        
        If InStr(1, "|" & mstrѡ���� & "|", "|����|") <> 0 Then
            mstrѡ���� = Replace("|" & mstrѡ���� & "|", "|����|", "|������|")
            mstrѡ���� = Left(mstrѡ����, Len(mstrѡ����) - 1)
            mstrѡ���� = Mid(mstrѡ����, 2)
        End If
        
        If mstr������ <> "" Then
            str�������� = mstrѡ���� & "|" & mstr������
        Else
            str�������� = mstrѡ����
        End If
        arr���� = Split(str����, "|")
        arr�������� = Split(str��������, "|")
        If UBound(arr����) <> UBound(arr��������) Or InStr(1, "|" & mstr������ & "|", "|������|") <> 0 Or InStr(1, "|" & mstrѡ���� & "|", "|������|") = 0 Or InStr(1, "|" & mstr������ & "|", "|�ɹ���|") <> 0 Or InStr(1, "|" & mstrѡ���� & "|", "|�ɹ���|") <> 0 Then
            mstrѡ���� = "ҩ��|ҩƷ��Դ|����ҩ��|���|������|ԭ����|����|��������|Ч��|��λ|����|��������|�ɱ���|�ɱ����|" & _
                        "�ۼ�|�ۼ۽��|���|��׼�ĺ�|���"
            mstr������ = "���ۼ�|���۵�λ|���۽��|���۲��"
            zlDataBase.SetPara "ѡ����", mstrѡ����, glngSys, ģ���.�������
            zlDataBase.SetPara "������", mstr������, glngSys, ģ���.�������
        End If
    Else
        mstrѡ���� = "ҩ��|ҩƷ��Դ|����ҩ��|���|������|ԭ����|����|��������|Ч��|��λ|����|��������|�ɱ���|�ɱ����|" & _
                    "�ۼ�|�ۼ۽��|���|��׼�ĺ�|���"
        mstr������ = "���ۼ�|���۵�λ|���۽��|���۲��"
        zlDataBase.SetPara "ѡ����", mstrѡ����, glngSys, ģ���.�������
        zlDataBase.SetPara "������", mstr������, glngSys, ģ���.�������
    End If


'    mstr������ = "|" & mstr������ & "|"
    With mshBill
        For intCol = 1 To .Cols - 1
            If InStr("|" & mstr������ & "|", "|" & .TextMatrix(0, intCol) & "|") > 0 Then
                .ColWidth(intCol) = 0
                .ColData(intCol) = 5
            End If
        Next
    End With
    
    strColumn_All = "ҩ��,2|ҩƷ��Դ,4|����ҩ��,5|���,7|������,11|ԭ����,12|��λ,13|����,14|��������,15|Ч��,16|��׼�ĺ�,17|���,18|����,19|��������,20|�ɱ���,21|�ɱ����,22|" & _
                    "�ۼ�,23|�ۼ۽��,24|���,25|���ۼ�,26|���۵�λ,27|���۽��,28|���۲��,29"

    '��װ��ȱʡ����
    mconIntCol�к� = 1
    mconIntColҩ�� = 2
    mconIntCol��Ʒ�� = 3
    mconIntCol��Դ = 4
    mconIntCol����ҩ�� = 5
    mconIntCol��� = 6
    mconIntCol��� = 7
    mconIntColԭ������ = 8
    mconIntColԭ���� = 9
    mconIntCol����ϵ�� = 10
    mconIntCol���� = 11
    mconIntColԭ���� = 12
    mconIntCol��λ = 13
    mconIntCol���� = 14
    mconIntCol�������� = 15
    mconIntColЧ�� = 16
    mconIntCol��׼�ĺ� = 17
    mconIntCol��� = 18
    mconIntCol���� = 19
    mconIntCol�������� = 20
    mconIntCol�ɱ��� = 21
    mconIntCol�ɱ���� = 22
    mconIntCol�ۼ� = 23
    mconIntCol�ۼ۽�� = 24
    mconintCol��� = 25
    mconintCol���ۼ� = 26
    mconintCol���۵�λ = 27
    mconintCol���۽�� = 28
    mconintCol���۲�� = 29
    mconintCol��ʵ���� = 30
    mconIntCol�������� = 31
    mconIntCol�Ƿ����� = 32
    mconIntColҩƷ��������� = 33
    mconIntColҩƷ���� = 34
    mconIntColҩƷ���� = 35
    mconIntCol���� = 36
    
    mintLastCol = 36
    
    '�����û����õ�����˳��
    arrColumn_All = Split(strColumn_All, "|")
    arrColumn_Selected = Split(mstrѡ����, "|")
    intCols = UBound(arrColumn_Selected)
    For intCol = 0 To intCols
        Call SetColumnValue(arrColumn_Selected(intCol), Split(arrColumn_All(intCol), ",")(1))
    Next
    
    intCol = intCols + 1
    intColumns = 0
    arrColumn_UnSelected = Split(mstr������, "|")
    intCols = UBound(arrColumn_All)
    For intCol = intCol To intCols
        If UBound(arrColumn_UnSelected) >= intColumns Then
            Call SetColumnValue(arrColumn_UnSelected(intColumns), Split(arrColumn_All(intCol), ",")(1), False)
            intColumns = intColumns + 1
        Else
            Call SetColumnValue(Split(arrColumn_All(intCol), ",")(0), Split(arrColumn_All(intCol), ",")(1), False)
        End If
    Next
    
    Exit Sub
ErrHand:
    MsgBox "�ָ�������ʱ�������������½��������ã�", vbInformation, gstrSysName
End Sub


Private Sub Setʱ�۷���ҩƷ���ۼ�(ByVal intRow As Integer, ByVal dblPrice As Double)
    Dim Dbl���� As Double

    With mshBill
        If .TextMatrix(intRow, mconIntColԭ����) = "" Then Exit Sub
        If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) <> 1 Or Val(.TextMatrix(intRow, mconIntCol��������)) <> 1 Then Exit Sub
        
       .TextMatrix(intRow, mconintCol���ۼ�) = zlStr.FormatEx(dblPrice, gtype_UserDrugDigits.Digit_���ۼ�, , True) '���ۼ��ֶα���������С��λ����˲�����ҩƷ���ľ������ý��п��ƣ�ֱ�Ӱ���7λ������ʾ
        
        If mint�༭״̬ = 6 Then
            Dbl���� = Val(.TextMatrix(intRow, mconIntCol��������)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
        Else
            Dbl���� = Val(.TextMatrix(intRow, mconIntCol����)) * Val(.TextMatrix(intRow, mconIntCol����ϵ��))
        End If
        
        If Val(.TextMatrix(intRow, mconIntCol�ɱ���)) = Val(.TextMatrix(intRow, mconIntCol�ۼ�)) Then
            'ͨ�����������������۹����⹺��⣬��ֹ���ֳ���������
            .TextMatrix(intRow, mconintCol���۽��) = .TextMatrix(intRow, mconIntCol�ۼ۽��)
        Else
            .TextMatrix(intRow, mconintCol���۽��) = zlStr.FormatEx(Dbl���� * Val(.TextMatrix(intRow, mconintCol���ۼ�)), mintMoneyDigit, , True)
        End If
        .TextMatrix(intRow, mconintCol���۲��) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(.TextMatrix(intRow, mconIntCol�ɱ����)), mintMoneyDigit, , True)
    End With
End Sub

Private Sub GetҩƷ��������(intBillRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim int�������� As Integer      '0-������;1-����
    Dim intҩ����� As Integer      '0-������;1-����
    Dim intҩ������ As Integer      '0-������;1-����
    Dim bln�Ƿ����ҩ������ As Boolean  'True-����ҩ������;False-������ҩ������
    
    If Val(mshBill.TextMatrix(intBillRow, 0)) = 0 Then Exit Sub
    On Error GoTo errHandle
    strSQL = "SELECT NVL(ҩ�����, 0) ҩ�����,NVL(ҩ������, 0) ҩ������ " & _
            " From ҩƷ��� WHERE ҩƷID = [1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "ȡҩƷ�ⷿ��������", Val(mshBill.TextMatrix(intBillRow, 0)))
    
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
            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "ȡ��������", cboStock.ItemData(Me.cboStock.ListIndex))
            
            bln�Ƿ����ҩ������ = (rsTemp.RecordCount > 0)
                    
            If bln�Ƿ����ҩ������ Then
                int�������� = 0
            Else
                int�������� = 1
            End If
        End If
    End If
    
    mshBill.TextMatrix(intBillRow, mconIntCol��������) = int��������
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
        '.Width = .Left - .Left
    End With
    
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 4
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
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ����������", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
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
    mblnSave = False
    SaveCheck = False
    
    Dim strҩƷ As String
    Dim n As Integer
    Dim m As Integer
    Dim dbl�ϼ����� As Double
    Dim lngҩƷID As Long
    
    '�����
    strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, mconIntCol����, mconIntCol����ϵ��, 1, , mintNumberDigit)
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
    
    gstrSQL = "zl_ҩƷ�������_Verify('" & txtNo.Tag & "','" & UserInfo.�û����� & "')"
    
    On Error GoTo errHandle
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
errHandle:
    'MsgBox "���ʧ�ܣ�", vbInformation, gstrSysName
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
    mshBill.Value = Format(Sys.Currentdate, "YYYY-MM-DD")
    mshBill.TextMatrix(Row, mconIntCol�Ƿ�����) = "��"
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
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
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow As Integer
    Dim rsProvider As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
        
    intOldRow = mshBill.Row
    
    On Error GoTo errHandle
    Select Case mshBill.Col
    Case mconIntColҩ��
        Dim RecReturn As Recordset
        
        mblnChange = True
        mshBill.CmdEnable = False
'        Set RecReturn = FrmҩƷѡ����.ShowME(Me, 1, , cboStock.ItemData(cboStock.ListIndex))
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(1, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
        End If
        
        Set RecReturn = frmSelector.ShowME(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , , , , mstrPrivs)
        
        mshBill.CmdEnable = True
        If RecReturn.RecordCount > 0 Then
            With mshBill
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    intRow = .Row
                    .TextMatrix(intRow, mconIntCol�к�) = .Row
                    SetColValue .Row, RecReturn!ҩƷID, _
                        "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                        nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, IIf(IsNull(RecReturn!���), "", RecReturn!���), _
                        IIf(IsNull(RecReturn!����), "", RecReturn!����), Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                        IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), RecReturn!ָ�������� * Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                        IIf(IsNull(RecReturn!����), "!", RecReturn!����), RecReturn!���Ч��, Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                        RecReturn!ʱ��, RecReturn!ҩ������, RecReturn!�ӳ��� / 100, _
                        IIf(IsNull(RecReturn!��������), "", Format(RecReturn!��������, "yyyy-mm-dd")), RecReturn!�ۼ۵�λ, RecReturn!ԭ����
'                    If .TextMatrix(.Row, mconIntColԭ������) = "!" Then
'                        .Col = mconIntCol����
'                    Else
'                        .Col = mconIntCol����
'                    End If
                    
                    .Col = GetNextEnableCol(mconIntColҩ��)
                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If

                    .Row = .rows - 1
                    RecReturn.MoveNext
                Next
                .Row = intOldRow
            End With
            RecReturn.Close
        End If
    Case mconIntCol����
        gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
        True, vRect.Left + 7000, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If rsProvider Is Nothing Then

            Exit Sub
        End If
        If Not rsProvider.EOF Then
            mshBill.TextMatrix(mshBill.Row, mconIntCol����) = rsProvider!����
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol����), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
            Else
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = ""
            End If
        End If
    Case mconIntColԭ����
        
        gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
        True, vRect.Left + 7800, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
        
        If rsProvider Is Nothing Then
            Exit Sub
        End If
        If Not rsProvider.EOF Then
            mshBill.TextMatrix(mshBill.Row, mconIntColԭ����) = rsProvider!����
        End If
    Case mconIntCol���
        Dim rs��� As New Recordset
                    
        gstrSQL = "Select ����,����,���� From ҩƷ��� Order By ����"
        Set rs��� = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ���")
                
        If rs���.EOF Then
            rs���.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rs���
            .StrNode = "����ҩƷ���"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol���) = .CurrentName
            End If
        End With
        Unload FrmSelect
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
        
        If .Col = mconIntCol���� Or .Col = mconIntCol�������� Or .Col = mconIntCol�ɱ��� Or .Col = mconIntCol�ۼ� Or .Col = mconintCol���ۼ� Or .Col = mconIntCol�ɱ���� Then
            Select Case .Col
                Case mconIntCol����, mconIntCol��������
                    intDigit = mintNumberDigit
                Case mconIntCol�ɱ���
                   intDigit = mintCostDigit
                Case mconIntCol�ۼ�
                    intDigit = mintPriceDigit
                Case mconintCol���ۼ�
                    intDigit = gtype_UserDrugDigits.Digit_���ۼ�
                Case mconIntCol�ɱ����
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
    Dim lngRow As Long
    Dim strxq As String
    Dim dblTemp�ۼ� As Double
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        If .Row <> .LastRow Then
            lngRow = .LastRow
            If PicInput.Visible Then
                '���¼������ۼۡ����
                dblTemp�ۼ� = Val(.TextMatrix(lngRow, mconIntCol�ɱ���)) * (1 + (Val(Txt�Ӽ���) / 100))
                .TextMatrix(lngRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(lngRow, 0)), Val(.TextMatrix(lngRow, mconIntCol�ɱ���)), Val(Txt�Ӽ���) / 100, dblTemp�ۼ�, lngRow), mintPriceDigit, , True)
                .TextMatrix(lngRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(lngRow, mconIntCol�ۼ�)) * Val(.TextMatrix(lngRow, mconIntCol����)), mintMoneyDigit, , True)
                .TextMatrix(lngRow, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(lngRow, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(lngRow, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(lngRow, mconIntCol�ɱ����) = "", 0, .TextMatrix(lngRow, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                PicInput.Visible = False
            End If
        End If
        SetInputFormat .Row
        
        'Modified by zyb 2002-10-30
        If Not (.Col = mconIntCol�ɱ��� Or .Col = mconIntCol�ɱ����) Then PicInput.Visible = False
        If .Col = mconIntCol�ɱ���� And PicInput.Visible Then Txt�Ӽ���.SetFocus: Exit Sub
        
        Select Case .Col
            Case mconIntColҩ��
                .txtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                
            Case mconIntCol����
                OS.OpenIme True
                .txtCheck = False
                .MaxLength = mlng�����̳���
                .TxtSetFocus
                
            Case mconIntColԭ����
                OS.OpenIme True
                .txtCheck = False
                .MaxLength = mlngԭ���س���
                .TxtSetFocus
                
            Case mconIntCol����
                .txtCheck = False
                '.TextMask = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
                .MaxLength = mintBatchNoLen
            Case mconIntCol��������
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol����) <> "" And Len(.TextMatrix(.Row, mconIntCol����)) = 8 Then
                    strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                    If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                        If IsNumeric(strxq) Then
                            If Trim(.TextMatrix(.Row, mconIntCol��������)) = "" Then
                                strxq = TranNumToDate(strxq)
                                If Trim(strxq) = "" Then Exit Sub
                                .TextMatrix(.Row, mconIntCol��������) = Format(strxq, "yyyy-mm-dd")
                            End If
                         End If
                    End If
                End If
            Case mconIntColЧ��
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If Trim(.TextMatrix(.Row, mconIntColԭ����)) = "" Then
                    Exit Sub
                End If
                If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0) = "0" Then
                    Exit Sub
                End If
                If .TextMatrix(.Row, mconIntCol��������) <> "" Then
                    If Trim(.TextMatrix(.Row, mconIntColЧ��)) = "" Then
                        strxq = UCase(.TextMatrix(.Row, mconIntCol��������))
                    End If
                ElseIf .TextMatrix(.Row, mconIntCol����) <> "" And Len(.TextMatrix(.Row, mconIntCol����)) = 8 Then
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) Then
                       If Trim(.TextMatrix(.Row, mconIntColЧ��)) = "" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                            End If
                        End If
                    End If
                End If
                If Trim(strxq) = "" Then Exit Sub
                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                    '����Ϊ��Ч��
                    .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                End If
                
                Call CheckLapse(.TextMatrix(.Row, mconIntColЧ��))
            Case mconIntCol�ɱ���
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                
            Case mconIntCol�ɱ����
                .txtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
                
            Case mconIntCol����
                .txtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mconIntCol��׼�ĺ�
                .txtCheck = False
                .MaxLength = 40
            Case mconIntCol�ۼ�
                .txtCheck = True
                .MaxLength = 11
                .TextMask = ".1234567890"
            Case mconIntCol���
                .txtCheck = False
                .MaxLength = 100
        End Select
        
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim dbl�ӳ��� As Double
    Dim strUnitQuantity As String
    Dim dblָ�����ۼ� As Double
    Dim rsTemp As ADODB.Recordset
    Dim strxq As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    Dim dblTemp�ۼ� As Double
    Dim rsProvider As Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim rsMaxs As New Recordset
    Dim ints���� As Integer, strCodes As String
    
    intOldRow = mshBill.Row
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
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
                If .TextMatrix(.Row, .Col) = "" Then
                    .TextMatrix(.Row, .Col) = " "
                End If
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 1, , cboStock.ItemData(cboStock.ListIndex), , strkey, sngLeft, sngTop)
                    
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(1, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
                    End If
                    
                    Set RecReturn = frmSelector.ShowME(Me, 1, 1, strKey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , , , , , , , mstrPrivs)
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                            If SetColValue(.Row, RecReturn!ҩƷID, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, _
                               IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, IIf(IsNull(RecReturn!���), "", RecReturn!���), _
                               IIf(IsNull(RecReturn!����), "", RecReturn!����), Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                               IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), RecReturn!ָ�������� * Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                               IIf(IsNull(RecReturn!����), "!", RecReturn!����), RecReturn!���Ч��, Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), RecReturn!ʱ��, _
                               RecReturn!ҩ������, RecReturn!�ӳ��� / 100, IIf(IsNull(RecReturn!��������), "", Format(RecReturn!��������, "yyyy-mm-dd")), RecReturn!�ۼ۵�λ, RecReturn!ԭ����) = False Then
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
                        Cancel = True
                    End If
                End If
                Call ��ʾ�����
                'End If
            Case mconIntCol����
                '����Ҳ�����Ӧ�Ĳ��أ�����������Ϊ����
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, .Col) = ""
                        .Text = " "
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
'                    Cancel = True
                    Exit Sub
                Else
                    vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
                    
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    If Trim(.Text) = "" Then Exit Sub
                    
                    gstrSQL = "Select ���� as id,���� ,����,���� From ҩƷ������ " _
                            & "Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And (upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & strKey & "%') " _
                                & "Order By ���� "
                                
                    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "������", False, "", "������ѡ��", False, False, _
                    True, vRect.Left + 7000, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
                    
                    If blnCancel = True Then mshBill.Text = "": .TextMatrix(.Row, mconIntCol����) = "": Exit Sub '��ѡ����ʱ����Esc�������´���
                    
                    If rsProvider Is Nothing Then
                        If MsgBox("ҩƷ������û���ҵ�������������̣���Ҫ��������ҩƷ����������", vbYesNo + vbQuestion, MStrCaption) = vbNo Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol����) = ""
                            mshBill.Text = ""
'                            Cancel = True
                            Exit Sub
                        Else
                            If LenB(StrConv(strKey, vbFromUnicode)) > mlng�����̳��� Then
                                MsgBox "���������ƹ���(���" & mlng�����̳��� & "���ַ���" & Int(mlng�����̳��� / 2) & "������)!", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                            
                            If rsMaxs.State = 1 Then rsMaxs.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱��볤��")
                            ints���� = rsMaxs!length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(����," & ints���� & ",'0')),'00') As Code FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱���")
                            strCodes = rsMaxs!Code
                            
                            ints���� = Len(strCodes)
                            strCodes = strCodes + 1
                            If ints���� >= Len(strCodes) Then
                                strCodes = String(ints���� - Len(strCodes), "0") & strCodes
                            End If
                            
                            gstrSQL = "ZL_ҩƷ������_INSERT('" & strCodes & "','" & strKey & "',zlSpellCode('" & strKey & "',10))"
                            
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        End If
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntCol����) = rsProvider!����
                        mshBill.Text = rsProvider!����
                        
                        gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol����), mshBill.TextMatrix(mshBill.Row, 0))
                        If Not rsProvider.EOF Then
                            mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
                        Else
                            mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = ""
                        End If
                    End If
                End If
                OS.OpenIme
            Case mconIntColԭ����
                '����Ҳ�����Ӧ�Ĳ��أ�����������Ϊ����
                If Trim(.Text) = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, .Col) = ""
                        .Text = " "
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    Exit Sub
                Else
                
                    vRect = zlControl.GetControlRect(mshBill.MsfObj.hWnd)
                    .Text = UCase(Trim(.Text))
                    strKey = Trim(.Text)
                    
                    gstrSQL = "Select ���� as id,���� ,����,���� From ҩƷ������ " _
                            & "Where (վ�� = '" & gstrNodeNo & "' Or վ�� is Null) And (upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & IIf(gstrMatchMethod = "0", "%", "") & strKey & "%' or Upper(����) like '" & strKey & "%') " _
                                & "Order By ���� "
                                
                    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "ԭ����", False, "", "ԭ����ѡ��", False, False, _
                    True, vRect.Left + 7800, vRect.Top, 300, blnCancel, False, True, gstrNodeNo)
                    
                    If blnCancel = True Then .Text = "": .TextMatrix(.Row, mconIntColԭ����) = "": Exit Sub '��ѡ����ʱ����Esc�������´���
                    
                    If rsProvider Is Nothing Then
                        If MsgBox("ҩƷ������û���ҵ��������ԭ���أ���Ҫ��������ҩƷ����������", vbYesNo + vbQuestion, MStrCaption) = vbNo Then
                            mshBill.TextMatrix(mshBill.Row, mconIntColԭ����) = ""
                            mshBill.Text = ""
                            Cancel = True
                            Exit Sub
                        Else
                            If LenB(StrConv(strKey, vbFromUnicode)) > mlngԭ���س��� Then
                                MsgBox "ԭ�������ƹ���(���" & mlngԭ���س��� & "���ַ���" & Int(mlngԭ���س��� / 2) & "������)!", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                Exit Sub
                            End If
                        
                            If rsMaxs.State = 1 Then rsMaxs.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱��볤��")
                            ints���� = rsMaxs!length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(����," & ints���� & ",'0')),'00') As Code FROM ҩƷ������"
                            Set rsMaxs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "-ҩƷ�����̱���")
                            strCodes = rsMaxs!Code
                            
                            ints���� = Len(strCodes)
                            strCodes = strCodes + 1
                            If ints���� >= Len(strCodes) Then
                                strCodes = String(ints���� - Len(strCodes), "0") & strCodes
                            End If
                            
                            gstrSQL = "ZL_ҩƷ������_INSERT('" & strCodes & "','" & strKey & "',zlSpellCode('" & strKey & "',10))"
                            
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                        End If
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntColԭ����) = rsProvider!����
                        mshBill.Text = rsProvider!����
                    End If
                End If
                OS.OpenIme
            Case mconIntCol����
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .Text = " "
                        .TextMatrix(.Row, mconIntCol����) = ""
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    Exit Sub
                End If
            Case mconIntCol��������
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "�Բ����������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        .Text = strKey
                        .TextMatrix(.Row, mconIntCol��������) = .Text
                        
                        '����Ч��
                        If Trim(.TextMatrix(.Row, mconIntColԭ����)) = "" Then
                            Exit Sub
                        End If
                        If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0) = "0" Then
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol��������) <> "" Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol��������))
                        ElseIf .TextMatrix(.Row, mconIntCol����) <> "" And Len(.TextMatrix(.Row, mconIntCol����)) = 8 Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                If IsNumeric(strxq) Then
                                    If Trim(.TextMatrix(.Row, mconIntColЧ��)) = "" Then
                                        strxq = TranNumToDate(strxq)
                                    Else
                                        Exit Sub
                                    End If
                                Else
                                    strxq = ""
                                End If
                            Else
                                strxq = ""
                            End If
                        End If
                        If Trim(strxq) = "" Then Exit Sub
                        
                        .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntColԭ����), "||")(0), strxq), "yyyy-mm-dd")
                        
                        If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                            '����Ϊ��Ч��
                            .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                        End If
                        
                        Call CheckLapse(.TextMatrix(.Row, mconIntColЧ��))
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "�Բ����������ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mconIntCol��������) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    
                    Cancel = True
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
                
            Case mconIntCol�ɱ���
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬳ɱ��۱���Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintCostDigit, , True)
                    .Text = strKey
                End If
                
                '��ʱ��ҩƷ�Ĵ���
                If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɱ���) And .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                    If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                        'Modified by zyb 2002-10-30
                        .Text = zlStr.FormatEx(strKey, mintCostDigit, , True)
                        
                        '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ���
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                            If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                            End If
                        Else
                            If mbln�Ӽ��� Then
                                If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  '���δ��ѡȡ�ϴ��ۼۣ��ҹ�ѡ���ֹ�¼��ӳ��ʲ����򵯳��ӳ��ʿ����û�ѡ��
                                    sngLeft = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                    sngTop = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                                    If sngTop + 1700 > Screen.Height Then
                                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 1700
                                    End If
                                    
                                    With PicInput
                                        .Top = sngTop
                                        .Left = sngLeft
                                        .Visible = True
                                    End With
                                    If Txt�Ӽ���.Text = "" Then Txt�Ӽ���.Text = "15.00000"
                                    .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(strKey), Val(Txt�Ӽ���) / 100, Val(strKey) * (1 + (Val(Txt�Ӽ���) / 100))), mintPriceDigit, , True)
                                    If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And Val(strKey) <> 0 Then
                                        Txt�Ӽ��� = zlStr.FormatEx(����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), Val(strKey)), 5, , True)
                                    End If
                                    Txt�Ӽ���.Tag = Txt�Ӽ���
                                    Txt�Ӽ���.SetFocus
                                End If
                            Else
                                If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
                                    If get�ֶμӳ��ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), strKey, dbl�ӳ���, dblTemp�ۼ�) = False Then
                                        .TxtSetFocus
                                        Cancel = True
                                        Exit Sub
                                    End If
                                Else
                                    dbl�ӳ��� = Val(Split(.TextMatrix(.Row, mconIntColԭ����), "||")(1))
                                    dblTemp�ۼ� = strKey * (1 + dbl�ӳ���)
                                End If
                                
                                If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then
                                    .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), strKey, dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                End If
                                
                                If .TextMatrix(.Row, mconIntCol����) <> "" Then
                                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                                End If
                            End If
                        End If
                    
                    Else
                        '���ۿ��ƣ�����ҩƷ�����¼��ĳɱ����Ƿ�����ۼ�
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol�ۼ�)) Then
                            If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ɱ���Ӧ���ۼ�(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�), mintPriceDigit, , True) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(.Row, mconIntCol�ۼ�)
                                .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                                .Text = strKey
'                                Cancel = True
'                                .TxtSetFocus
'                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                '���ý��
                If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɱ���) And .TextMatrix(.Row, mconIntCol����) <> "" Then
                    .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * strKey, mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * Val(.TextMatrix(.Row, mconIntCol�ۼ�)), mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                End If
                ��ʾ�ϼƽ��
                If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" And .TextMatrix(.Row, mconIntCol����ϵ��) <> "" Then
                    Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
                End If
            Case mconIntCol�ɱ����
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�Բ��𣬳ɱ�������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) * Val(.TextMatrix(.Row, mconIntCol����)) < 0 Then
                        MsgBox "�ɱ�������Ӧ����������һ�£�", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                '��ʽ�����
                If strKey <> "" Then
                    strKey = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                    .Text = strKey
                End If
                
                If strKey <> "" And strKey <> .TextMatrix(.Row, mconIntCol�ɱ����) Then
                    If .TextMatrix(.Row, mconIntCol����) <> "" Then
                        '���ۿ��ƣ�����ҩƷ�����ܵ����������Ϊ�ۼ۹̶����ۼ۽��Ҳ�̶���
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 0 And strKey <> .TextMatrix(.Row, mconIntCol�ۼ۽��) Then
                                MsgBox "�ö���ҩƷ���������۹���ģʽ�����ܵ��������", vbInformation + vbOKOnly, gstrSysName
                                strKey = .TextMatrix(.Row, mconIntCol�ۼ۽��)
                                .Text = strKey
                                Cancel = True
'                                .TxtSetFocus
                                Exit Sub
                            End If
                        Else
                            If mbln�Ӽ��� Then
                                'ȡ�øı�ɱ����ǰ�ļӼ���
                                mdbl�Ӽ��� = 15
                                If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And Val(.TextMatrix(.Row, mconIntCol�ɱ���)) <> 0 Then
                                    mdbl�Ӽ��� = ����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)))
                                End If
                            End If
                            
                            .Text = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                            .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(strKey / .TextMatrix(.Row, mconIntCol����), mintCostDigit, , True)
                            '��ʱ��ҩƷ�Ĵ���
                            If .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                                If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                                    '���ڴ��ڲ�������ȵĴ���,��Ҫ���ӳ��ʼ���,��˽�ָ�������ת���ɼӳ��ʼ��� ��ʽ���ӳ���=1/(1-�����)-1
                                    If mbln�Ӽ��� Then
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), (mdbl�Ӽ��� / 100), Val(.TextMatrix(.Row, mconIntCol�ɱ���)) * (1 + (mdbl�Ӽ��� / 100))), mintPriceDigit, , True)
                                        End If
                                        
                                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                                    Else
                                        If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
                                            If get�ֶμӳ��ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), dbl�ӳ���, dblTemp�ۼ�) = False Then
                                                .TxtSetFocus
                                                Cancel = True
                                                Exit Sub
                                            End If
                                        Else
                                            dbl�ӳ��� = Val(Split(.TextMatrix(.Row, mconIntColԭ����), "||")(1))
                                            dblTemp�ۼ� = .TextMatrix(.Row, mconIntCol�ɱ���) * (1 + dbl�ӳ���)
                                        End If
                                        
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                        End If
                                        
                                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(.Row, mconIntCol����)) <> 0 Then
                        .TextMatrix(.Row, mconIntCol�ɱ���) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol����)), mintCostDigit, , True)
                    End If
                    
                    '���ۿ��ƣ�����ҩƷ�����ܵ����������Ϊ�ۼ۹̶����ۼ۽��Ҳ�̶���
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                        .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(strKey / Val(.TextMatrix(.Row, mconIntCol����)), mintCostDigit, , True)
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                    End If
                    
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - strKey, mintMoneyDigit, , True)
                    .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(strKey, mintMoneyDigit, , True)
                    
                    Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
                End If
                ��ʾ�ϼƽ��
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
                    If Abs(Val(strKey)) = 0 Then
                        MsgBox "�Բ��������ľ���ֵ���������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    If mint�༭״̬ = 2 And Val(.TextMatrix(.Row, mconIntCol����)) <> 0 And .TextMatrix(.Row, mconIntCol�Ƿ�����) = "��" Then
                        If Not ��ͬ����(Val(strKey), Val(.TextMatrix(.Row, mconIntCol����))) Then
                            MsgBox "�Բ��������ķ���Ӧ����ԭ���������ķ���һ�£�", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    If Val(strKey) < 0 Then
                        If Not zlStr.IsHavePrivs(mstrPrivs, "��������") Then
                            MsgBox "�Բ�����û�и���������Ȩ�ޣ������䣡", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        If .TextMatrix(.Row, mconIntCol��������) = 1 Then
                            MsgBox "����ҩƷ����������⣬������", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                                        
                    strKey = zlStr.FormatEx(strKey, mintNumberDigit, , True)
                    .Text = strKey
                    
                    If Trim(.TextMatrix(.Row, mconIntCol�ɱ���)) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɱ���) * strKey, mintMoneyDigit, , True)
                        
                        '���ۿ���
                        If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                            '������������۹������������ۼ�
                        Else
                            'ʱ��ҩƷ�Ĵ���
                            If .TextMatrix(.Row, mconIntColԭ����) <> "" Then
                                If Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 Then
                                    'Modified by ZYB 2002-10-30
                                    If mbln�Ӽ��� Then
                                        mdbl�Ӽ��� = 15
                                        If Val(.TextMatrix(.Row, mconIntCol�ۼ�)) <> 0 And Val(.TextMatrix(.Row, mconIntCol�ɱ���)) <> 0 Then
                                            mdbl�Ӽ��� = ����ӳ���(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ۼ�)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)))
                                        End If
                                        
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), (mdbl�Ӽ��� / 100), Val(.TextMatrix(.Row, mconIntCol�ɱ���)) * (1 + (mdbl�Ӽ��� / 100))), mintPriceDigit, , True)
                                        End If
                                        
                                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * strKey, mintMoneyDigit, , True)
                                        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                                    Else
                                        If mintʱ�۷ֶμӳɷ�ʽ = 1 Then
                                            If get�ֶμӳ��ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����ϵ��)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), dbl�ӳ���, dblTemp�ۼ�) = False Then
                                                .TxtSetFocus
                                                Cancel = True
                                                Exit Sub
                                            End If
                                        Else
                                            dbl�ӳ��� = Split(.TextMatrix(.Row, mconIntColԭ����), "||")(1)
                                            dblTemp�ۼ� = .TextMatrix(.Row, mconIntCol�ɱ���) * (1 + dbl�ӳ���)
                                        End If
                                        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then  'û�й�ѡʱ��ȡ�ϴ��ۼ۲���
                                            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), dbl�ӳ���, dblTemp�ۼ�), mintPriceDigit, , True)
                                        End If
                                        
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                    
                    .TextMatrix(.Row, mconIntCol����) = strKey
                    Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
                End If
                ��ʾ�ϼƽ��
            Case mconIntCol��������
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
                    If Not ��ͬ����(Val(strKey), Val(.TextMatrix(.Row, mconIntCol����))) Then
                        MsgBox "�Բ��𣬳��������ķ���Ӧ����ԭ������һ�£�", vbInformation, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
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
                    
                    If .TextMatrix(.Row, mconIntCol�ɱ���) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ɱ����) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɱ���) * strKey, mintMoneyDigit, , True)
                    End If
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�) * strKey, mintMoneyDigit, , True)
                    End If
                    .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                    
                    .TextMatrix(.Row, mconIntCol��������) = strKey
                    Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconintCol���ۼ�)))
                End If
                ��ʾ�ϼƽ��
            Case mconIntCol��׼�ĺ�
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .Text = " "
                        .TextMatrix(.Row, mconIntCol��׼�ĺ�) = ""
                    End If
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
                    Exit Sub
                End If
            Case mconIntCol�ۼ�
                '������ۼ۲��ܴ���ָ�����ۼ�
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "�ۼ۱���Ϊ�����ͣ������䣡", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strKey = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ۼ�), mintPriceDigit, , True)
                
                '�ж���������ۼ���ָ�����ۼ�
                gstrSQL = "Select ָ�����ۼ� From ҩƷĿ¼ Where ҩƷID=[1] "
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(.Row, 0)))
                
                dblָ�����ۼ� = Round(rsTemp!ָ�����ۼ� * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), 5)
                strKey = Round(strKey, 5)
                If Val(strKey) > dblָ�����ۼ� Then
                    MsgBox "��������ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ��ۣ�ֻ��ʱ��ҩƷ�����޸��ۼ�
                If gtype_UserSysParms.P275_���۹���ģʽ = 2 And Split(.TextMatrix(.Row, mconIntColԭ����), "||")(2) = 1 And Val(strKey) <> Val(.TextMatrix(.Row, mconIntCol�ɱ���)) Then
                    If IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                        MsgBox "��ʱ��ҩƷ���������۹���ģʽ���ۼ�Ӧ�ͳɱ���(" & zlStr.FormatEx(.TextMatrix(.Row, mconIntCol�ɱ���), mintPriceDigit, , True) & ")��ȣ�", vbInformation + vbOKOnly, gstrSysName
                        strKey = .TextMatrix(.Row, mconIntCol�ɱ���)
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
                    End If
                End If
                
                .Text = zlStr.FormatEx(strKey, mintPriceDigit, , True)
                .TextMatrix(.Row, .Col) = .Text
                
                '������
                .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
                .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
            Case mconintCol���ۼ�
                '��������ۼ۲��ܴ���ָ�����ۼ�
                If Val(.TextMatrix(.Row, 0)) = 0 Then Exit Sub
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "���ۼ۱���Ϊ�����ͣ������䣡", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If .TxtVisible = False Then strKey = zlStr.FormatEx(.TextMatrix(.Row, mconintCol���ۼ�), gtype_UserDrugDigits.Digit_���ۼ�, , True)
                
                '�ж���������ۼ���ָ�����ۼ�
                gstrSQL = "Select ָ�����ۼ� From ҩƷĿ¼ Where ҩƷID=[1] "
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", Val(.TextMatrix(.Row, 0)))
                
                dblָ�����ۼ� = Round(rsTemp!ָ�����ۼ�, gtype_UserDrugDigits.Digit_���ۼ�)
                If strKey <> "" Then strKey = Round(strKey, gtype_UserDrugDigits.Digit_���ۼ�)
                If Val(strKey) > dblָ�����ۼ� Then
                    MsgBox "��������ۼ۲��ܴ���ָ�����ۼۣ�ָ�����ۼۣ���" & dblָ�����ۼ� & "��", vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                .Text = zlStr.FormatEx(strKey, gtype_UserDrugDigits.Digit_���ۼ�, , True)
                .TextMatrix(.Row, .Col) = .Text
                
                .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(Val(.TextMatrix(.Row, .Col)) * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), mintPriceDigit, , True)
                .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
                
                If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(.Row, 0))) = True Then
                    .TextMatrix(.Row, mconIntCol�ɱ���) = .TextMatrix(.Row, mconIntCol�ۼ�)
                    .TextMatrix(.Row, mconIntCol�ɱ����) = .TextMatrix(.Row, mconIntCol�ۼ۽��)
                End If
                
                .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
                
                Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.Text))
                Call ��ʾ�����
            Case mconIntCol���
                '�޴���
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol���) = ""
                        .Text = " "
                    End If
                    
                    If .TextMatrix(.Row, .Col) = "" Then
                        .TextMatrix(.Row, .Col) = " "
                    End If
'                    .Col = mconIntCol����
'                    Cancel = True
                    Exit Sub
                Else
                    Dim rs��� As New Recordset
                    
                    gstrSQL = "Select ����,����,���� From ҩƷ��� " _
                            & "Where upper(����) like [1] or Upper(����) like [2] or Upper(����) like [3] "
                    Set rs��� = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", strKey & "%")
                    
                    If rs���.EOF Then
                        .TextMatrix(.Row, mconIntCol���) = .Text
'                        .Col = mconIntCol����
'                        Cancel = True
                        Exit Sub
                    Else
                        If rs���.RecordCount = 1 Then
                            .TextMatrix(.Row, mconIntCol���) = rs���.Fields("����")
                            .Text = rs���.Fields("����")
                        Else
                            Set msh����.Recordset = rs���
                            With msh����
                                .Redraw = False
                                .Left = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                                .Top = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                                .Visible = True
                                .SetFocus
                                .ColWidth(0) = 800
                                .ColWidth(1) = 800
                                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                                .Row = 1
                                .Col = 0
                                .TopRow = 1
                                .ColSel = .Cols - 1
                                .Redraw = True
                                Cancel = True
                                Exit Sub
                            End With
                        End If
                    End If
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
Private Function SetColValue(ByVal intRow As Integer, ByVal lngҩƷID As Long, ByVal strҩƷ���� As String, _
    ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, ByVal str����ҩ��, _
    ByVal str��� As String, ByVal str���� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, _
    ByVal numָ�������� As Double, ByVal strԭ������ As String, _
    ByVal intԭЧ�� As Integer, dbl����ϵ�� As Double, _
    ByVal int�Ƿ��� As Integer, ByVal intҩ������ As Integer, ByVal dbl�ӳ��� As Double, ByVal str�������� As String, ByVal str�ۼ۵�λ As String, ByVal strԭ���� As String) As Boolean
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbl�ɱ��� As Double
    Dim rsPrice As New Recordset
    Dim lngDepartid As Long
    Dim strҩ�� As String
    Dim rsProvider As ADODB.Recordset
    Dim rsRecord As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim str������ As String
    Dim rs�ۼ� As ADODB.Recordset
    
    SetColValue = False
    lngDepartid = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    On Error GoTo errHandle
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        gstrSQL = "SELECT Nvl(a.���������,0) ���������,nvl(a.����,0) ����,Nvl(a.�б�ҩƷ,0) �б�ҩƷ,nvl(a.�ɱ���,0) �ɱ���,a.�ϴ���׼�ĺ�, a.��׼�ĺ�,a.�ϴβ��� ,b.����,a.ԭ����,a.�ϴ���������" & _
                " from ҩƷ��� a,�շ���ĿĿ¼ b where a.ҩƷid=b.id and ҩƷid=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ����]", lngҩƷID)
        
        dbl�ɱ��� = rsTemp!�ɱ���
        
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
        .TextMatrix(intRow, mconIntColԭ����) = strԭ����
        
        '���ء���׼�ĺš��������ڹ��򣬸��ݲ�������ȡ
        '���������ȴ��ϴ����ȡ
        '���أ�ֱ�Ӵӹ�����ȡ�ϴβ��أ����û������շ���Ŀ��ȡ���أ�û���������
        '��׼�ĺţ����ȴӹ�����ȡ�ϴ���׼�ĺţ����û����ӹ�����ȡ��׼�ĺţ���û��������׼�ĺ�
        '�������ڣ����ȴӹ�����ȡ�ϴ��������ڣ����û������
        '�ɱ��ۣ��ӹ�����ȡ�ɱ���
        
        '���������ȴ�����������ȡ
        '���أ����ȴӿ������������ȡ���أ����û������շ���Ŀ��ȡ���أ�û���������
        '��׼�ĺţ����ȴӿ������������ȡ��׼�ĺţ����û����ӹ�����ȡ��׼�ĺţ���û��������׼�ĺ�
        '�������ڣ����ȴӿ������������ȡ�������ڣ����û������
        '�ɱ��ۣ����ȴ�ҩƷ�������������ȡ�ϴγɱ��ۣ�û����ӹ�����ȡ�ɱ���
        If IIf(IsNull(rsTemp!�ϴβ���), "", rsTemp!�ϴβ���) <> "" Then
            .TextMatrix(intRow, mconIntCol����) = rsTemp!�ϴβ���
        Else
            .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        End If
        .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsTemp!�ϴ���������), "", rsTemp!�ϴ���������)
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsTemp!�ϴ���׼�ĺ�), IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�), rsTemp!�ϴ���׼�ĺ�)
        
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(num�ۼ� * dbl����ϵ��, mintPriceDigit, , True)
        .TextMatrix(intRow, mconIntColԭ������) = IIf(IsNull(strԭ������), "", strԭ������)
        .TextMatrix(intRow, mconIntColԭ����) = IIf(IsNull(intԭЧ��), "0", intԭЧ��) & "||" & dbl�ӳ��� & "||" & int�Ƿ��� & "||" & intҩ������
        .TextMatrix(intRow, mconIntCol����ϵ��) = dbl����ϵ��
        
        SetInputFormat intRow
        
        '��������
        Call GetҩƷ��������(intRow)
        
        '˵�����������ַ�������Ͳ����������Ŀ������������ٶȡ�
        '�������Բ�����Щ��ֱ���õ�һ��SQL���ʵ�֣�����������ҩƷ�Ͷ������ݿ���ɨ��һ��
        '0-���ȴ�ҩƷ���ȡ;1-���ȴ�ҩƷ���ȡ��
        If mintȡ�ϴγɱ��۷�ʽ = 0 Then
            If Val(.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                gstrSQL = "select �ϴβɹ��� as �ϴγɱ���,�ϴβ���,��׼�ĺ�,�ϴ��������� from ҩƷ��� where ����=1 and �ⷿid=[1] and ҩƷid=[2] " & _
                        " and nvl(����,0) =(select max(nvl(����,0)) from ҩƷ��� where ����=1 and �ⷿid=[1] )"
            Else
                gstrSQL = "select �ϴβɹ��� as �ϴγɱ���,�ϴβ���,��׼�ĺ�,�ϴ��������� from ҩƷ��� where ����=1 and �ⷿid=[1] and ҩƷid=[2]"
            End If
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ�ϴγɱ���]", lngDepartid, lngҩƷID)
            
            If Not rsPrice.EOF Then
                .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsPrice!�ϴβ���), IIf(IsNull(rsTemp!����), "", rsTemp!����), rsPrice!�ϴβ���)
                'mintʱ������ۼۼӳɷ�ʽ
                If nvl(rsPrice!�ϴγɱ���) = 0 Then
                    If dbl�ɱ��� >= 0 Then
                        .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(dbl�ɱ��� * dbl����ϵ��, mintCostDigit, , True)
                    End If
                Else
                    .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(rsPrice!�ϴγɱ��� * dbl����ϵ��, mintCostDigit, , True)
                End If
                .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsPrice!��׼�ĺ�), IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�), rsPrice!��׼�ĺ�)
                .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsPrice!�ϴ���������), "", Format(rsPrice!�ϴ���������, "yyyy-mm-dd"))
            Else
                .TextMatrix(intRow, mconIntCol��������) = ""
                If dbl�ɱ��� >= 0 Then
                    .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(dbl�ɱ��� * dbl����ϵ��, mintCostDigit, , True)
                End If
            End If
        Else
            If dbl�ɱ��� >= 0 Then
                .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(dbl�ɱ��� * dbl����ϵ��, mintCostDigit, , True)
            End If
        End If
        
        'ʱ��ҩƷ����
        If int�Ƿ��� = 1 Then
            '���ۿ��ƣ�ʱ��ҩƷ���ۼ�ֱ�ӵ��ڳɱ���
            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True Then
                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ɱ���), mintPriceDigit, , True)
                If .TextMatrix(intRow, mconIntCol����) <> "" Then
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(.Row, mconIntCol����) * .TextMatrix(.Row, mconIntCol�ۼ�), mintMoneyDigit, , True)
                End If
            Else
                If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� = 1 Then
                    gstrSQL = "select nvl(�ϴ��ۼ�,0) �ϴ��ۼ� from ҩƷ��� where ҩƷid=[1]"
                                     
                    Set rs�ۼ� = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�ۼ�", lngҩƷID)
                    If rs�ۼ�!�ϴ��ۼ� > 0 Then
                        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rs�ۼ�!�ϴ��ۼ� * dbl����ϵ��, mintPriceDigit, , True)
                    Else
                        .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), dbl�ӳ���, Val(.TextMatrix(intRow, mconIntCol�ɱ���)) * (1 + dbl�ӳ���)), mintPriceDigit, , True)
                    End If
                Else
                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), dbl�ӳ���, Val(.TextMatrix(intRow, mconIntCol�ɱ���)) * (1 + dbl�ӳ���)), mintPriceDigit, , True)
                End If
                
            End If
        Else
            '���ۿ��ƣ�����ҩƷ���ɱ���Ĭ�ϵ����ۼ�
            If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(Val(.TextMatrix(intRow, 0))) = True Then
                .TextMatrix(intRow, mconIntCol�ɱ���) = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�), mintPriceDigit, , True)
            End If
        End If
        
        If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
            If mintUnit <> mconint�ۼ۵�λ And Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                .TextMatrix(intRow, mconintCol���۵�λ) = str�ۼ۵�λ
            End If
        End If
        
        If .TextMatrix(intRow, mconIntCol����) <> "" Then
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
            Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol����), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
               .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
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

Private Sub SetInputFormat(ByVal intRow As Integer)
    If mblnEdit = False Then Exit Sub
    With mshBill
    
        If mint�༭״̬ = 1 Then
            .ColData(mconIntCol����) = 1
            .ColData(mconIntColԭ����) = 1
        End If
        
        If .TextMatrix(intRow, mconIntColԭ����) <> "" Then
            .ColData(mconIntColЧ��) = 2                '���������
            '�����ʱ��ҩƷ�������������ۼ�
            If Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2) = 1 Then
                .ColData(mconIntCol�ۼ�) = IIf(Getʱ��ҩƷֱ��ȷ���ۼ�, 4, 5)
            Else
                .ColData(mconIntCol�ۼ�) = 5
            End If
        Else
            .ColData(mconIntColЧ��) = 5
        End If
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            If mshBill.TextMatrix(intRow, mconIntColԭ����) <> "" Then
                mshBill.ColData(mconintCol���ۼ�) = 5
                If Val(Split(mshBill.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(mshBill.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                    mshBill.ColData(mconintCol���ۼ�) = 4
                End If
            End If
        End If
    End With
End Sub


Private Sub mshBill_LeaveCell(Row As Long, Col As Long)
    OS.OpenIme
End Sub

Private Sub mshBill_LostFocus()
    OS.OpenIme
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

Private Sub msh����_DblClick()
    msh����_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    On Error GoTo errHandle
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh����.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, mconIntCol����) = msh����.TextMatrix(msh����.Row, 2)
            
            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
            Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mconIntCol����), mshBill.TextMatrix(mshBill.Row, 0))
            If Not rsProvider.EOF Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
            Else
                mshBill.TextMatrix(mshBill.Row, mconIntCol��׼�ĺ�) = ""
            End If
            
            msh����.Visible = False
            .Col = mconIntCol����
            .SetFocus
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh����_LostFocus()
    If msh����.Visible Then
        msh����.Visible = False
    End If
End Sub

Private Function ValidData() As Boolean
    ValidData = False
    Dim intLop As Integer
    Dim rsStock As New Recordset
    Dim blnStock As Boolean
    
    On Error GoTo errHandle
    gstrSQL = "SELECT count(*)" _
              & " From ��������˵�� " _
             & " WHERE ((�������� LIKE '%ҩ��') " _
                  & "   OR (�������� LIKE '�Ƽ���')) " _
               & " AND ����id =[1]"
    Set rsStock = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���]", cboStock.ItemData(cboStock.ListIndex))
               
               
    If rsStock.Fields(0) > 0 Then
        blnStock = False
    Else
        blnStock = True
    End If
    
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
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol�ɱ���))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ���Ϊ���ˣ����飡", vbInformation, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ���
                        Exit Function
                    End If
                    
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol�ɱ����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ����
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
                    
                    If Split(.TextMatrix(intLop, mconIntColԭ����), "||")(0) <> "0" Then
                        If blnStock = True And (.TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntColЧ��) = "") Then
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
                    
                    '����ҩƷ����¼����غ�����
                    If Val(.TextMatrix(intLop, mconIntCol��������)) = 1 And (.TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntCol����) = "") Then
                        MsgBox "��" & intLop & "�е�ҩƷ�Ƿ���ҩƷ,������������̺�����" & vbCrLf & "��Ϣ���뵥���У�", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol����) = "" Then
                            .Col = mconIntCol����
                        Else
                            .Col = mconIntCol����
                        End If
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɱ���)) > 9999999999# Then
                        MsgBox "  ��" & intLop & "��ҩƷ�ĳɱ��۴��������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ���
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
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɱ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol�ɱ����
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
                                                            
                    '���۹�������Ƿ���ڲ��������۵�ҩƷ
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 Then
                        If IsPriceAdjustMod(Val(.TextMatrix(intLop, 0))) = True Then
                            If Val(.TextMatrix(intLop, mconIntCol�ɱ���)) <> Val(.TextMatrix(intLop, mconIntCol�ۼ�)) Then
                                MsgBox "��" & intLop & "��ҩƷ���������۹�������ⵥ���ۼۺͳɱ��۲�һ�£����ܽ���ҵ�����飡", vbInformation + vbOKOnly, gstrSysName
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
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngInOutTypeID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim strProducingArea As String
    Dim strOldProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
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
    Dim intRow As Integer
    Dim datTimeProduct As String
    Dim str��׼�ĺ� As String
    Dim n As Integer
    Dim m As Integer
    Dim lngҩƷID As Long
    Dim str��� As String
    Dim dbl���� As Double
    
    SaveCard = False
    On Error GoTo errHandle
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = Sys.GetNextNo(24, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        Me.txtNo.Tag = chrNo
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        lngInOutTypeID = cboType.ItemData(cboType.ListIndex)
        strBrief = Trim(txtժҪ.Text)
        strBooker = Trim(Txt������)
        datBookDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strAssessor = Trim(Txt�����)
        
'        gcnOracle.BeginTrans
        If mint�༭״̬ = 2 Or blnǿ�Ʊ��� Then        '�޸�
            gstrSQL = "zl_ҩƷ�������_Delete('" & mstr���ݺ� & "')"
            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            strBooker = Trim(Txt������)
            datBookDate = Format(Txt��������, "yyyy-mm-dd hh:mm:ss")
            strModifier = Trim(UserInfo.�û�����)
            datModifyDate = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = Trim(.TextMatrix(intRow, mconIntCol����))
                strOldProducingArea = Trim(.TextMatrix(intRow, mconIntColԭ����))
                strBatchNo = Trim(.TextMatrix(intRow, mconIntCol����))
                datTimeProduct = IIf(Trim(.TextMatrix(intRow, mconIntCol��������)) = "", "", Trim(.TextMatrix(intRow, mconIntCol��������)))
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntColЧ��)) = "", "", Trim(.TextMatrix(intRow, mconIntColЧ��)))
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                    '����ΪʧЧ��������
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = .TextMatrix(intRow, mconIntCol����) * .TextMatrix(intRow, mconIntCol����ϵ��)
                dblPurchasePrice = Round(.TextMatrix(intRow, mconIntCol�ɱ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblPurchaseMoney = .TextMatrix(intRow, mconIntCol�ɱ����)
                dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dblSaleMoney = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                dblMistakePrice = .TextMatrix(intRow, mconintCol���)
                
'                If Val(.TextMatrix(intRow, mconIntCol���)) = 0 Then
'                    lngSerial = intRow
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol���))
'                End If
                lngSerial = intRow
                
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)))
                str��� = Trim(.TextMatrix(intRow, mconIntCol���))
                
                If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 0 And mintUnit <> 4 Then
                    '����Ƕ���ҩƷ�����ۼ�ȡԭʼ�۸񱣴�
                    dblSalePrice = Get�ۼ�(Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1, lngDrugID, lngStockid, 0)
                                    
                    If gtype_UserSysParms.P275_���۹���ģʽ = 2 And IsPriceAdjustMod(lngDrugID) = True Then
                        '�����ʵ�����۹����ҩƷ���ɱ���ҲҪ���ۼ�һ��
                        dblPurchasePrice = dblSalePrice
                    End If
                End If
                
                'ʱ�۷���ҩƷ����
                If Val(Split(.TextMatrix(intRow, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intRow, mconIntCol��������)) = 1 Then
                    dblSalePrice = zlStr.FormatEx(.TextMatrix(intRow, mconintCol���ۼ�), gtype_UserDrugDigits.Digit_���ۼ�)
                    dblSaleMoney = .TextMatrix(intRow, mconintCol���۽��)
                    dblMistakePrice = .TextMatrix(intRow, mconintCol���۲��)
                    dbl���� = zlStr.FormatEx(Val(.TextMatrix(intRow, mconintCol���۽��)) - Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)), mintMoneyDigit, , True)
                End If
                
                gstrSQL = "zl_ҩƷ�������_INSERT("
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '���
                gstrSQL = gstrSQL & "," & lngSerial
                '�ⷿID
                gstrSQL = gstrSQL & "," & lngStockid
                '������ID
                gstrSQL = gstrSQL & "," & lngInOutTypeID
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngDrugID
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
                '������
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-MM-dd") & "','yyyy-mm-dd')")
                '��������
                gstrSQL = gstrSQL & "," & IIf(datTimeProduct = "", "Null", "to_date('" & Format(datTimeProduct, "yyyy-MM-dd") & "','yyyy-mm-dd')")
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '���
                gstrSQL = gstrSQL & ",'" & str��� & "'"
                '����
                gstrSQL = gstrSQL & "," & IIf(dbl���� <> 0, dbl����, "NULL")
                'ԭ����
                gstrSQL = gstrSQL & ",'" & strOldProducingArea & "'"
                '�޸���
                gstrSQL = gstrSQL & ",'" & strModifier & "'"
                '�޸�����
                gstrSQL = gstrSQL & "," & IIf(datModifyDate = "", "Null", "to_date('" & datModifyDate & "','yyyy-mm-dd HH24:MI:SS')")
                gstrSQL = gstrSQL & ")"

                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
            recSort.MoveNext
        Next
        
'        gcnOracle.CommitTrans
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
'    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
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
    Dim strժҪ As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim strҩƷ As String
    
    arrSql = Array()
    SaveStrike = False
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
        
        '�����
        strҩƷ = CheckNumStock(mshBill, cboStock.ItemData(cboStock.ListIndex), 0, mconIntCol����, mconIntCol��������, mconIntCol����ϵ��, 2, , mintNumberDigit)
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
                If .TextMatrix(intRow, mconIntCol��������) = .TextMatrix(intRow, mconIntCol����) Then
                    ��������_IN = Val(.TextMatrix(intRow, mconintCol��ʵ����))
                Else
                    ��������_IN = zlStr.FormatEx(.TextMatrix(intRow, mconIntCol��������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����, , True)
                End If
                
                strժҪ = txtժҪ.Text
                ���_IN = .TextMatrix(intRow, mconIntCol���)
                
                gstrSQL = "ZL_ҩƷ�������_STRIKE("
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
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strժҪ & "'"
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
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    Dim dblʱ�۷��� As Boolean
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol�ɱ����))
'            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            If .TextMatrix(intLop, mconIntColԭ����) <> "" Then
                If Val(Split(.TextMatrix(intLop, mconIntColԭ����), "||")(2)) = 1 And Val(.TextMatrix(intLop, mconIntCol��������)) = 1 Then
                    dblʱ�۷��� = True
                    Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconintCol���۽��))
                Else
                    Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
                End If
            Else
                Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
            End If
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & zlStr.FormatEx(curTotal, mintMoneyDigit, , True)
    
    If dblʱ�۷��� = True Then
        lblSalePrice.Caption = "�ۼ۽��(ʱ�۷��������۽��)�ϼƣ�" & zlStr.FormatEx(Cur���ʽ��, mintMoneyDigit, , True)
        lblDifference.Caption = "���(ʱ�۷��������۲��)�ϼƣ�" & zlStr.FormatEx(Cur���ʲ��, mintMoneyDigit, , True)
    Else
        lblDifference.Caption = "��ۺϼƣ�" & zlStr.FormatEx(Cur���ʲ��, mintMoneyDigit, , True)
        lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & zlStr.FormatEx(Cur���ʽ��, mintMoneyDigit, , True)
    End If
End Sub

Private Sub ��ʾ�����()
    Dim RecTmp As New ADODB.Recordset
    Dim Dbl���� As Double
    Dim str��λ As String
    Dim intID As Long
    Dim strUnit As String
    Dim strQuantity As String
    
    On Error GoTo errHandle
    If mshBill.TextMatrix(mshBill.Row, mconIntColҩ��) = "" Then
        staThis.Panels(2).Text = ""
        Exit Sub
    End If
    If mshBill.TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
    intID = mshBill.TextMatrix(mshBill.Row, 0)

    Select Case mintUnit
        Case mconint�ۼ۵�λ
            strUnit = "C.���㵥λ"
            strQuantity = "�������� "
        Case mconint���ﵥλ
            strUnit = "B.���ﵥλ"
            strQuantity = "��������/�����װ "
        Case mconintסԺ��λ
            strUnit = "B.סԺ��λ"
            strQuantity = "��������/סԺ��װ "
        Case mconintҩ�ⵥλ
            strUnit = "B.ҩ�ⵥλ"
            strQuantity = "��������/ҩ���װ "
    End Select
    
    gstrSQL = "Select b.ҩƷID," & strUnit & " as ��λ, Sum(" & strQuantity & ") as ���� " & _
        " From ҩƷ��� a,ҩƷ��� b,�շ���ĿĿ¼ C " & _
        " Where a.����=1 and a.ҩƷid=b.ҩƷid And B.ҩƷID=C.ID " & _
        " And ��������<>0 And �ⷿID=[1] and b.ҩƷID=[2] " & _
        " Group by b.ҩƷID," & strUnit
    Set RecTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", cboStock.ItemData(cboStock.ListIndex), intID)
    
    With RecTmp
        If .EOF Then
            staThis.Panels(2).Text = ""
            Exit Sub
        End If
        Dbl���� = IIf(IsNull(!����), 0, !����)
        
        staThis.Panels(2).Text = "��ҩƷ��ǰ�����Ϊ[" & zlStr.FormatEx(Dbl����, mintNumberDigit, , True) & "]" & !��λ
        
    End With
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
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1302", "zl8_bill_1302"), mint��¼״̬, int��λϵ��, 1302, "ҩƷ������ⵥ", strNo
End Sub


'ȡָ�������۶��۵�λ������ֵ��ȱʡΪ0-���ۼ۵�λ���ۣ���ѡΪ1-��ҩ�ⵥλ���ۣ�
Private Function GetUnit() As Integer
    GetUnit = gtype_UserSysParms.P29_ָ�������۶��۵�λ
End Function

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Call zlDataBase.OpenRecordset(rsBatchNolen, gstrSQL, "ȡ�ֶγ���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub PicInput_LostFocus()
    Dim strActive As String
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDYES,CMDNO,TXT�Ӽ���", strActive) <> 0 Then
        Exit Sub
    Else
        If strActive = "MSHBILL" Then
            If mshBill.Col = mconIntCol�ɱ��� Or mshBill.Col = mconIntCol�ɱ���� Then Exit Sub
        End If
    End If
    PicInput.Visible = False
End Sub

Private Sub Txt�Ӽ���_GotFocus()
    Txt�Ӽ���.SelStart = 0
    Txt�Ӽ���.SelLength = Len(Txt�Ӽ���)
End Sub

Private Sub Txt�Ӽ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call cmdYes_Click
End Sub

Private Sub Txt�Ӽ���_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub

Private Sub Txt�Ӽ���_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub cmdYes_Click()
    If Val(Txt�Ӽ���) > 9900 Or Val(Txt�Ӽ���) < 0 Then
        MsgBox "������Ϸ��ļӳ��ʣ���0-9900��", vbInformation, gstrSysName
        Txt�Ӽ���.SetFocus
        Exit Sub
    End If
    
    With mshBill
        '���¼������ۼۡ����
        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then
            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), Val(Txt�Ӽ���) / 100, Val(.TextMatrix(.Row, mconIntCol�ɱ���)) * (1 + (Val(Txt�Ӽ���) / 100))), mintPriceDigit, , True)
        End If
        
        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
        
        Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
    End With
    
    PicInput.Visible = False
    mshBill.SetFocus
End Sub

Private Sub CmdYes_LostFocus()
    Call PicInput_LostFocus
End Sub

Private Sub cmdNo_Click()
    With mshBill
        mdbl�Ӽ��� = Val(Txt�Ӽ���.Tag)
        
        '���¼������ۼۡ����
        If gtype_UserSysParms.P183_ʱ��ȡ�ϴ��ۼ� <> 1 Then
            .TextMatrix(.Row, mconIntCol�ۼ�) = zlStr.FormatEx(ʱ��ҩƷ���ۼ�(Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol�ɱ���)), mdbl�Ӽ��� / 100, Val(.TextMatrix(.Row, mconIntCol�ɱ���)) * (1 + (mdbl�Ӽ��� / 100))), mintPriceDigit, , True)
        End If
        
        .TextMatrix(.Row, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(.Row, mconIntCol�ۼ�)) * Val(.TextMatrix(.Row, mconIntCol����)), mintMoneyDigit, , True)
        .TextMatrix(.Row, mconintCol���) = zlStr.FormatEx(IIf(.TextMatrix(.Row, mconIntCol�ۼ۽��) = "", 0, .TextMatrix(.Row, mconIntCol�ۼ۽��)) - IIf(.TextMatrix(.Row, mconIntCol�ɱ����) = "", 0, .TextMatrix(.Row, mconIntCol�ɱ����)), mintMoneyDigit, , True)
        
        Call Setʱ�۷���ҩƷ���ۼ�(.Row, Val(.TextMatrix(.Row, mconIntCol�ۼ�)) / Val(.TextMatrix(.Row, mconIntCol����ϵ��)))
    End With
    PicInput.Visible = False
End Sub

Private Sub CmdNO_LostFocus()
    Call PicInput_LostFocus
End Sub

'ȡʱ��ҩƷ���ʱ���Ƿ��������Ӽ���
Private Function Get�Ӽ���() As Boolean
    Get�Ӽ��� = (gtype_UserSysParms.P54_ʱ��ҩƷ�ԼӼ������ = 1)
End Function

Private Function Getʱ��ҩƷֱ��ȷ���ۼ�() As Boolean
    Getʱ��ҩƷֱ��ȷ���ۼ� = (gtype_UserSysParms.P76_ʱ��ҩƷֱ��ȷ���ۼ� = 1)
End Function
Private Sub GetSysParm()
    mbln�¿������� = (gtype_UserSysParms.P96_ҩƷ��¿��ÿ�� = 1)
End Sub
Private Function ʱ��ҩƷ���ۼ�(ByVal lngҩƷID As Long, ByVal sin�ɱ��� As Double, ByVal sin�ӳ��� As Double, ByVal sin�ۼ� As Double, Optional ByVal lngLastRow As Long = -1) As Double
    Dim sin���ۼ� As Double, sinָ�����ۼ� As Double, sin��������� As Double
    Dim rsTemp As New ADODB.Recordset
    Dim sin������� As Double
    'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɱ���*(1+�ӳ���)
    '��Ϊ:�ɱ���*(1+�ӳ���)+(ָ�����ۼ�-�ɱ���*(1+�ӳ���))*(1-���������)
    '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɱ���*(1+�ӳ���))*(1-���������)
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    
    On Error GoTo errHandle
    gstrSQL = "Select ָ�����ۼ�,Nvl(���������,100) ��������� From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������

    ʱ��ҩƷ���ۼ� = 0
    If sin��������� = 100 Then
        ʱ��ҩƷ���ۼ� = sin�ۼ�
        Exit Function
    End If
    
    sin���ۼ� = sin�ɱ��� * (1 + sin�ӳ���)
    If sin���ۼ� / Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��)) >= sinָ�����ۼ� Then
        ʱ��ҩƷ���ۼ� = sin�ۼ�
        Exit Function
    End If
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��))
    sin������� = (sinָ�����ۼ� - sin���ۼ�) * (1 - sin��������� / 100)
    
    ʱ��ҩƷ���ۼ� = IIf(sin�ۼ� + sin������� > sinָ�����ۼ�, sinָ�����ۼ�, sin�ۼ� + sin�������)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ����ӳ���(ByVal lngҩƷID As Long, ByVal sin���ۼ� As Double, ByVal sin�ɱ��� As Double) As Double
    Dim sinָ�����ۼ� As Double, sin��������� As Double
    Dim rsTemp As New ADODB.Recordset
    '�������ۼ۷���ɱ���,����ʱ��ҩƷ��ʽ�ı仯,����ԭ������ӳ��ʵĹ�ʽ��Ч,�����¼���
    'ԭ��ʽ:(���ۼ�/�ɱ���-1)*100
    '�ֹ�ʽ������:�������ۼ��ǰ��ӳ����������,�ټ������������ǲ��ֽ��,���ʵ�ʰ��ӳ�����������ۼ�=ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    '������ԭ��ʽ���ʵ�ʵļӳ���
    ����ӳ��� = 0.15
    On Error GoTo errHandle
    gstrSQL = " Select ָ�����ۼ�,Nvl(���������,100) ���������,Nvl(�Ƿ���,0) ʱ�� " & _
              " From ҩƷ��� A,�շ���ĿĿ¼ B Where A.ҩƷID=B.ID AND A.ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sin��������� = rsTemp!���������
    If rsTemp!ʱ�� = 0 Then Exit Function
    
    'ָ�����ۼ�-(ָ�����ۼ�-���ۼ�)/���������
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(mshBill.Row, mconIntCol����ϵ��))
    If sin��������� <> 100 And sin��������� > 0 Then
        sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�) / sin��������� * 100
    Else
        sin���ۼ� = sinָ�����ۼ� - (sinָ�����ۼ� - sin���ۼ�)
    End If
    ����ӳ��� = (Val(sin���ۼ�) / Val(sin�ɱ���) - 1) * 100
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function У�����ۼ�(ByVal sin���ۼ� As Double, Optional ByVal lngLastRow As Long = -1) As Double
    '�õ�����ǰ��λϵ�����������ָ�����ۼۣ����ʱ��ҩƷ������������ۼ۴���ָ�����ۼۣ���ָ�����ۼ�Ϊ׼
    Dim sinָ�����ۼ� As Double
    Dim rsTemp As New ADODB.Recordset
    
    If lngLastRow = -1 Then lngLastRow = mshBill.Row
    
    On Error GoTo errHandle
    gstrSQL = " Select ָ�����ۼ�,Nvl(���������,100) ��������� " & _
              " From ҩƷ��� Where ҩƷID=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", Val(mshBill.TextMatrix(lngLastRow, 0)))
    
    sinָ�����ۼ� = rsTemp!ָ�����ۼ�
    sinָ�����ۼ� = sinָ�����ۼ� * Val(mshBill.TextMatrix(lngLastRow, mconIntCol����ϵ��))
    
    У�����ۼ� = IIf(sin���ۼ� > sinָ�����ۼ�, sinָ�����ۼ�, sin���ۼ�)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function get�ֶμӳ��ۼ�(ByVal lngҩƷID As Long, ByVal lng����ϵ�� As Long, ByVal dbl�ɱ��� As Double, ByRef dblR�ӳ��� As Double, ByRef dbl�ۼ� As Double) As Boolean
    '����:������ʱ��ҩƷ�ֶμӳ����󣬸��ݳɱ��ۼ������Ӧ���ۼ�
    '�ۼۼ��㹫ʽ�������۸���2000Ԫ/֧��ƿ��У���2000Ԫ�����µ�ҩƷ��������ۼ۸�=ʵ�ʹ����ۡ���1+����ʣ�+��۶
    '               �����۸���2000Ԫ/֧��ƿ��У�����2000Ԫ�����ϵ�ҩƷ��������ۼ۸� = ʵ�ʹ����� + ��۶�˶��Ѿ��������������ã�

    '�������ɱ���
    Dim dbl�ӳ��� As Double
    Dim dbl��۶� As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    dbl�ӳ��� = 0
    dbl��۶� = 0
    
    gstrSQL = "select ��� from  �շ���ĿĿ¼ a where a.id=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��ҩƷ���ʷ���", lngҩƷID)
    If rsTemp!��� = 7 Then
        mrs�ֶμӳ�.Filter = "����=1"
    Else
        mrs�ֶμӳ�.Filter = "����=0"
    End If
      
    If mrs�ֶμӳ�.RecordCount <> 0 Then
        mrs�ֶμӳ�.MoveFirst
        Do While Not mrs�ֶμӳ�.EOF
            With mrs�ֶμӳ�
                If dbl�ɱ��� > !��ͼ� And dbl�ɱ��� <= !��߼� Then
                    dbl�ӳ��� = IIf(IsNull(!�ӳ���), 0, !�ӳ���) / 100
                    dblR�ӳ��� = dbl�ӳ���
                    dbl��۶� = IIf(IsNull(!��۶�), 0, !��۶�)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs�ֶμӳ�.MoveNext
        Loop
    End If
    
    If blnData = False Then
        If rsTemp!��� = 7 Then
            MsgBox "����ҩ��δ���ý���Ϊ��" & dbl�ɱ��� & " " & "�ķֶμӳ����ݣ��뵽ҩƷĿ¼�����зֶμӳ������ã�", vbInformation, gstrSysName
        Else
            MsgBox "����ҩ/��ҩ��δ���ý���Ϊ��" & dbl�ɱ��� & " " & "�ķֶμӳ����ݣ��뵽ҩƷĿ¼�����зֶμӳ������ã�", vbInformation, gstrSysName
        End If
        get�ֶμӳ��ۼ� = False
    End If
    
    dbl�ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���) + dbl��۶�
    
    Set rsTemp = Nothing
    gstrSQL = "Select ָ�����ۼ� From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    If rsTemp!ָ�����ۼ� * lng����ϵ�� < dbl�ۼ� Then
        dbl�ۼ� = rsTemp!ָ�����ۼ� * lng����ϵ��
    End If
    
    get�ֶμӳ��ۼ� = True
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ����ۼ�() As Boolean
    '���ܣ��⹺����ʱ���ж϶���ҩƷ�Ƿ��������ۼۣ������޸ĺ���ʾ
    Dim strMsg As String '������ʾ��Ϣ
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    
    On Error GoTo errHandle
    
    ����ۼ� = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" Then
                
                If Val(Split(.TextMatrix(i, mconIntColԭ����), "||")(2)) = 0 Then '�ж϶���

                    dbl���ۼ� = zlStr.FormatEx(Get�ۼ�(False, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol����))) * Val(.TextMatrix(i, mconIntCol����ϵ��)), mintPriceDigit)
                    
                    If .TextMatrix(i, mconIntCol�ۼ�) <> dbl���ۼ� Then
                        intSum = intSum + 1 '��¼�����˼�������
                        
                        dbl�ɱ��� = Val(.TextMatrix(i, mconIntCol�ɱ���))
                        Dbl���� = Val(.TextMatrix(i, mconIntCol����))
                        dbl�ɱ���� = dbl�ɱ��� * Dbl����
                        dbl���۽�� = dbl���ۼ� * Dbl����
                        dbl��� = dbl���۽�� - dbl�ɱ����
                        
                        '�����ۼ��������
                        .TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, mintPriceDigit, , True)
                        .TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(dbl���۽��, mintMoneyDigit, , True)
                        .TextMatrix(i, mconintCol���) = zlStr.FormatEx(dbl���, mintMoneyDigit, , True)
                        
                    End If
                End If
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "�м�¼δʹ�������ۼۣ��������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
            ����ۼ� = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetNextEnableCol(ByVal intCurrCol As Integer) As Integer
    '������һ���ɼ������õ��к�
    Dim n As Integer
    Dim intNextCol As Integer

    If intCurrCol > mshBill.Cols Or intCurrCol + 1 >= mintLastCol Then
        GetNextEnableCol = mintLastCol
        Exit Function
    End If
    
    With mshBill
        For n = intCurrCol + 1 To .Cols - 1
            If .ColWidth(n) > 0 And .ColData(n) <> 5 Then
                intNextCol = n
                Exit For
            End If
        Next
    End With
    
    GetNextEnableCol = IIf(intNextCol = 0, mintLastCol, intNextCol)
End Function

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo errHandle
    Dim rsTmp As New ADODB.Recordset
     
    gstrSQL = "Select t.�ϴβ��� as ������, t.ԭ���� as ԭ���� From ҩƷ��� T Where Rownum < 1"
    Call zlDataBase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng�����̳��� = rsTmp.Fields("������").DefinedSize
    mlngԭ���س��� = rsTmp.Fields("ԭ����").DefinedSize
   
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
