VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmTransferCard 
   Caption         =   "�����ƿⵥ"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11130
   Icon            =   "frmTransferCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11130
   StartUpPosition =   2  '��Ļ����
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2175
      Left            =   2520
      TabIndex        =   35
      Top             =   1200
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
   Begin VB.CommandButton cmdRequestTransfer 
      Caption         =   "���깺���ƿ�(&T)"
      Height          =   350
      Left            =   3840
      TabIndex        =   34
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdExpend 
      Caption         =   "�Զ��ֽ�(&A)"
      Height          =   350
      Left            =   4950
      TabIndex        =   7
      Top             =   5490
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6180
      TabIndex        =   30
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7500
      TabIndex        =   29
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   12
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7560
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   13
      Top             =   0
      Width           =   11715
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "������ʵ�:F3"
         Top             =   150
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
         Height          =   360
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         IMEMode         =   2  'OFF
         Left            =   9930
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   165
         Width           =   1425
      End
      Begin VB.ComboBox cboEnterStock 
         Height          =   300
         Left            =   9240
         TabIndex        =   3
         Text            =   "cboEnterStock"
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   4
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
         TabIndex        =   6
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Text            =   "cboStock"
         Top             =   585
         Width           =   2745
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
         Left            =   7950
         TabIndex        =   23
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
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
         Width           =   915
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
         TabIndex        =   19
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
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "���������ƿⵥ"
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
         TabIndex        =   18
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƴ��ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ⷿ(&I)"
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
            Picture         =   "frmTransferCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1000
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
            Picture         =   "frmTransferCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   28
      Top             =   6492
      Width           =   11124
      _ExtentX        =   19632
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTransferCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13282
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":3080
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
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6000
      TabIndex        =   8
      Top             =   5025
      Width           =   1335
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
End
Attribute VB_Name = "frmTransferCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5,6-����,10-����,11-����ⵥ��ȡ����

Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mbln���쵥 As Boolean               '�Ƿ������쵥�������������ִ���Զ��ֽ�Ĺ���
Private mbln��ȷ���� As Boolean             '�Ƿ���ȷ���Σ��������쵥��Ч
Private mbln�ƿ���ȷ���� As Boolean         '�Ƿ���ȷ���Σ������ƿⵥ��Ч

Private mint����� As Integer             '��ʾ���ĳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrEnterSQL As String
Private mblnNoClick As Long
Private mblnCostView As Boolean                 '�鿴�ɱ��� true-����鿴 false-������鿴
Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������
Private Const mstrCaption As String = "�����ƿⵥ"
Private mstr�˲��� As String                '���쵥��ʹ�ã���¼����˲���
Private mstr�˲����� As String              '���쵥��ʹ�ã���¼����˲�����
Private mbln����˲� As Boolean             '���쵥��ʹ�ã���¼�����Ƿ���Ҫ�˲����� true-��Ҫ false-����Ҫ

Private mbln�����������Ų��ؿ��� As Boolean  '�Ƿ�������������Ų����Ƿ�¼��

Private mstrRequestNO As String     '���깺���ƿ�NO ���մ��������깺����ʽ�ƿ⣬�������깺���ƿ�
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private mbln����ʾ�п������  As Boolean

Dim mstrPrivs As String                     'Ȩ��
Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��

Private mstrTime_Start As String                        '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��
Private mblnFirst As Boolean
Private mint�ƿ⴦������ As Integer                    '1-��Ҫ��ҩ�����͡�������һ����  0-����Ҫ��һ����
Private mstr��ⵥ�� As String
Private mstr�ظ����� As String '��¼�ظ�������

Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

'=========================================================================================
Private Const mlngModule = 1716

Private mbln��������    As Boolean          '����ʱ���ݺ��ۼ�1
Private mintUnit  As Integer                '��ʾ��λ:0-ɢװ��λ,1-��װ��λ
Private mint������ʽ As Integer             '0������������ʽ��1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��

Private Enum mBillCol
     C_�к� = 1
     C_���� = 2
     c_��� = 3
     c_��� = 4
     C_�ⷿ���� = 5
     C_���Ч�� = 6
     C_�������� = 7
     C_ָ������� = 8
     C_ʵ�ʽ�� = 9
     C_ʵ�ʲ�� = 10
     C_����ϵ�� = 11
     c_���� = 12
     C_���� = 13
     C_��׼�ĺ� = 14
     c_��λ = 15
     c_���� = 16
     C_Ч�� = 17
     C_һ���Բ��� = 18
     C_���Ч�� = 19
     C_������� = 20
     C_���ʧЧ�� = 21
     C_��д���� = 22
     C_ʵ������ = 23
     c_ԭʼ���� = 24
     C_�ɹ��� = 25
     C_�ɹ���� = 26
     C_�ۼ� = 27
'     C_�ۼ۽�� = 28
'
'     C_��� = 29
End Enum
Private mconintcol�ۼ۽�� As Integer
Private mconintcol��� As Integer

Private Const mBillCols  As Integer = 30              '������
Private mlng����ⷿ As Long
Private mlngPreEnterId As Long      '�ϴ�����ⷿ
Private mlngPreStockId As Long  '�ϴ��Ƴ��ⷿ

Private Function Auto�����ƿ�����() As Boolean
    '�Զ������ƿ����� 1������ 2������ 3������
    
    On Error GoTo ErrHandle
    
    If Not ��鵥��(19, txtNO.Tag, False) And Not mblnUpdate Then
        MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
        Call RefreshBill
        mblnUpdate = True
        Exit Function
    End If
        
    If Not ���ϵ������(Txt������.Caption) Then Exit Function
    
    '2-
    If Not ValidData Then Exit Function
    If Not CheckStock Then Exit Function
    
    '��ɾ�����쵥�������ݵ�ǰ���ݲ����ƿⵥ
    If Not SaveCard(True) Then
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    '����
    gstrSQL = "zl_�����ƿ�_Prepare('" & txtNO.Tag & "','" & UserInfo.�û��� & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
                    
    '���ͣ��³���ⷿ�Ĳ��Ͽ��ÿ�棩
    gstrSQL = "zl_�����ƿ�_Prepare('" & txtNO.Tag & "')"
    zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
       
    '3-
    If SaveCheck() = True Then
        If IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '��ӡ
            If zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
            End If
        End If
        Unload Me
    Else
        GoTo ErrHandle
    End If
    
    Auto�����ƿ����� = True
    Exit Function
ErrHandle:
    Auto�����ƿ����� = False
End Function

'=========================================================================================


'�������������
Private Function GetDepend() As Boolean
    Dim strMsg As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    GetDepend = False
    With rsTemp
        '��������������Ƿ�����
        strMsg = "û�����������ƿ����⼰�����������������������ã�"
        
        gstrSQL = "" & _
            "   SELECT B.Id,B.ϵ�� " & _
            "   FROM ҩƷ�������� A, ҩƷ������ B " & _
            "   Where A.���id = B.ID  AND A.���� = 34"
            
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "�����ƿ����"
        
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "ϵ��=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û�����������ƿ������������������������ã�"
            GoTo ErrHand
        End If
        .Filter = "ϵ��=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û�����������ƿ�ĳ����������������������ã�"
            GoTo ErrHand
        End If
        .Filter = 0
        .Close
    End With
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsTemp.Close
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(frmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
    Optional int��¼״̬ As Integer = 1, Optional strPrivs As String, Optional blnSuccess As Boolean = False, Optional int������ʽ As Integer = 0)
    Dim strReg As String
    
    mblnSave = False
    mblnSuccess = False
    
    mstr��ⵥ�� = ""
    mstr���ݺ� = ""
    If int�༭״̬ = 11 Then
        mstr��ⵥ�� = str���ݺ�
    Else
        mstr���ݺ� = str���ݺ�
    End If
    
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mint������ʽ = int������ʽ
    
    mblnSuccess = blnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    If Not GetDepend Then Exit Sub
    mblnCostView = zlStr.IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    'û�гɱ���Ȩ�޽��ۼ۽��Ͳ��˳��һ�£���ֹ���ɼ���������ϳ���
    mconintcol�ۼ۽�� = IIf(mblnCostView = False, 29, 28)
    mconintcol��� = IIf(mblnCostView = False, 28, 29)
    
    Call GetRegInFor(g˽��ģ��, "�����ƿ����", "���ݺ��ۼ�", strReg)
    mbln�������� = IIf(strReg = "", True, Val(strReg) = 1)
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 11 Then
        
        mblnEdit = True

        txtNO.Locked = True
        txtNO.TabStop = True

        txtNO = mstr���ݺ�
        txtNO.Tag = txtNO.Text
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If InStr(mstrPrivs, "���ݴ�ӡ") = 0 Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint�༭״̬ = 6 Then
        mblnEdit = False
        If mint������ʽ = 0 Then '��������
            CmdSave.Caption = "����(&O)"
        ElseIf mint������ʽ = 1 Then    '�������
            CmdSave.Caption = "�������(&O)"
        ElseIf mint������ʽ = 2 Then    '��������������
            CmdSave.Caption = "��˳���(&O)"
        End If
        If mint������ʽ = 2 Then
            cmdAllSel.Visible = False
            cmdAllCls.Visible = False
        Else
            cmdAllSel.Visible = True
            cmdAllCls.Visible = True
        End If
        
    ElseIf mint�༭״̬ = 10 Then
        mblnEdit = False
        CmdSave.Caption = "����(&S)"
        CmdSave.Visible = True
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, frmMain
    blnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboEnterStock_Click()
    If mblnNoClick Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex)
End Sub

Private Sub cboEnterStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If cboEnterStock.ListCount = 0 Then Call zlControl.ControlSetFocus(mshBill): Exit Sub
    
    If cboEnterStock.ListIndex >= 0 Then
        If mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
            OS.PressKey vbKeyTab
            'Call zlControl.ControlSetFocus(mshBill, True)
            Exit Sub
        End If
    End If
    If Select����ѡ����(Me, cboEnterStock, Trim(cboEnterStock.Text), "", False, mstrEnterSQL) = False Then
        Exit Sub
    End If
    If cboEnterStock.ListIndex >= 0 Then
        mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    End If
End Sub

Private Sub cboEnterStock_LostFocus()
    Dim i As Long
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboEnterStock.ListIndex < 0 Then
        For i = 0 To cboEnterStock.ListCount - 1
            If mlngPreEnterId = cboEnterStock.ItemData(i) Then
                mblnNoClick = True
                cboEnterStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cboEnterStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboEnterStock
        If .ListCount = 0 Then Exit Sub
        If .ListIndex < 0 Then Exit Sub
        If .ListIndex <> Val(.Tag) Then
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("����ı�����ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ��������" & vbCrLf & "��Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�������ĵ�λ�ı�
                    cboEnterStock.Tag = .ListIndex
                    mshBill.ClearBill
                Else
                    .ListIndex = Val(.Tag)
                End If
            Else
                .Tag = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsTemp As New ADODB.Recordset
    If mblnNoClick Then Exit Sub
    If cboEnterStock.ListIndex >= 0 Then mlngPreEnterId = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    
    '��鲢װ������ⷿ
    err = 0: On Error Resume Next
    Set rsTemp = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), mstrCaption, True, mstrEnterSQL, 1716)
    With rsTemp
        cboEnterStock.Clear
        Do While Not .EOF
            cboEnterStock.AddItem !����
            cboEnterStock.ItemData(cboEnterStock.NewIndex) = !Id
            If mint�༭״̬ = 11 Then
                If Val(zlStr.NVL(!Id)) = mfrmMain.cboEnterStock.ItemData(mfrmMain.cboEnterStock.ListIndex) Then
                    cboEnterStock.ListIndex = cboEnterStock.NewIndex
                End If
            End If
            .MoveNext
        Loop
        If cboEnterStock.ListIndex < 0 Then cboEnterStock.ListIndex = 0
        If mint�༭״̬ <> 11 Then
            If cboEnterStock.ListCount <> 0 Then cboEnterStock.ListIndex = Val(cboEnterStock.Tag)
        End If
    End With
    
    mint�ƿ⴦������ = IIf(Val(zlDatabase.GetPara("�ƿ�����", glngSys, mlngModule, "0", , , , cboStock.ItemData(cboStock.ListIndex))) = 1, 1, 0)
    
    mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then OS.PressKey vbKeyTab: Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If mlngPreStockId = cboStock.ItemData(cboStock.ListIndex) Then
           OS.PressKey vbKeyTab
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), "V,K,W", Not zlStr.IsHavePrivs(mstrPrivs, "���пⷿ")) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        mlngPreStockId = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim i As Integer
        Dim blnreturn As Boolean
        blnreturn = False
        cboStock_Validate blnreturn
        If blnreturn = True Then Exit Sub
        
        OS.PressKey (vbKeyTab)
    End If
    
End Sub

Private Sub cboEnterStock_KeyPress(KeyAscii As Integer)
    Dim blnreturn As Boolean

    If KeyAscii <> 13 Then Exit Sub
    blnreturn = False
    cboEnterStock_Validate blnreturn
    If blnreturn = True Then Exit Sub

    With mshBill
        .Row = 1
        .Col = mBillCol.C_����
    End With
    zlControl.ControlSetFocus mshBill, True
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
    If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If mlngPreStockId = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboStock
        If .ListIndex < 0 Then Exit Sub
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.Rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.Rows Then
                If MsgBox("����ı��Ƴ��ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ��" & vbCrLf & "��Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�������ĵ�λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
        
    End With
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
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mBillCol.C_ʵ������) = Format(0, mFMT.FM_����)
                .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mconintcol�ۼ۽��) = Format(0, mFMT.FM_���)
                .TextMatrix(intRow, mconintcol���) = Format(0, mFMT.FM_���)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mBillCol.C_ʵ������) = .TextMatrix(intRow, mBillCol.C_��д����)
                .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(.TextMatrix(intRow, mBillCol.C_��д����) * .TextMatrix(intRow, mBillCol.C_�ɹ���), mFMT.FM_���)
                .TextMatrix(intRow, mconintcol�ۼ۽��) = Format(.TextMatrix(intRow, mBillCol.C_��д����) * .TextMatrix(intRow, mBillCol.C_�ۼ�), mFMT.FM_���)
                .TextMatrix(intRow, mconintcol���) = Format(.TextMatrix(intRow, mconintcol�ۼ۽��) - .TextMatrix(intRow, mBillCol.C_�ɹ����), mFMT.FM_���)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExpend_Click()
    Call AutoExpend
End Sub

'����
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
        
        cmdRequestTransfer.Left = txtCode.Left + txtCode.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    Else
        FindRownew mshBill, mBillCol.C_����, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
        
        cmdRequestTransfer.Left = cmdFind.Left + cmdFind.Width + (cmdFind.Left - cmdHelp.Left - cmdHelp.Width)
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdRequestTransfer_Click()
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim blnDo As Boolean
    Dim str���Ч�� As String
    Dim blnҩ�� As Boolean
    Dim dblPrice As Double
    Dim strЧ�� As String
    Dim dbl���� As Double
    Dim lng����ID As Long
    Dim bln���� As Boolean
    Dim dbl�깺����  As Double
    Dim dbl�ѵ����� As Double
    
    If cboEnterStock.ListCount = 0 Then  '������ⷿ
        MsgBox "����ⷿ����Ϊ�գ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)) = 0 Then
        MsgBox "����ⷿ����Ϊ�գ�", vbInformation, gstrSysName
        cboEnterStock.SetFocus
        Exit Sub
    End If
    
    mstrRequestNO = frmDrawCondition.ShowMe(Me, mintUnit, cboStock.Text, Val(cboStock.ItemData(cboStock.ListIndex)), cboEnterStock.Text, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)))
    If mstrRequestNO <> "" Then
        blnDo = False
        mstrRequestNO = Mid(mstrRequestNO, 1, LenB(StrConv(mstrRequestNO, vbFromUnicode)) - 1)

        blnҩ�� = True
        gstrSQL = "Select Distinct 0 " & _
                                    "From ��������˵�� " & _
                                    "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
        If rsTemp.RecordCount = 0 Then
            blnҩ�� = False
        End If

        gstrSQL = "Select a.Id as ����id, d.���� as �ƻ�����,a.����,a.���� ,a.���,c.�ּ� as �ۼ�,a.���㵥λ as ɢװ��λ,a.�Ƿ��� as ʱ��,b.��װ��λ,b.����ϵ��,b.ָ�������,b.���Ч��" & vbNewLine & _
                    ",e.�ϴβ��� as ����,e.�ϴ����� as ����,nvl(e.����,0) as ����,e.Ч��,e.���Ч��,e.��������,nvl(e.ʵ������,0) as ʵ������,e.ʵ�ʽ��,e.ʵ�ʲ��,e.���ۼ�,e.ƽ���ɱ���,e.��׼�ĺ�,b.�ⷿ����,b.���÷���, nvl(b.���ٲ���,0) as ���ٲ���" & vbNewLine & _
                    "From �շ���ĿĿ¼ A, �������� B, �շѼ�Ŀ C," & vbNewLine & _
                    "     (Select  b.����id, Sum(b.�ƻ�����) As ����" & vbNewLine & _
                    "       From ���ϲɹ��ƻ� A, ���ϼƻ����� B" & vbNewLine & _
                    "       Where a.Id = b.�ƻ�id and a.����=1 And a.No In (Select * From Table(Cast(f_Str2list([1]) As Zltools.t_Strlist)))" & vbNewLine & _
                    "       Group By b.����id) D,ҩƷ��� e" & vbNewLine & _
                    "Where a.Id = b.����id And b.����id = c.�շ�ϸĿid And a.Id = d.����id and b.����id=e.ҩƷid(+)  and e.�ⷿid=[2] and e.ʵ������>0 and e.����=1 And Sysdate Between c.ִ������ And c.��ֹ����"

        If gSystem_Para.P156_�����㷨 = 0 Then '���λ���Ч�������ȳ���
            gstrSQL = gstrSQL & " Order by a.id,Nvl(e.����, 0)"
        Else
            gstrSQL = gstrSQL & " Order by a.id,e.Ч��,Nvl(e.����, 0)"
        End If

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cmdRequestTransfer_Click", mstrRequestNO, cboStock.ItemData(cboStock.ListIndex))

        Do While Not rsTemp.EOF
            With mshBill
                For lngRow = 1 To .Rows - 1
                    If Val(.TextMatrix(lngRow, 0)) <> 0 Then
                        If Val(.TextMatrix(lngRow, 0)) = rsTemp!����ID And Val(.TextMatrix(lngRow, mBillCol.c_����)) = rsTemp!���� Then
                            blnDo = True
                            MsgBox "�ظ�����" & "[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "������ӣ�", vbInformation, gstrSysName
                            Exit For
                        End If
                    End If
                Next

                If Val(.TextMatrix(.Rows - 1, 0)) = 0 Then
                    lngRow = .Rows - 1
                Else
                    .Rows = .Rows + 1
                    lngRow = .Rows - 1
                End If

                str���Ч�� = IIf(IsNull(rsTemp!���Ч��), "", Format(rsTemp!���Ч��, "yyyy-MM-dd"))
                If Format(str���Ч��, "yyyy-mm-dd") < Format(zlDatabase.Currentdate, "yyyy-mm-dd") And Trim(str���Ч��) <> "" Then
                   If MsgBox("[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "�����Ѿ��������Ч��,�Ƿ�Ҫ���ã�", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbYes Then
                        blnDo = True
                   End If
                End If

'                strЧ�� = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd"))
'                If IsDate(strЧ��) Then
'                    If Format(strЧ��, "yyyy-MM-dd") < Format(zldatabase.Currentdate, "yyyy-MM-dd") Then
'                        MsgBox "[" & rsTemp!���� & "-" & rsTemp!���� & "]" & "���������Ѿ�ʧЧ�ˣ�", vbInformation, gstrSysName
'                    End If
'                End If

                'ȡ�ۼ�
                If rsTemp!ʱ�� = 1 Then
                    If rsTemp!���÷��� = 0 Then
                        If rsTemp!�ⷿ���� = 1 And blnҩ�� = False Then
                            bln���� = True
                        Else
                            bln���� = False
                        End If
                    Else
                        bln���� = True
                    End If

                    If bln���� = True Then
                        If IsNull(rsTemp!���ۼ�) Then
                            If rsTemp!ʵ������ = 0 Then
                                dblPrice = 0
                            Else
                                dblPrice = rsTemp!ʵ�ʽ�� / rsTemp!ʵ������
                            End If
                        Else
                            dblPrice = rsTemp!���ۼ�
                        End If
                    Else
                        If rsTemp!ʵ������ = 0 Then
                            dblPrice = 0
                        Else
                            dblPrice = rsTemp!ʵ�ʽ�� / rsTemp!ʵ������
                        End If
                    End If
                Else
                    dblPrice = IIf(IsNull(rsTemp!�ۼ�), 0, rsTemp!�ۼ�)
                End If

                If lng����ID = rsTemp!����ID Then
                    If rsTemp!�������� + dbl�ѵ����� > rsTemp!�ƻ����� Then
                        If rsTemp!�ƻ����� - dbl�ѵ����� <> 0 Then
                            dbl���� = rsTemp!�ƻ����� - dbl�ѵ�����
                            dbl�ѵ����� = dbl�ѵ����� + dbl����
                        Else
                            blnDo = True
                        End If
                    Else
                        dbl���� = rsTemp!��������
                        dbl�ѵ����� = dbl�ѵ����� + dbl����
                    End If
                Else
                    If rsTemp!�������� > rsTemp!�ƻ����� Then
                        dbl���� = rsTemp!�ƻ�����
                    Else
                        dbl���� = rsTemp!��������
                    End If
                    dbl�ѵ����� = dbl����
                End If
                lng����ID = rsTemp!����ID

                If dbl���� = 0 Then
                    blnDo = True
                End If

'                'ֻ�в��ظ��Ĳ���ӵ������ȥ
'                If blnDo = False Then
'                    SetRequestColValue lngRow, rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, _
'                                IIf(IsNull(rsTemp!���), "", rsTemp!���), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
'                                IIf(mintUnit = 0, rsTemp!ɢװ��λ, rsTemp!��װ��λ), _
'                                dblPrice, rsTemp!ƽ���ɱ���, IIf(IsNull(rsTemp!����), "", rsTemp!����), _
'                                IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd")), _
'                                IIf(IsNull(rsTemp!���Ч��), "", Format(rsTemp!���Ч��, "yyyy-MM-dd")), _
'                                rsTemp!�ƻ�����, _
'                                IIf(IsNull(rsTemp!��������), "0", rsTemp!��������), _
'                                dbl����, _
'                                IIf(IsNull(rsTemp!ָ�������), "0", rsTemp!ָ�������), _
'                                IIf(mintUnit = 0, 1, rsTemp!����ϵ��), IIf(IsNull(rsTemp!����), 0, rsTemp!����), rsTemp!ʱ��, rsTemp!���÷���, IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�), rsTemp!���ٲ���, rsTemp!�ⷿ����
'                End If

                'ֻ�в��ظ��Ĳ���ӵ������ȥ
                If blnDo = False Then
                    SetRequestColValue lngRow, rsTemp!����ID, "[" & rsTemp!���� & "]" & rsTemp!����, _
                    IIf(IsNull(rsTemp!���), "", rsTemp!���), IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                    IIf(mintUnit = 0, rsTemp!ɢװ��λ, rsTemp!��װ��λ), _
                    rsTemp!�ۼ�, IIf(IsNull(rsTemp!����), "", rsTemp!����), _
                    IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-MM-dd")), _
                    IIf(IsNull(rsTemp!���Ч��), "", Format(rsTemp!���Ч��, "yyyy-MM-dd")), _
                    IIf(IsNull(rsTemp!���Ч��), "0", rsTemp!���Ч��), _
                    rsTemp!�ⷿ����, _
                    IIf(IsNull(rsTemp!��������), "0", rsTemp!��������), _
                    IIf(IsNull(rsTemp!ʵ�ʽ��), "0", rsTemp!ʵ�ʽ��), _
                    IIf(IsNull(rsTemp!ʵ�ʲ��), "0", rsTemp!ʵ�ʲ��), _
                    IIf(IsNull(rsTemp!ָ�������), "0", rsTemp!ָ�������), _
                    IIf(mintUnit = 0, 1, rsTemp!����ϵ��), IIf(IsNull(rsTemp!����), 0, rsTemp!����), rsTemp!ʱ��, rsTemp!���÷���, IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
                    
                    With mshBill
                        .Row = lngRow
                        .TextMatrix(lngRow, mBillCol.C_�к�) = lngRow
                    
                        .TextMatrix(lngRow, mBillCol.C_��д����) = Format(dbl���� / IIf(mintUnit = 0, 1, rsTemp!����ϵ��), mFMT.FM_����)
                        
                        If .TextMatrix(lngRow, mBillCol.C_�ۼ�) <> "" Then
                            .TextMatrix(lngRow, mconintcol�ۼ۽��) = Format(.TextMatrix(lngRow, mBillCol.C_�ۼ�) * .TextMatrix(lngRow, mBillCol.C_��д����), mFMT.FM_���)
                        End If
                        
                        .TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(Get�ɱ���(Val(.TextMatrix(lngRow, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(lngRow, mBillCol.c_����))) * Val(.TextMatrix(lngRow, mBillCol.C_����ϵ��)), mFMT.FM_�ɱ���)
                        .TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(Val(.TextMatrix(lngRow, mBillCol.C_�ɹ���)) * .TextMatrix(lngRow, mBillCol.C_��д����), mFMT.FM_���)
                        .TextMatrix(lngRow, mconintcol���) = Format(Val(.TextMatrix(lngRow, mconintcol�ۼ۽��)) - Val(.TextMatrix(lngRow, mBillCol.C_�ɹ����)), mFMT.FM_���)
    
    
                        .TextMatrix(lngRow, mBillCol.C_ʵ������) = Format(dbl���� / IIf(mintUnit = 0, 1, rsTemp!����ϵ��), mFMT.FM_����)
                    End With
                    
                End If
                
                blnDo = False
                rsTemp.MoveNext
            End With
        Loop
    End If
End Sub





Private Sub CmdSave_Click()
    Dim blnSuccess As Boolean
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    Dim strReg As String
    
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
    
   
    If mint�༭״̬ = 10 Then        '����
        '����������ֽ⣬����������ˣ���˴˴�����飬ǿ���û��ֹ�����ֽ⹦��
        If Not ValidData Then Exit Sub
        If Not CheckStock Then Exit Sub
        
        '����Ƿ��ѱ�ҩ
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ����=19 And NO=[1] And ��ҩ�� Is Not NULL"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���", txtNO.Tag)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "�õ����ѱ���������Աȡ�����ϣ���ǰ������ֹ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����Ƿ��ѷ���
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ����=19 And NO=[1] And ��ҩ���� Is Not NULL"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ���", txtNO.Tag)
        If rsTemp.RecordCount <> 0 Then
            MsgBox "�õ����ѱ���������Ա���ͣ���ǰ������ֹ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        blnTrans = True
        gcnOracle.BeginTrans
        
        '��ɾ�����쵥�������ݵ�ǰ���ݲ����ƿⵥ
        If Not SaveCard(True) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
        
        '����
        gstrSQL = "zl_�����ƿ�_Prepare('" & txtNO.Tag & "','" & Txt�����.Caption & "')"
        zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
                        
        
        '���ͣ��³���ⷿ�Ĳ��Ͽ��ÿ�棩
        gstrSQL = "zl_�����ƿ�_Prepare('" & txtNO.Tag & "')"
        zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
        
        gcnOracle.CommitTrans
        blnTrans = True
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then        '���
        '�ж��Ƿ��Զ�ִ���ƿ����̣�����Ǿ��Զ���ɱ��ϡ����͡����չ���
        If mint�ƿ⴦������ = 0 Then
            blnSuccess = Auto�����ƿ�����
            Exit Sub
        End If
    
        If Not CheckSend Then Exit Sub
        If Not ���ϵ������(Txt������.Caption) Then Exit Sub
        
        If Not ��鵥��(19, txtNO.Tag, False) And Not mblnUpdate Then
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        gcnOracle.BeginTrans
        '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
            
            '����
            gstrSQL = "zl_�����ƿ�_Prepare('" & txtNO.Tag & "','" & UserInfo.�û��� & "')"
            zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
            '���ͣ��³���ⷿ�Ĳ��Ͽ��ÿ�棩
            gstrSQL = "zl_�����ƿ�_Prepare('" & txtNO.Tag & "')"
            zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
        End If

        If SaveCheck(True) = True Then
            strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
            If Val(strReg) = 1 Then
                '��ӡ
                If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                    printbill
                End If
            End If
            blnTrans = False
            gcnOracle.CommitTrans
            Unload Me
        Else
            gcnOracle.RollbackTrans: Exit Sub
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then '����
        If SaveStrike Then
            If mint������ʽ = 2 Then
                strReg = IIf(Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
                If Val(strReg) = 1 Then
                    '��ӡ
                    If InStr(mstrPrivs, "���ݴ�ӡ") <> 0 Then
                        printbill
                    End If
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If ValidData = False Then Exit Sub
    
    blnSuccess = SaveCard
        
    If blnSuccess = True Then
            
        strReg = IIf(Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModule, "0")) = 1, 1, 0)
        If Val(strReg) = 1 Then
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
    
    If mint�༭״̬ = 11 Then
        Unload Me
        Exit Sub
    End If
    
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
    
    txtժҪ.Text = ""
    If cboEnterStock.Enabled Then cboEnterStock.SetFocus
    mblnChange = False
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lng����ID As Long
    Dim dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsprice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    
    On Error GoTo ErrHandle
    
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))" & _
                   GetPriceClassString("") & ") B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 19 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(b.�ּ�, " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0" & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 19 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") <> Round(decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�), " & g_С��λ��.obj_ɢװС��.���ۼ�С�� & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, b.ƽ���ɱ��� As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B" & _
            " Where a.���� = 19 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ")<>round(b.ƽ���ɱ���," & g_С��λ��.obj_ɢװС��.�ɱ���С�� & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
            " Order By ����, ����id, ���"

    Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNO.Text))
    
    If rsprice.EOF Then Exit Sub
    
    lngRows = mshBill.Rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl���� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������))
        dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���))
        dbl���ۼ� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�))
        dbl�ɱ���� = dbl�ɱ��� * dbl����
        dbl���۽�� = dbl���ۼ� * dbl����
        dbl��� = dbl���۽�� - dbl�ɱ����
'
        If lng����ID <> 0 Then
            rsprice.Filter = "����='�ۼ�' And ����id=" & lng����ID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl���ۼ� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��)), mFMT.FM_���ۼ�))
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            rsprice.Filter = "����='�ɱ���' And ����id=" & lng����ID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
            If rsprice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(Format(dbl���ۼ� * dbl����, mFMT.FM_���))
                dbl�ɱ��� = Val(Format(rsprice!�ּ� * Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��)), mFMT.FM_���))
                dbl�ɱ���� = Val(Format(dbl�ɱ��� * dbl����, mFMT.FM_���))
                dbl��� = Val(Format(dbl���۽�� - dbl�ɱ����, mFMT.FM_���))
            End If

            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
                mshBill.TextMatrix(lngRow, mconintcol�ۼ۽��) = Format(dbl���۽��, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
                mshBill.TextMatrix(lngRow, mconintcol���) = Format(dbl���, mFMT.FM_���)
            End If
        End If
    Next
    rsprice.Filter = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
'    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 6 Then
                MsgBox "�õ�����û�п��Գ��������ģ����飡", vbOKOnly, gstrSysName
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
        stbThis.Panels("PY").Visible = True
        stbThis.Panels("WB").Visible = True
        gSystem_Para.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram stbThis, gSystem_Para.int���뷽ʽ
    Else
        stbThis.Panels("PY").Visible = False
        stbThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRownew mshBill, mBillCol.C_����, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If stbThis.Panels("PY").Bevel = sbrRaised Then
            Logogram stbThis, 0
        Else
            Logogram stbThis, 1
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strStock As String
    Dim rsEnterStock As New Recordset
    Dim rsPara As New ADODB.Recordset
    Dim strReg As String
    
    On Error GoTo ErrHandle
    mblnFirst = True
    mblnUpdate = False
    
    mbln����˲� = IIf((zlDatabase.GetPara("������Ҫ�˲������ƿ�", glngSys, 1722, "0")) = 0, False, True)
    strReg = Val(zlDatabase.GetPara("���ĵ�λ", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
     
    mintBatchNoLen = GetBatchNoLen()

    txtNO = mstr���ݺ�
    txtNO.Tag = txtNO.Text
    
    strStock = "And b.���� In('V','K','W','12') "
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "   Where c.�������� = b.���� And (a.վ��=[1] or a.վ�� is null) " & _
        "        " & strStock & _
        "       AND a.id = c.����id " & _
        "       AND a.����ʱ�� = to_date('3000-01-01','yyyy-MM-dd')"
    Set rsEnterStock = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, gstrNodeNo)
    
    With cboEnterStock
        .Clear
        Do While Not rsEnterStock.EOF
            .AddItem rsEnterStock.Fields(1)
            .ItemData(.NewIndex) = rsEnterStock.Fields(0)
            rsEnterStock.MoveNext
        Loop
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
        .Tag = 0
    End With
    
    'ȡϵͳ��������ȷ�����������Ρ�
    mbln��ȷ���� = IS��������
    
    'ȡϵͳ��������ȷ�ƿ��������Ρ�
    mbln�ƿ���ȷ���� = IS�����ƿ�
    
    mbln�����������Ų��ؿ��� = Val(zlDatabase.GetPara(305, glngSys, 0)) = 1
    
    '����ⷿȱʡΪ�����浱ǰѡ��Ŀⷿ������������Ч
    On Error Resume Next
    mlng����ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    
    Call initCard
    mstrTime_Start = GetBillInfo(19, mstr���ݺ�)
    '�ָ����Ի���������
    RestoreWinState Me, App.ProductName, mstrCaption
    '�ָ����Ի��������ú󣬻���Ҫ��Ȩ�޿��Ƶ��н�һ������
    With mshBill
        .ColWidth(mBillCol.C_�ɹ���) = IIf(mblnCostView = True, 800, 0)
        .ColWidth(mBillCol.C_�ɹ����) = IIf(mblnCostView = True, 900, 0)
        .ColWidth(mconintcol���) = IIf(mblnCostView = True, 800, 0)
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim str���� As String
    Dim strArray As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim lng����ⷿ As Long, lng���ⷿ As Long
    
    '�ⷿ
    On Error GoTo ErrHandle
    mbln���쵥 = False
    mstr�˲��� = ""
    mstr�˲����� = ""
    
    strOrder = zlDatabase.GetPara("��������", glngSys, mlngModule, "00")
    strOrder = IIf(strOrder = "", "00", strOrder)
    strCompare = Mid(strOrder, 1, 1)
    
    'ȡָ�����ݵĳ���ⷿ�����ⷿ
    gstrSQL = " Select �ⷿID,�Է�����ID From ҩƷ�շ���¼" & _
              " Where NO=[1] And ����=19 And ���ϵ��=-1 And Rownum<2"
    
    Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, "ȡָ�����ݵĳ���ⷿ�����ⷿ", mstr���ݺ�)
              
    If rsInitCard.RecordCount <> 0 Then
        lng����ⷿ = rsInitCard!�ⷿID
        lng���ⷿ = rsInitCard!�Է�����id
    End If
    If lng����ⷿ = 0 Then lng����ⷿ = mlng����ⷿ
    
    mint�ƿ⴦������ = IIf(Val(zlDatabase.GetPara("�ƿ�����", glngSys, mlngModule, "0", , , , lng����ⷿ)) = 1, 1, 0)
        
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                If .ItemData(i) = lng����ⷿ Then cboStock.ListIndex = cboStock.ListCount - 1
            Next
            mintcboIndex = cboStock.ListIndex
            '���û��ָ���Ĳ��ţ��������
            If mintcboIndex = -1 Then
                gstrSQL = "Select ID,���� From ���ű� Where ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���û��ָ���ĳ��ⲿ�ţ��������", lng����ⷿ)
                
                cboStock.AddItem rsTemp!����
                cboStock.ItemData(cboStock.NewIndex) = rsTemp!Id
                cboStock.ListIndex = cboStock.ListCount - 1
            End If
            mintcboIndex = cboStock.ListIndex
            cboStock.Enabled = .Enabled
        End With
        
    End If
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û���
            Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            
            If cboEnterStock.ListCount <> 0 Then
                If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                    If cboEnterStock.ListCount > 1 Then
                        cboEnterStock.ListIndex = cboEnterStock.ListIndex + 1
                    End If
                End If
            End If
        
        Case 2, 3, 4, 6, 10, 11
            initGrid
            '���õ����Ƿ������쵥��
            gstrSQL = "" & _
                "   Select Nvl(��ҩ��ʽ,0) ����,�˲���,�˲����� From ҩƷ�շ���¼ " & _
                "   Where ����=19 And NO=[1] And ���=1"
                
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
            
            If Not rsTemp.EOF Then
                mbln���쵥 = (rsTemp!���� = 1)
                mstr�˲��� = IIf(IsNull(rsTemp!�˲���), "", rsTemp!�˲���)
                mstr�˲����� = IIf(IsNull(rsTemp!�˲�����), "", rsTemp!�˲�����)
                If mbln���쵥 Then LblTitle.Caption = GetUnitName & "�������쵥"
            End If
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "" & _
                    "   Select distinct b.id,b.���� " & _
                    "   From ҩƷ�շ���¼ a,���ű� b " & _
                    "   Where a.�ⷿid=b.id and A.���� = 19 and a.no=[1] and a.���ϵ��=-1"
                
                Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, mstr���ݺ�)
                    
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsInitCard!����
                    .ItemData(.NewIndex) = rsInitCard!Id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
            
                Case 0
                    strUnitQuantity = "c.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                Case Else
                    strUnitQuantity = "B.��װ��λ AS ��λ,(A.��д���� / B.����ϵ��) AS ��д����,(A.ʵ������ / B.����ϵ��) AS ʵ������,a.�ɱ���*B.����ϵ�� as �ɱ���,a.���ۼ�*B.����ϵ�� as ���ۼ�,B.����ϵ�� as ����ϵ��,"
            End Select
            
            
            Select Case mint�༭״̬
            Case 6
                If mint������ʽ <> 2 Then
                    gstrSQL = "" & _
                        "   select w.*,z.��������/w.����ϵ�� as  ��������,z.ʵ�ʽ��,z.ʵ�ʲ�� " & _
                        "   From (  SELECT distinct a.����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ," & _
                        "                       zlSpellCode(c.����) ����,c.���,c.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,b.�ⷿ����," & _
                        "                       b.���Ч��,A.Ч��,A.�������,A.���Ч�� as ���ʧЧ��,B.һ���Բ���,b.���Ч��,A.��д���� ԭʼ����," & strUnitQuantity & _
                        "                       A.�ɱ����,0 ���۽��, 0 ���,a.ժҪ,a.�ⷿid,a.�Է�����id,c.�Ƿ���,b.���÷���  " & _
                        "           FROM (  Select min(id) as id, sum(ʵ������) as ��д����,0 ʵ������,sum(�ɱ����) as �ɱ����,ҩƷid ����ID,���,����,��׼�ĺ�, ����,Ч��,�������,���Ч�� ," & _
                        "                           Nvl(����,0) ����,����,�ɱ���,���ۼ�,ժҪ,�ⷿID,�Է�����ID,������ID" & _
                        "                   From ҩƷ�շ���¼ x " & _
                        "                   WHERE NO=[1] AND ����=19 And ���ϵ��=-1 " & _
                        "                   group by ҩƷID,���,����,��׼�ĺ�,����,Ч��,�������,���Ч��,Nvl(����,0),����,�ɱ���,���ۼ�,ժҪ,�ⷿID,�Է�����ID,������ID" & _
                        "                   having sum(ʵ������)<>0 ) A, �������� B,�շ���ĿĿ¼ C " & _
                        "           Where A.����id = B.����id  and A.����id=c.id " & _
                        "       ) w,(   Select  ҩƷid ����id,Nvl(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                        "               From ҩƷ��� " & _
                        "               where �ⷿid=[2]  and ����=1)  z " & _
                        "   Where w.����id=z.����id(+) and nvl(w.����,0)=nvl(z.����(+),0) " & _
                        "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                Else
                    '������˳���ʱ����ʾδ��˵������������
                    gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                        " FROM " & _
                        "     (SELECT DISTINCT A.ҩƷID as ����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ,zlSpellCode(c.����) as ����," & _
                        "     C.���,C.���� AS ԭ����,A.����,A.��׼�ĺ�, A.����,A.����,B.ָ�������,B.�ⷿ����," & _
                        "     B.���Ч��,A.Ч��,A.�������,A.���Ч�� as ���ʧЧ��,B.һ���Բ���,b.���Ч��,A.��д���� ԭʼ����," & strUnitQuantity & "A.�ɱ����,A.���۽��, A.���,A.��ҩ��, " & _
                        "     A.ժҪ,������,��������,�����,�������,A.�ⷿID,A.�Է�����ID,C.�Ƿ���,B.���÷���,NVL(A.��ҩ��λID,0) �ϴι�Ӧ��ID" & _
                        "     FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ C,�շ���Ŀ���� E " & _
                        "     WHERE A.ҩƷID = B.����ID AND B.����ID=C.ID AND B.����ID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                        "     AND A.��¼״̬ =[3] " & _
                        "     AND A.���� = 19 AND A.���ϵ��=-1 AND A.NO =[1] ) W," & _
                        "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                        "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                        " WHERE W.����id=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                        "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                End If
            Case 11
                Dim bln�߱����ٲ��� As Boolean
                bln�߱����ٲ��� = �ж�ֻ�߱����ϲ���(cboEnterStock.ItemData(cboEnterStock.ListIndex))

                gstrSQL = "" & _
                    "   Select w.����ID,w.���,w.������Ϣ,W.����,w.���,w.ԭ����,w.���� ,w.��׼�ĺ�,w.����,w.����,w.ָ�������, " & _
                    "           w.�ⷿ����,w.���Ч��,w.Ч��,w.�������,w.���ʧЧ��, " & _
                    "           w.һ���Բ���,w.���Ч��,w.��λ,w.ԭʼ���� ԭʼ����,w.��д����,w.ʵ������,w.���ۼ�,w.���۽��,w.����ϵ��, " & _
                    "           (w.���۽�� - Decode(Sign(nvl(z.ʵ�ʽ��,0)),1,w.���۽�� * (nvl(z.ʵ�ʲ��,0) / z.ʵ�ʽ��),w.���۽�� * w.ָ������� / 100)) / decode(w.ʵ������,0,1,w.ʵ������)  �ɱ���, " & _
                    "           (w.���۽�� - Decode(Sign(z.ʵ�ʽ��),1,w.���۽�� * (z.ʵ�ʲ�� / z.ʵ�ʽ��),w.���۽�� * w.ָ������� / 100)) �ɱ����, " & _
                    "           Decode(Sign(z.ʵ�ʽ��),1,w.���۽�� * (z.ʵ�ʲ�� / z.ʵ�ʽ��),w.���۽�� * w.ָ������� / 100) ���, " & _
                    "            w.ժҪ,w.������,w.��������,w.��ҩ��, w.�����, w.�������,w.�ⷿid,w.�Է�����id,w.�Ƿ���,w.���÷���,z.��������/w.����ϵ�� as  ��������,z.ʵ�ʽ��,z.ʵ�ʲ��   " & _
                    "    From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ,  " & _
                    "                    zlSpellCode(C.����) ����,c.���,C.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,b.�ⷿ����,  " & _
                    "                    b.���Ч��,A.Ч��,A.�������,A.���Ч�� as ���ʧЧ��,B.һ���Բ���,b.���Ч��,A.��д���� ԭʼ����, " & strUnitQuantity & _
                    "                    A.�ɱ����,A.���۽��, A.���,   " & _
                    "                    a.ժҪ,a.������,A.��������,A.��ҩ��, A.�����, A.�������,a.�ⷿid,a.�Է�����id,c.�Ƿ���,b.���÷���   " & _
                    "            FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ C  " & _
                    "            Where A.ҩƷid = B.����id and a.ҩƷid=c.id   " & IIf(bln�߱����ٲ���, " and nvl(B.��������,0)=1 ", "") & _
                    "                    AND A.��¼״̬ =[3]  " & _
                    "                    AND A.���� = 15 AND A.No = [1]  " & _
                    "           ) w, (  Select ҩƷid ����id,Nvl(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ��   " & _
                    "                    From ҩƷ��� where �ⷿid=[2]  and ����=1)  z, " & _
                    "  (Select Distinct �շ�ϸĿid From �շ�ִ�п��� f Where ִ�п���id = [4]) y " & _
                    "    Where w.����id=z.����id(+)  AND W.����id=Y.�շ�ϸĿid  and nvl(w.����,0)=nvl(z.����(+),0)   " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
                    
            Case Else
                gstrSQL = "" & _
                    "   Select w.*,z.��������/w.����ϵ�� as  ��������,z.ʵ�ʽ��,z.ʵ�ʲ�� " & _
                    "   From (  SELECT distinct a.ҩƷid ����id,A.���,('[' || c.���� || ']' || c.����) AS ������Ϣ," & _
                    "                   zlSpellCode(C.����) ����,c.���,C.���� as ԭ����,A.����,A.��׼�ĺ�, A.����,a.����,b.ָ�������,b.�ⷿ����," & _
                    "                   b.���Ч��,A.Ч��,A.�������,A.���Ч�� as ���ʧЧ��,B.һ���Բ���,b.���Ч��,A.��д���� ԭʼ����," & strUnitQuantity & _
                    "                   A.�ɱ����,A.���۽��, A.���, " & _
                    "                   a.ժҪ,������,��������,A.��ҩ��,�����,�������,a.�ⷿid,a.�Է�����id,c.�Ƿ���,b.���÷��� " & _
                    "           FROM ҩƷ�շ���¼ A, �������� B,�շ���ĿĿ¼ C " & _
                    "           Where A.ҩƷid = B.����id and a.ҩƷid=c.id " & _
                    "                   AND A.��¼״̬ =[3]" & _
                    "                   AND A.���� = 19 and a.���ϵ��=-1 AND A.No = [1]" & _
                    "          ) w, (  Select ҩƷid ����id,Nvl(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "                   From ҩƷ��� where �ⷿid=[2]  and ����=1)  z " & _
                    "   Where w.����id=z.����id(+) and nvl(w.����,0)=nvl(z.����(+),0) " & _
                    "   ORDER BY " & IIf(strCompare = "0", "���", IIf(strCompare = "1", "������Ϣ", "����")) & IIf(Right(strOrder, 1) = "0", " Asc", " Desc")
            
            End Select
            
            Set rsInitCard = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(mint�༭״̬ = 11, mstr��ⵥ��, mstr���ݺ�), cboStock.ItemData(cboStock.ListIndex), mint��¼״̬, cboEnterStock.ItemData(cboEnterStock.ListIndex))
            
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint�༭״̬
            Case 2, 6, 10, 11
                Txt������ = UserInfo.�û���
                Txt�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint�༭״̬ = 6 Or mint�༭״̬ = 10 Then
                    Txt����� = UserInfo.�û���
                    Txt������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
                If mint�༭״̬ = 10 Then
                    Txt����� = zlStr.NVL(rsInitCard!��ҩ��)
                    Txt������ = rsInitCard!������
                    Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                    Lbl�����.Caption = "������"
                    Lbl�������.Caption = "��������"
                End If
            Case Else
                Txt������ = rsInitCard!������
                Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            End Select
            
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            Dim intCount As Integer
            With cboEnterStock
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = rsInitCard!�Է�����id Then
                        .ListIndex = intCount
                        .Tag = intCount
                        Exit For
                    End If
                Next
            End With
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    .Rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    .TextMatrix(intRow, mBillCol.C_����) = rsInitCard!������Ϣ
                    .TextMatrix(intRow, mBillCol.c_���) = rsInitCard!���
                    .TextMatrix(intRow, mBillCol.c_���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mBillCol.C_����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mBillCol.c_��λ) = rsInitCard!��λ
                    .TextMatrix(intRow, mBillCol.c_����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mBillCol.C_Ч��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    
                    .TextMatrix(intRow, mBillCol.C_һ���Բ���) = zlStr.NVL(rsInitCard!һ���Բ���)
                    .TextMatrix(intRow, mBillCol.C_���Ч��) = zlStr.NVL(rsInitCard!���Ч��)
                    .TextMatrix(intRow, mBillCol.C_�������) = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd"))
                    .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = IIf(IsNull(rsInitCard!���ʧЧ��), "", Format(rsInitCard!���ʧЧ��, "yyyy-mm-dd"))
        
                    .TextMatrix(intRow, mBillCol.C_��д����) = Format(IIf(mint�༭״̬ = 6 And mint������ʽ = 2, -1, 1) * rsInitCard!��д����, mFMT.FM_����)
                    .TextMatrix(intRow, mBillCol.C_ʵ������) = Format(IIf(mint�༭״̬ = 6 And mint������ʽ = 2, -1, 1) * rsInitCard!ʵ������, mFMT.FM_����)
                    
                    
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 6 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Or mint�༭״̬ = 11 Then
                        .TextMatrix(intRow, mBillCol.c_ԭʼ����) = Format(IIf(mint�༭״̬ = 6 And mint������ʽ = 2, -1, 1) * rsInitCard!ԭʼ����, mFMT.FM_����)
                    End If

                    .TextMatrix(intRow, mBillCol.C_�ɹ���) = Format(rsInitCard!�ɱ���, mFMT.FM_�ɱ���)
                    .TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(IIf(mint�༭״̬ = 6 And mint������ʽ <> 2, 0, IIf(mint�༭״̬ = 6 And mint������ʽ = 2, -1, 1)) * rsInitCard!�ɱ����, mFMT.FM_���)
                    .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(rsInitCard!���ۼ�, mFMT.FM_���ۼ�)
                    .TextMatrix(intRow, mconintcol�ۼ۽��) = Format(IIf(mint�༭״̬ = 6 And mint������ʽ = 2, -1, 1) * rsInitCard!���۽��, mFMT.FM_���)
                    .TextMatrix(intRow, mconintcol���) = Format(IIf(mint�༭״̬ = 6 And mint������ʽ = 2, -1, 1) * rsInitCard!���, mFMT.FM_���)
                    .TextMatrix(intRow, mBillCol.C_���Ч��) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!���÷���
                    .TextMatrix(intRow, mBillCol.c_����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mBillCol.C_����ϵ��) = rsInitCard!����ϵ��
                    
                    .TextMatrix(intRow, mBillCol.C_ָ�������) = rsInitCard!ָ�������
                    .TextMatrix(intRow, mBillCol.C_�ⷿ����) = IIf(IsNull(rsInitCard!�ⷿ����), "0", rsInitCard!�ⷿ����)
                    .TextMatrix(intRow, mBillCol.C_��������) = Format(IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������), mFMT.FM_����)
                    .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = IIf(IsNull(rsInitCard!ʵ�ʲ��), "0", rsInitCard!ʵ�ʲ��)
                    .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = IIf(IsNull(rsInitCard!ʵ�ʽ��), "0", rsInitCard!ʵ�ʽ��)
                    
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!����ID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str���� = rsInitCard!����ID & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                        If mint�༭״̬ = 2 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!��д����), "0", rsInitCard!��д����)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!ʵ������), "0", rsInitCard!ʵ������)
                        End If
                        mcolUsedCount.Add Array(str����, strArray), str����
                    End If
                    rsInitCard.MoveNext
                Loop
                .Rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, 1)
    
    SetEdit         '���ñ༭����
    '���ġ��޸Ļ����ʱ�����ݿ��������������ʾ����
    If (mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 4 Or mint�༭״̬ = 10) Then
        If mbln���쵥 Then Call ShowColor
        Select Case mint�༭״̬
        Case 2, 10
            cmdExpend.Visible = True
        End Select
    End If
    If mint�ƿ⴦������ = 0 And mint�༭״̬ = 3 Then
        cmdExpend.Visible = True
    End If
    
    Call ��ʾ�ϼƽ��
    Exit Sub
ErrHandle:
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
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = False
            If mint�༭״̬ = 6 Then
                If mint������ʽ <> 2 Then
                    .ColData(mBillCol.C_ʵ������) = 4
                End If
            End If
        Else
            .ColData(0) = 5
            .ColData(mBillCol.C_����) = 1
            .ColData(mBillCol.c_���) = 5
            .ColData(mBillCol.c_���) = 5
            .ColData(mBillCol.C_����) = 5
            .ColData(mBillCol.c_��λ) = 5
            .ColData(mBillCol.c_����) = 5
            .ColData(mBillCol.C_Ч��) = 5
            .ColData(mBillCol.C_һ���Բ���) = 5
            .ColData(mBillCol.C_���Ч��) = 5
            .ColData(mBillCol.C_���ʧЧ��) = 5
            .ColData(mBillCol.C_�������) = 5
            .ColData(mBillCol.c_ԭʼ����) = 5
            If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                .ColData(mBillCol.C_��д����) = 4
                .ColData(mBillCol.C_�������) = 5
                .ColData(mBillCol.C_ʵ������) = 5
            ElseIf mint�༭״̬ = 3 Then
                .ColData(mBillCol.C_��д����) = 5
                .ColData(mBillCol.C_ʵ������) = 4
            ElseIf mint�༭״̬ = 11 Then
                .ColData(mBillCol.C_��д����) = 5
                .ColData(mBillCol.C_ʵ������) = 5
            End If
            
            .ColData(mBillCol.C_�ɹ���) = 5
            .ColData(mBillCol.C_�ɹ����) = 5
            .ColData(mBillCol.C_�ۼ�) = 5
            .ColData(mconintcol�ۼ۽��) = 5
            .ColData(mconintcol���) = 5
            
            .ColData(mBillCol.C_�ⷿ����) = 5
            .ColData(mBillCol.C_��������) = 5
            .ColData(mBillCol.C_���Ч��) = 5
            
            .ColData(mBillCol.C_ָ�������) = 5
            .ColData(mBillCol.C_ʵ�ʽ��) = 5
            .ColData(mBillCol.C_ʵ�ʲ��) = 5
            .ColData(mBillCol.C_����ϵ��) = 5
            .ColData(mBillCol.c_����) = 5
        
            .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
            .ColAlignment(mBillCol.c_���) = flexAlignLeftCenter
            .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
            .ColAlignment(mBillCol.c_��λ) = flexAlignCenterCenter
            .ColAlignment(mBillCol.c_����) = flexAlignLeftCenter
            .ColAlignment(mBillCol.C_Ч��) = flexAlignLeftCenter
            .ColAlignment(mBillCol.C_��д����) = flexAlignRightCenter
            .ColAlignment(mBillCol.C_ʵ������) = flexAlignRightCenter
            
            .ColAlignment(mBillCol.C_�ɹ���) = flexAlignRightCenter
            .ColAlignment(mBillCol.C_�ɹ����) = flexAlignRightCenter
            .ColAlignment(mBillCol.C_�ۼ�) = flexAlignRightCenter
            .ColAlignment(mconintcol�ۼ۽��) = flexAlignRightCenter
            .ColAlignment(mconintcol���) = flexAlignRightCenter
            
            If mint�༭״̬ = 11 Then
                '���ת��Ҳ���ܽ��б༭
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            If mint�༭״̬ = 11 Then
                cboEnterStock.Enabled = False
            Else
                cboEnterStock.Enabled = True
            End If
            txtժҪ.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    With mshBill
        .Active = (mint�༭״̬ <> 11)
        .Cols = mBillCols
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mBillCol.C_�к�) = ""
        .TextMatrix(0, mBillCol.C_����) = "���������"
        .TextMatrix(0, mBillCol.c_���) = "���"
        .TextMatrix(0, mBillCol.c_���) = "���"
        .TextMatrix(0, mBillCol.C_����) = "����"
        .TextMatrix(0, mBillCol.C_��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mBillCol.c_��λ) = "��λ"
        .TextMatrix(0, mBillCol.c_����) = "����"
        .TextMatrix(0, mBillCol.C_Ч��) = "ʧЧ��"
        
        .TextMatrix(0, mBillCol.C_һ���Բ���) = "һ���Բ���"
        .TextMatrix(0, mBillCol.C_���Ч��) = "���Ч��"
        .TextMatrix(0, mBillCol.C_���ʧЧ��) = "���ʧЧ��"
        .TextMatrix(0, mBillCol.C_�������) = "�������"
        
        .TextMatrix(0, mBillCol.C_��д����) = IIf(mint�༭״̬ = 6, "����", "��д����")
        .TextMatrix(0, mBillCol.C_ʵ������) = IIf(mint�༭״̬ = 6, "��������", "ʵ������")
        
        .TextMatrix(0, mBillCol.C_�ɹ���) = "�ɱ���"
        .TextMatrix(0, mBillCol.C_�ɹ����) = "�ɱ����"
        .TextMatrix(0, mBillCol.C_�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconintcol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintcol���) = "���"
        
        .TextMatrix(0, mBillCol.C_��������) = "��������"
        .TextMatrix(0, mBillCol.C_�ⷿ����) = "�ⷿ����"
        .TextMatrix(0, mBillCol.C_���Ч��) = "���Ч��"
        .TextMatrix(0, mBillCol.C_ʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mBillCol.C_ʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mBillCol.C_ָ�������) = "ָ�������"
        .TextMatrix(0, mBillCol.C_����ϵ��) = "����ϵ��"
        .TextMatrix(0, mBillCol.c_����) = "����"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mBillCol.C_�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mBillCol.C_�к�) = 300
        .ColWidth(mBillCol.C_����) = 2000
        .ColWidth(mBillCol.c_���) = 0
        .ColWidth(mBillCol.c_���) = 900
        .ColWidth(mBillCol.C_����) = 800
        .ColWidth(mBillCol.C_��׼�ĺ�) = 800
        .ColWidth(mBillCol.c_��λ) = 500
        .ColWidth(mBillCol.c_����) = 800
        .ColWidth(mBillCol.C_Ч��) = 1000
     
        .ColWidth(mBillCol.C_һ���Բ���) = 0
        .ColWidth(mBillCol.C_���Ч��) = 0
        .ColWidth(mBillCol.C_���ʧЧ��) = 1000
        .ColWidth(mBillCol.C_�������) = 0
          
        .ColWidth(mBillCol.C_��д����) = 800
        .ColWidth(mBillCol.C_ʵ������) = 800
        .ColWidth(mBillCol.C_�ɹ���) = IIf(mblnCostView = False, 0, 800)
        .ColWidth(mBillCol.C_�ɹ����) = IIf(mblnCostView = False, 0, 900)
        .ColWidth(mBillCol.C_�ۼ�) = 800
        .ColWidth(mconintcol�ۼ۽��) = 900
        .ColWidth(mconintcol���) = IIf(mblnCostView = False, 0, 800)
        
        .ColWidth(mBillCol.C_�ⷿ����) = 0
        .ColWidth(mBillCol.C_��������) = 0
        .ColWidth(mBillCol.C_���Ч��) = 0
        .ColWidth(mBillCol.C_ʵ�ʲ��) = 0
        .ColWidth(mBillCol.C_ʵ�ʽ��) = 0
        .ColWidth(mBillCol.C_ָ�������) = 0
        .ColWidth(mBillCol.C_����ϵ��) = 0
        .ColWidth(mBillCol.c_����) = 0
        .ColWidth(mBillCol.c_ԭʼ����) = 0
        
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        
        .ColData(mBillCol.C_�к�) = 5
        .ColData(mBillCol.c_���) = 5
        .ColData(mBillCol.c_���) = 5
        .ColData(mBillCol.C_����) = 5
        .ColData(mBillCol.C_��׼�ĺ�) = 5
        .ColData(mBillCol.c_��λ) = 5
        .ColData(mBillCol.c_����) = 5
        .ColData(mBillCol.C_Ч��) = 5
   
        .ColData(mBillCol.C_һ���Բ���) = 5
        .ColData(mBillCol.C_���Ч��) = 5
        .ColData(mBillCol.C_���ʧЧ��) = 5
        .ColData(mBillCol.C_�������) = 5
        .ColData(mBillCol.c_ԭʼ����) = 5
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            cboEnterStock.Enabled = True
            txtժҪ.Enabled = True
            
            cboStock.Enabled = True
   
            .ColData(mBillCol.C_����) = 1
            .ColData(mBillCol.C_��д����) = 4
            .ColData(mBillCol.C_ʵ������) = 5
            .ColData(mBillCol.C_�������) = 5
            
        ElseIf mint�༭״̬ = 3 Or mint�༭״̬ = 6 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mBillCol.C_����) = 5
            .ColData(mBillCol.C_��д����) = 5
            .ColData(mBillCol.C_ʵ������) = 4

        ElseIf mint�༭״̬ = 4 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mBillCol.C_��д����) = 5
            .ColData(mBillCol.C_ʵ������) = 5
            .ColData(mBillCol.C_����) = 5
        End If
        
        .ColData(mBillCol.C_�ɹ���) = 5
        .ColData(mBillCol.C_�ɹ����) = 5
        .ColData(mBillCol.C_�ۼ�) = 5
        .ColData(mconintcol�ۼ۽��) = 5
        .ColData(mconintcol���) = 5
        
        .ColData(mBillCol.C_�ⷿ����) = 5
        .ColData(mBillCol.C_��������) = 5
        .ColData(mBillCol.C_���Ч��) = 5
        .ColData(mBillCol.C_ʵ�ʲ��) = 5
        .ColData(mBillCol.C_ʵ�ʽ��) = 5
        .ColData(mBillCol.C_ָ�������) = 5
        .ColData(mBillCol.C_����ϵ��) = 5
        .ColData(mBillCol.c_����) = 5
        
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_���) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mBillCol.c_��λ) = flexAlignCenterCenter
        .ColAlignment(mBillCol.c_����) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_Ч��) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_��д����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_ʵ������) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_�ɹ���) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ɹ����) = flexAlignRightCenter
        .ColAlignment(mBillCol.C_�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconintcol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintcol���) = flexAlignRightCenter
        
        .ColAlignment(mBillCol.C_һ���Բ���) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_���Ч��) = flexAlignLeftCenter
        .ColAlignment(mBillCol.C_���ʧЧ��) = flexAlignCenterCenter
        .ColAlignment(mBillCol.C_�������) = flexAlignCenterCenter
        
        .PrimaryCol = mBillCol.C_����
        .LocateCol = mBillCol.C_����
        If InStr(1, "346", mint�༭״̬) <> 0 Then .ColData(mBillCol.C_����) = 0
    End With
    txtժҪ.MaxLength = sys.FieldsLength("ҩƷ�շ���¼", "ժҪ")
    chkIn.Visible = (mint�༭״̬ = 1)
    txtIn.Visible = (mint�༭״̬ = 1)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
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
    With txtNO
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cboEnterStock.Left = mshBill.Left + mshBill.Width - cboEnterStock.Width
    
    LblEnterStock.Left = cboEnterStock.Left - LblEnterStock.Width - 100
    
    
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
        
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    If mblnCostView = False Then
        lblPurchasePrice.Visible = False
    End If
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnCostView = False Then
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
    
    With cmdRequestTransfer
        .Top = cmdFind.Top
        
        .Visible = (mint�༭״̬ = 1 Or mint�༭״̬ = 2) '�������޸Ĳſɼ�
        
    End With
    
    With cmdExpend
        .Top = CmdSave.Top
        .Left = CmdSave.Left - 150 - .Width
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If msh����.Visible = True Then '���ص��б�����ȹرղ��ص��б�
        msh����.Visible = False
        mshBill.SetFocus
        mshBill.Col = mBillCol.C_����
        Cancel = True
        Exit Sub
    End If
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, mstrCaption
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, mstrCaption
    End If
    
End Sub

Private Function SaveCheck(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim rs��� As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim lng�Է�����id As Long
    Dim str����� As String
    
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng������ As Long
    Dim dbl��д���� As Double
    Dim dblʵ������ As Double
    Dim dbl�ɱ��� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl�ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim lng�����id As Long
    Dim lng�����id As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim str������� As String
    Dim str������� As String
    Dim int���к� As Integer
    Dim n As Long
    
    Dim arrSQL As Variant
    
    On Error GoTo ErrHandle
    arrSQL = Array()
    mblnSave = False
    SaveCheck = False
    
    '���õ����Ƿ��ڽ���༭����󣬱���������Ա�޸�
    mstrTime_End = GetBillInfo(19, mstr���ݺ�)
    If mstrTime_End = "" Then
        MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    If Not blnǿ�Ʊ��� And mint�ƿ⴦������ <> 0 Then
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng�Է�����id = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    str����� = UserInfo.�û���
    strNo = txtNO.Tag
    
    gstrSQL = "" & _
        "   SELECT b.ϵ��,b.id AS ���id " & _
        "   FROM ҩƷ�������� a, ҩƷ������ b " & _
        "   Where a.���id = b.ID " & _
        "           AND a.���� = 34 "
    
    zlDatabase.OpenRecordset rs���, gstrSQL, mstrCaption
    
    If rs���.EOF Then
        MsgBox "����������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rs���.RecordCount < 2 Then
        MsgBox "����������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rs���.MoveFirst
    Do While Not rs���.EOF
        If rs���!ϵ�� = 1 Then
            lng�����id = rs���!���ID
        Else
            lng�����id = rs���!���ID
        End If
        rs���.MoveNext
    Loop
    rs���.Close
    
    str������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mBillCol.C_����)
                lng������ = .TextMatrix(intRow, mBillCol.c_����)
                dbl��д���� = Round(Val(.TextMatrix(intRow, mBillCol.C_��д����)) * .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_���С��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_���С��.����С��)
                If Val(Format(Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����)) / Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), mFMT.FM_����)) = Val(.TextMatrix(intRow, mBillCol.C_��д����)) Then
                    If dbl��д���� = dblʵ������ Then
                        dblʵ������ = Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����))
                        dbl��д���� = dblʵ������
                    End If
                End If
                
                dbl�ɱ��� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ���)) / .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ����)), g_С��λ��.obj_���С��.���С��)
                dbl�ۼ� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ�)) / .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mconintcol�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mconintcol���)), g_С��λ��.obj_���С��.���С��)
                str���� = .TextMatrix(intRow, mBillCol.c_����)
                strЧ�� = IIf(.TextMatrix(intRow, mBillCol.C_Ч��) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_Ч��) & "','yyyy-mm-dd')")
                str������� = IIf(.TextMatrix(intRow, mBillCol.C_���ʧЧ��) = "", "Null", "to_date('" & .TextMatrix(intRow, mBillCol.C_���ʧЧ��) & "','yyyy-mm-dd')")
                int���к� = Val(.TextMatrix(intRow, mBillCol.c_���))
                'zl_�����ƿ�_VERIFY( /*�ⷿID_IN*/, /*�Է�����ID_IN*/, /*����ID_IN*/,
                    '����_IN*/, /*������_IN*/, /*��д����_IN*/, /*ʵ������_IN*/, /*�ɱ���_IN*/,
                    '/*�ɱ����_IN*/, /*���۽��_IN*/, /*���_IN*/, /*�����ID_IN*/, /*�����ID_IN*/,
                    '/*NO_IN*/, /*�����_IN*/, /*����_IN*/, /*Ч��_IN*/���ʧЧ��/������� ,�ƿⵥ��־);
                        
                gstrSQL = "" & _
                    "zl_�����ƿ�_Verify(" & int���к� & "," & lng�ⷿID & "," & lng�Է�����id & "," & _
                     lng����ID & ",'" & str���� & "'," & lng������ & "," & dbl��д���� & "," & _
                     dblʵ������ & "," & dbl�ɱ��� & "," & dbl�ɱ���� & "," & dbl���۽�� & "," & _
                     dbl��� & "," & lng�����id & "," & lng�����id & ",'" & _
                     strNo & "','" & str����� & "','" & str���� & "'," & strЧ�� & "," & str������� & ",to_date('" & str������� & "','yyyy-mm-dd HH24:MI:SS')," & IIf(mbln���쵥 = True, 0, 1) & "," & dbl�ۼ� & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng����ID) & ";" & vbCrLf & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSQL, mstrCaption, False, Not blnǿ�Ʊ���) Then Exit Function
'    If Not ��鵥��(19, txtNO.Tag) Then
'
'        If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
'        Exit Function
'    End If
    If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mBillCol.C_�к�, mshBill.Row)
    If mbln���쵥 Then Call ShowColor
End Sub



Private Sub mshBill_BeforeAddRow(Row As Long)
    If mshBill.ColData(mBillCol.C_����) = 0 Then
        Exit Sub
    End If
        
        
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If mint�༭״̬ = 10 Then
        Cancel = True
        Exit Sub
    End If
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        If mint�༭״̬ = 3 And mbln���쵥 Then Exit Sub
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ���������ģ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim int����� As Integer
    Dim rsTemp As Recordset
    
    On Error GoTo ErrHandle
    
    int����� = mshBill.Row
    
    If cboEnterStock.ListCount = 0 Then Exit Sub
    
    If mshBill.Col = mBillCol.C_���� Then
        mbln����ʾ�п������ = gSystem_Para.para_������¿��ÿ�� And mint����� = 2
        If Not mbln���쵥 Then
            Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                mbln�ƿ���ȷ����, True, False, False, True, , , , , mbln����ʾ�п������, , , mstrPrivs, mbln�ƿ���ȷ����, False)
        Else
            Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                mbln��ȷ����, mbln��ȷ����, False, False, True, , , , , mbln����ʾ�п������, , , mstrPrivs, IIf(mbln���쵥 = True, mbln��ȷ����, True), False)
        End If
        If RecReturn.RecordCount > 0 Then
            mblnChange = True
            With mshBill
                Dim intUnit As Integer
                
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                        IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                        IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                        RecReturn!�ۼ�, IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                        IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                        IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
                        IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                        RecReturn!�ⷿ����, _
                        IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                        IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                        IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                        IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                        IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) Then
                        
                        If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                        .Row = .Row + 1
                    End If
                    
                    .Col = mBillCol.C_��д����
                    RecReturn.MoveNext
                Next
                
                mshBill.Row = int�����
                
                If mstr�ظ����� <> "" Then
                    MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                    mstr�ظ����� = ""
                End If
                
    '            If RecReturn.RecordCount = 1 Then
    '                SetColValue .Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
    '                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
    '                    IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
    '                    RecReturn!�ۼ�, IIf(IsNull(RecReturn!����), "", RecReturn!����), _
    '                    IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
    '                    IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
    '                    IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
    '                    RecReturn!�ⷿ����, _
    '                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
    '                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
    '                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
    '                    IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
    '                    IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)
    '
    '                .Col = mBillCol.C_��д����
    '            End If
            End With
            RecReturn.Close
        End If
    Else
        gstrSQL = "Select rownum as id,null as �ϼ�id,����,����,����,1 as ĩ�� From ���������� "
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 1, "����������ѡ��", True, , "ѡ���������������̻���")
        
        '     frmParent=��ʾ�ĸ�����
        '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
        '     bytStyle=ѡ�������
        '       Ϊ0ʱ:�б���:ID,��
        '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
        '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
        '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
        '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
        '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
        '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
        '             bytStyle=1ʱ,�����Ǳ��������
        '     strNote=ѡ������˵������
        '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
        '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
        '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
        '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
        '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
        '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
        If rsTemp Is Nothing Then Exit Sub
        If rsTemp.State <> 1 Then Exit Sub
        
        With rsTemp
            If CheckQualifications(mlngModule, 1, CStr(NVL(!����))) = False Then Exit Sub
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_����) = NVL(!����)
        End With
        
        gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", mshBill.TextMatrix(mshBill.Row, mBillCol.C_����), mshBill.TextMatrix(mshBill.Row, 0))
        If rsTemp.RecordCount > 0 Then
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
        Else
            mshBill.TextMatrix(mshBill.Row, mBillCol.C_��׼�ĺ�) = ""
        End If

    End If
    
    Exit Sub
ErrHandle:
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
        If .Col = mBillCol.C_��д���� Or mBillCol.C_ʵ������ Then
            strKey = .Text
            If strKey = "" Then
                strKey = .TextMatrix(.Row, .Col)
            End If
            Select Case .Col
                Case mBillCol.C_��д����, mBillCol.C_ʵ������
                    intDigit = IIf(mintUnit = 1, g_С��λ��.obj_��װС��.����С��, g_С��λ��.obj_ɢװС��.����С��)
            End Select
            
            If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                KeyAscii = 0
                Exit Sub
            End If
            
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
        If .Row <> .LastRow Or .LastRow = 1 Then 'Or .LastRow = 1���������Ϊ��һ�ν���.Row �� .LastRow �� = 1
            SetInputFormat .Row
        End If
        
        Select Case .Col
            Case mBillCol.C_����
                .TxtCheck = False
                .MaxLength = 80
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                
            Case mBillCol.c_����
                .TxtCheck = False
                .MaxLength = mintBatchNoLen
            
            Case mBillCol.C_Ч��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mBillCol.c_����) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mBillCol.c_����)) And .TextMatrix(.Row, mBillCol.C_���Ч��) <> "" Then
                        If Split(.TextMatrix(.Row, mBillCol.C_���Ч��), "||")(0) <> 0 Then
                            strxq = UCase(.TextMatrix(.Row, mBillCol.c_����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mBillCol.C_Ч��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mBillCol.C_���Ч��), "||")(0), strxq), "yyyy-mm-dd")
                            End If
                        End If
                    End If
                End If
            Case mBillCol.C_��д����, mBillCol.C_ʵ������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
            Case mBillCol.C_����
                ImeLanguage True
                .TxtCheck = False
                .MaxLength = 30
                .TxtSetFocus
        End Select
        
    End With
End Sub

Private Sub mshBill_GotFocus()
    If mintParallelRecord <> 1 Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then Exit Sub
    If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
        MsgBox "����ⷿ���Ƴ��ⷿ��ͬ�ˣ����������ѡ��", vbOKOnly + vbExclamation, gstrSysName
       If cboEnterStock.Enabled Then cboEnterStock.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim rsDrug As New Recordset
    Dim strUnit As String
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim int����� As Integer
    Dim rsTemp As New Recordset
    
    int����� = mshBill.Row
    
    On Error GoTo ErrHandle
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    With mshBill
'        .Text = UCase(Trim(.Text))
        strKey = Trim(.Text)
        
        If Mid(strKey, 1, 1) = "[" Then
            If InStr(2, strKey, "]") <> 0 Then
                strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
            Else
                strKey = Mid(strKey, 2)
            End If
        End If
        Select Case .Col
            
            Case mBillCol.C_����
                If strKey <> "" Then
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    mbln����ʾ�п������ = gSystem_Para.para_������¿��ÿ�� And mint����� = 2

                    If Not mbln���쵥 Then
                        Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                            strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, mbln�ƿ���ȷ����, True, False, False, True, , , , mbln����ʾ�п������, , , mstrPrivs, mbln�ƿ���ȷ����, False)
                    Else
                        Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), _
                            strKey, sngLeft, sngTop, mshBill.MsfObj.CellWidth, mshBill.MsfObj.CellHeight, mbln��ȷ����, mbln��ȷ����, False, False, True, , , , mbln����ʾ�п������, , , mstrPrivs, IIf(mbln���쵥 = True, mbln��ȷ����, True), False)
                    End If
                    
                    If RecReturn.RecordCount <= 0 Then
                        Cancel = True
                        Exit Sub
                    End If
                    
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
                                IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
                                IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                                RecReturn!�ⷿ����, _
                                IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                                IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) Then
                            
                            If .Row = .Rows - 1 Then .Rows = .Rows + 1 'ֻ�е�ǰ�������һ��ʱ��������
                            .Row = .Row + 1
                            
                            .Text = .TextMatrix(.Row, .Col)
                        Else
                            Cancel = True
                        End If
                        
                        RecReturn.MoveNext
                    Next
                    
                    mshBill.Row = int�����
                    
                    If mstr�ظ����� <> "" Then
                        MsgBox mstr�ظ����� & "�б����Ѿ������ˣ�" & vbCrLf & "�������Ĳ�����ӣ�", vbInformation + vbOKOnly, gstrSysName
                        mstr�ظ����� = ""
                    End If
                    
                    If mbln�ƿ���ȷ���� = False Then
                        .Col = mBillCol.C_��д����
                    End If
'                    If RecReturn.RecordCount = 1 Then
'                        If SetColValue(.Row, RecReturn!����ID, "[" & RecReturn!���� & "]" & RecReturn!����, _
'                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(mintUnit = 0, RecReturn!ɢװ��λ, RecReturn!��װ��λ), _
'                                IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
'                                IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!���ʧЧ��), "", Format(RecReturn!���ʧЧ��, "yyyy-MM-dd")), _
'                                IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
'                                RecReturn!�ⷿ����, _
'                                IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
'                                IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
'                                IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
'                                IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
'                                IIf(mintUnit = 0, 1, RecReturn!����ϵ��), IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!���÷���, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) = False Then
'                            Cancel = True
'                            Exit Sub
'                        End If
'                        .Text = .TextMatrix(.Row, .Col)
'                    Else
'                        Cancel = True
'                    End If
                    Call ��ʾ�����
                End If
            Case mBillCol.c_����
                '�޴���
                If strKey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mBillCol.c_����) = ""
                    End If
                    If .ColData(mBillCol.C_Ч��) = 2 Then
                        .Col = mBillCol.C_Ч��
                    Else
                        .Col = mBillCol.C_��д����
                    End If
                    
                    
                    Cancel = True
                    Exit Sub
                End If
                
            Case mBillCol.C_Ч��
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mBillCol.C_Ч��) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            Case mBillCol.C_�������
                '�д���
                If strKey <> "" Then
                    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
                        strKey = TranNumToDate(strKey)
                        If strKey = "" Then
                            MsgBox "������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strKey
                        'Exit Sub
                    End If
                    If Not IsDate(strKey) Then
                        MsgBox "������ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    If Format(sys.Currentdate, "yyyy-mm-dd") >= Format(DateAdd("m", Val(.TextMatrix(.Row, mBillCol.C_���Ч��)), CDate(strKey)), "yyyy-mm-dd") Then
                        If MsgBox("�������Ѿ��������ʧЧ��(" & Format(DateAdd("m", Val(.TextMatrix(.Row, mBillCol.C_���Ч��)), CDate(strKey)), "yyyy-mm-dd") & "),�Ƿ�Ҫ�������!", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                    
                    .Text = strKey
                    '����ʧЧ��
                    .TextMatrix(.Row, mBillCol.C_���ʧЧ��) = Format(DateAdd("m", Val(.TextMatrix(.Row, mBillCol.C_���Ч��)), CDate(strKey)), "yyyy-mm-dd")
                ElseIf strKey = "" And strKey <> .TextMatrix(.Row, mBillCol.C_�������) Then
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    Exit Sub
                End If
            Case mBillCol.C_��д����, mBillCol.C_ʵ������
                If .TextMatrix(.Row, 0) = "" Then .Text = "": .TextMatrix(.Row, mBillCol.C_��д����) = "": Exit Sub
                If .TextMatrix(.Row, .Col) = "" And strKey = "" Then
                    MsgBox "�����������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                If Not IsNumeric(strKey) And strKey <> "" Then
                    MsgBox "��������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strKey <> "" Then
                    If Val(strKey) = 0 And mint�༭״̬ <> 3 Then
                        MsgBox "�������������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Not CompareUsableQuantity(.Row, strKey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '�ɱ��۵Ĺ�ʽ��     ������=����*�ۼ�
                    '                  ������=������*��ʵ�ʲ��/ʵ�ʽ�
                    '                  if ʵ�ʽ��<=0 then  ������=������*ָ�������
                    '                  ���ۣ��ɱ��ۣ�=��������-�����ۣ�/����
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                    strKey = Format(strKey, mFMT.FM_����)
                    .Text = strKey
                    .TextMatrix(.Row, .Col) = .Text
                    
                    If .TextMatrix(.Row, mBillCol.C_�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconintcol�ۼ۽��) = Format(.TextMatrix(.Row, mBillCol.C_�ۼ�) * strKey, mFMT.FM_���)
                    End If
                    
                    If mint�༭״̬ <> 6 Then
                        Dim dbl��� As Double, dbl���� As Double, dbl�ɱ���� As Double
                        'cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����
'                        Call ��֤�����ۼ���(cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_����)), Val(.TextMatrix(.Row, mBillCol.C_����ϵ��)), Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʲ��)), Val(.TextMatrix(.Row, mBillCol.C_ʵ�ʽ��)), Val(Split(.TextMatrix(.Row, mBillCol.C_ָ�������), "||")(0)) / 100, Val(strKey), Val(.TextMatrix(.Row, mBillCol.C_�ۼ۽��)), dbl���, dbl����, dbl�ɱ����)
'                        .TextMatrix(.Row, mBillCol.C_���) = Format(dbl���, mFMT.FM_���)
                        .TextMatrix(.Row, mBillCol.C_�ɹ���) = Format(Get�ɱ���(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mBillCol.c_����))) * Val(.TextMatrix(.Row, mBillCol.C_����ϵ��)), mFMT.FM_�ɱ���)
'                        .TextMatrix(.Row, mBillCol.C_�ɹ����) = Format(dbl�ɱ����, mFMT.FM_���)
'                    Else
'                        .TextMatrix(.Row, mBillCol.C_�ɹ����) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ɹ���)) * strKey, mFMT.FM_���)
'                        .TextMatrix(.Row, mBillCol.C_���) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ۼ۽��)) - Val(.TextMatrix(.Row, mBillCol.C_�ɹ����)), mFMT.FM_���)
                    End If
                    .TextMatrix(.Row, mBillCol.C_�ɹ����) = Format(Val(.TextMatrix(.Row, mBillCol.C_�ɹ���)) * strKey, mFMT.FM_���)
                    .TextMatrix(.Row, mconintcol���) = Format(Val(.TextMatrix(.Row, mconintcol�ۼ۽��)) - Val(.TextMatrix(.Row, mBillCol.C_�ɹ����)), mFMT.FM_���)
                 
                    If .Col = mBillCol.C_��д���� Then
                        .TextMatrix(.Row, mBillCol.C_ʵ������) = strKey
                    End If
                End If
                ��ʾ�ϼƽ��
                If mbln���쵥 Then Call ShowColor(.Row)
            Case mBillCol.C_����
                '�޴���
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mBillCol.C_����) = ""
                    End If
                    .Col = mBillCol.c_����
                    Cancel = True
                    Exit Sub
                Else
                    Dim rs���� As New Recordset
                    
                    gstrSQL = "" & _
                        "   Select ����,����,���� From ���������� " & _
                        "   Where upper(����) like [1] or Upper(����) like [1] or Upper(����) like [1]"
                    
                    Set rs���� = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, IIf(gstrMatchMethod = "0", "%", "") & UCase(Trim(strKey)) & "%")
                    
                    
                    If rs����.EOF Then
                        If MsgBox("��������������û���ҵ�������Ĳ��أ���Ҫ��������������������������", vbYesNo + vbQuestion, mstrCaption) = vbNo Then
                            Cancel = True
                            Exit Sub
                        Else
                            Dim rsMax As New Recordset
                            Dim int���� As Integer, strCode As String, strSpecify As String
                            
                            If rsMax.State = 1 Then rsMax.Close
                            gstrSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ����������"
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            int���� = rsMax!Length
                            
                            gstrSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM ����������"
                            rsMax.Close
                            zlDatabase.OpenRecordset rsMax, gstrSQL, mstrCaption
                            strCode = rsMax!Code
                            
                            int���� = Len(strCode)
                            strCode = strCode + 1
                            
                            If int���� >= Len(strCode) Then
                                strCode = String(int���� - Len(strCode), "0") & strCode
                            End If
                            strSpecify = zlCommFun.SpellCode(strKey)
                            
                            
                            gstrSQL = "ZL_����������_INSERT('" & strCode & "','" & strKey & "','" & strSpecify & "')"
                            Call zlDatabase.ExecuteProcedure(gstrSQL, mstrCaption)
                        End If
                    Else
                        If rs����.RecordCount = 1 Then
                            If CheckQualifications(mlngModule, 1, rs����.Fields("����")) = False Then
                                Exit Sub
                            End If
                            
                            .TextMatrix(.Row, mBillCol.C_����) = rs����.Fields("����")
                            .Text = rs����.Fields("����")
                            
                            gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "mshbill_CommandClick", .TextMatrix(.Row, mBillCol.C_����), Val(.TextMatrix(.Row, 0)))
                            If rsTemp.RecordCount > 0 Then
                                .TextMatrix(.Row, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
                            Else
                                .TextMatrix(.Row, mBillCol.C_��׼�ĺ�) = ""
                            End If
                        Else
                            Set msh����.Recordset = rs����
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
                zlCommFun.OpenIme False
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'������Ŀ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
    ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal str���ʧЧ�� As String, ByVal int���Ч�� As Integer, ByVal int�ⷿ���� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal numָ������� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal str��׼�ĺ� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbltotal As Double
    Dim dblPrice As Double
    Dim intLop As Integer
    Dim rsprice As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsЧ�� As ADODB.Recordset
    Dim bln���� As Boolean
    
    On Error GoTo ErrHandle
    If str���ʧЧ�� <> "" Then
        If Format(str���ʧЧ��, "yyyy-mm-dd") <= Format(sys.Currentdate, "yyyy-mm-dd") Then
            If MsgBox("���ġ�" & str���� & "(" & lng���� & ")�������Ч���Ѿ�����,�Ƿ�Ҫ�����ƿ�?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    gstrSQL = "Select һ���Բ���,���Ч�� from �������� where ����id=[1]"
    Set rsЧ�� = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
    
    SetColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID And Val(.TextMatrix(lngRow, mBillCol.c_����)) = lng���� Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & str���� & "��"  '����¼�����ظ�������
                    'Call MsgBox("�������ϡ�" & str���� & "(" & lng���� & ")���Ѿ����ڣ���ϲ��������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If lng���� > 0 Then   '���Ƴ��ⷿ�ǿⷿ�������ǿⷿ���������ĵ��ж�
            If mint�༭״̬ = 1 Then
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And lng���� = .TextMatrix(intLop, mBillCol.c_����) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mBillCol.C_��д����)
                        End If
                    End If
                Next
                
                If dbltotal >= num�������� And dbltotal <> 0 Then
                    MsgBox "�����ĵĿ��ÿ��������û���ˣ������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            End If
        End If
        
        If int�Ƿ��� = 1 Then
            If int���÷��� = 0 Then
                If int�ⷿ���� = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From ��������˵�� " & _
                            "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln���� = True
                    End If
                End If
            Else
                bln���� = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(���ۼ�,0)*" & num����ϵ�� & " as  �����ۼ�,ʵ�ʽ��/ʵ������* " & num����ϵ�� & " as ƽ�����ۼ�" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "       and ҩƷid=[2]" & _
                "       and ����=1 and ʵ������>0 and " & _
                "       nvl(����,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
            If rsprice.EOF Then
                If (mbln��ȷ���� = True And mbln���쵥 = True) Or (mbln�ƿ���ȷ���� = True And mbln���쵥 = False) Then 'Or (mbln���쵥 = False And bln���� = False)
                    MsgBox "ʱ������û�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
                    Exit Function
                ElseIf mbln��ȷ���� = False And mbln���쵥 = True Then
                    dblPrice = num�ۼ� * num����ϵ��
                ElseIf mbln�ƿ���ȷ���� = False Then
                    dblPrice = Get���ۼ�(lng����ID, cboStock.ItemData(cboStock.ListIndex), lng����, num����ϵ��)
                End If
            Else
                If bln���� = True Then
                    dblPrice = rsprice!�����ۼ�
                Else
                    dblPrice = rsprice!ƽ�����ۼ�
                End If
            End If
        End If

        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_���) = str���
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_��λ) = str��λ
        .TextMatrix(intRow, mBillCol.c_����) = str����
        .TextMatrix(intRow, mBillCol.C_Ч��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_һ���Բ���) = zlStr.NVL(rsЧ��!һ���Բ���)
        .TextMatrix(intRow, mBillCol.C_���Ч��) = zlStr.NVL(rsЧ��!���Ч��)
        
        .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mBillCol.C_�ⷿ����) = int�ⷿ����
        .TextMatrix(intRow, mBillCol.C_��������) = Format(num�������� / num����ϵ��, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mBillCol.C_ָ�������) = numָ�������
        .TextMatrix(intRow, mBillCol.C_����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = str��׼�ĺ�
        
        If (mbln��ȷ���� = True And mbln���쵥 = True) Or mbln���쵥 = False Then
            .TextMatrix(intRow, mBillCol.c_����) = lng����
        Else
            .TextMatrix(intRow, mBillCol.c_����) = 0
        End If
        If int�Ƿ��� = 1 Then
            .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(dblPrice, mFMT.FM_���ۼ�)
        End If
        Call CheckLapse(strЧ��)
        SetInputFormat intRow
        
    End With
    Call ��ʾ�����
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'������Ŀ¼��ȡֵ��������Ӧ����
Private Function SetRequestColValue(ByVal intRow As Integer, ByVal lng����ID As Long, _
    ByVal str���� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal str���ʧЧ�� As String, ByVal int���Ч�� As Integer, ByVal int�ⷿ���� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal numָ������� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal int���÷��� As Integer, ByVal str��׼�ĺ� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbltotal As Double
    Dim dblPrice As Double
    Dim intLop As Integer
    Dim rsprice As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rsЧ�� As ADODB.Recordset
    Dim bln���� As Boolean
    
    On Error GoTo ErrHandle
    If str���ʧЧ�� <> "" Then
        If Format(str���ʧЧ��, "yyyy-mm-dd") <= Format(sys.Currentdate, "yyyy-mm-dd") Then
            If MsgBox("���ġ�" & str���� & "(" & lng���� & ")�������Ч���Ѿ�����,�Ƿ�Ҫ�����ƿ�?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    gstrSQL = "Select һ���Բ���,���Ч�� from �������� where ����id=[1]"
    Set rsЧ�� = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, lng����ID)
    
    SetRequestColValue = False
    With mshBill
        Dim lngRow As Long
        For lngRow = 1 To .Rows - 1
            If lngRow <> intRow And .TextMatrix(lngRow, 0) <> "" Then
                If .TextMatrix(lngRow, 0) = lng����ID And Val(.TextMatrix(lngRow, mBillCol.c_����)) = lng���� Then
                    If UBound(Split(mstr�ظ�����, "��")) < 3 Then mstr�ظ����� = mstr�ظ����� & str���� & "��"  '����¼�����ظ�������
                    'Call MsgBox("�������ϡ�" & str���� & "(" & lng���� & ")���Ѿ����ڣ���ϲ��������ӣ�", vbOKOnly + vbInformation + vbDefaultButton2, gstrSysName)
                    Exit Function
                End If
            End If
        Next
        
        If lng���� > 0 Then   '���Ƴ��ⷿ�ǿⷿ�������ǿⷿ���������ĵ��ж�
            If mint�༭״̬ = 1 Then
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And lng���� = .TextMatrix(intLop, mBillCol.c_����) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mBillCol.C_��д����)
                        End If
                    End If
                Next
                
                If dbltotal >= num�������� And dbltotal <> 0 Then
                    MsgBox "�����ĵĿ��ÿ��������û���ˣ������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            End If
        End If
        
        If int�Ƿ��� = 1 Then
            If int���÷��� = 0 Then
                If int�ⷿ���� = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From ��������˵�� " & _
                            "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex))
                    If rsTemp.RecordCount = 0 Then
                        bln���� = True
                    End If
                End If
            Else
                bln���� = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(���ۼ�,0)*" & num����ϵ�� & " as  �����ۼ�,ʵ�ʽ��/ʵ������* " & num����ϵ�� & " as ƽ�����ۼ�" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "       and ҩƷid=[2]" & _
                "       and ����=1 and ʵ������>0 and " & _
                "       nvl(����,0)=[3]"
            
            Set rsprice = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), lng����ID, lng����)
            If rsprice.EOF Then
                If (mbln��ȷ���� = True And mbln���쵥 = True) Or (mbln���쵥 = False And bln���� = False) Then
                    MsgBox "ʱ������û�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
                    Exit Function
                ElseIf mbln��ȷ���� = False And mbln���쵥 = True Then
                    dblPrice = num�ۼ� * num����ϵ��
                ElseIf mbln�ƿ���ȷ���� = False Then
                    dblPrice = Get���ۼ�(lng����ID, cboStock.ItemData(cboStock.ListIndex), lng����, num����ϵ��)
                End If
            Else
                If bln���� = True Then
                    dblPrice = rsprice!�����ۼ�
                Else
                    dblPrice = rsprice!ƽ�����ۼ�
                End If
            End If
        End If

        For intCol = 0 To .Cols - 1
            If intCol <> mBillCol.C_�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, mBillCol.C_�к�) = intRow
        .TextMatrix(intRow, 0) = lng����ID
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_���) = str���
        .TextMatrix(intRow, mBillCol.C_����) = str����
        .TextMatrix(intRow, mBillCol.c_��λ) = str��λ
        .TextMatrix(intRow, mBillCol.c_����) = str����
        .TextMatrix(intRow, mBillCol.C_Ч��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_���ʧЧ��) = Format(str���ʧЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mBillCol.C_һ���Բ���) = zlStr.NVL(rsЧ��!һ���Բ���)
        .TextMatrix(intRow, mBillCol.C_���Ч��) = zlStr.NVL(rsЧ��!���Ч��)
        
        .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(num�ۼ� * num����ϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(intRow, mBillCol.C_�ⷿ����) = int�ⷿ����
        .TextMatrix(intRow, mBillCol.C_��������) = Format(num��������, mFMT.FM_����)
        .TextMatrix(intRow, mBillCol.C_���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & int���÷���
        .TextMatrix(intRow, mBillCol.C_ʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mBillCol.C_ʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mBillCol.C_ָ�������) = numָ�������
        .TextMatrix(intRow, mBillCol.C_����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mBillCol.C_��׼�ĺ�) = str��׼�ĺ�
        '���깺���ƿ�����ȷ���εģ������ж��Ƿ������ƿ�
        .TextMatrix(intRow, mBillCol.c_����) = lng����
        
        If int�Ƿ��� = 1 Then
            .TextMatrix(intRow, mBillCol.C_�ۼ�) = Format(dblPrice, mFMT.FM_���ۼ�)
        End If
        Call CheckLapse(strЧ��)
        SetInputFormat intRow
        
    End With
    Call ��ʾ�����
    SetRequestColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    Dim rsData As ADODB.Recordset
    Dim bln���ⷿ As Boolean, bln����ⷿ As Boolean
    Dim bln������ As Boolean, bln���÷��� As Boolean
    '˵����1���Ƴ��ⷿΪ�ⷿ����Ϊ�ⷿ�������ģ�
    '         A����������ģ������������ڿ��ƿ�������ʱ����д���������˿�����������������⣬�Ҹõ㲻�ܿ�������Ӱ�죻
    '         B������ޣ�����ѡ����������
    '      2���Ƴ��ⷿ��Ϊ�ⷿ����Ϊ�ⷿ�������ģ�
    '         A���������ⷿΪ�ⷿ��������ѡ�����г���һ��û�����κ�Ч�ڣ�
    '                ��ʱ�������������κ�Ч��
    '         B���������ⷿ��Ϊ�ⷿ����
    '                ��ʱ���������������κ�Ч��
    '      3�����Ĳ�Ϊ�ⷿ��������
    '         ��ʱ���������������κ�Ч��
    
'    If mblnEdit = False Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If Val(mshBill.TextMatrix(intRow, 0)) = 0 Then Exit Sub
    
    With mshBill
        If .TextMatrix(intRow, mBillCol.C_�ⷿ����) = "0" Then  '���ǿⷿ�������ģ����������룬����������
            .ColData(mBillCol.c_����) = 5                    '��ֹ
            .ColData(mBillCol.C_Ч��) = 5
        Else
            If .TextMatrix(intRow, mBillCol.c_����) = "" Then        'And GetDrugUnit(cboEnterStock.ItemData(cboEnterStock.ListIndex), mfrmMain.Caption) = "���Ŀⵥλ"
                .ColData(mBillCol.c_����) = 4              '���ı�����
                If .TextMatrix(intRow, mBillCol.C_���Ч��) <> "" Then
                    If Split(.TextMatrix(intRow, mBillCol.C_���Ч��), "||")(0) <> 0 Then
                        .ColData(mBillCol.C_Ч��) = 2          '���������
                    Else
                        .ColData(mBillCol.C_Ч��) = 5
                    End If
                Else
                    .ColData(mBillCol.C_Ч��) = 5
                End If
            Else
                .ColData(mBillCol.c_����) = 5              '��ֹ
                .ColData(mBillCol.C_Ч��) = 5
            End If
        End If
        If .TextMatrix(intRow, mBillCol.C_һ���Բ���) = "1" Then
            .ColData(mBillCol.C_�������) = 5
            .ColData(mBillCol.C_���ʧЧ��) = 5
        Else
            .ColData(mBillCol.C_�������) = 5              '��ֹ
            .ColData(mBillCol.C_���ʧЧ��) = 5
        End If
        
        '���ⷿ���Ż����Ϊ�գ���ⷿ�����Ŀ��Զ����Ż���ؽ��б༭
        
        If mbln�����������Ų��ؿ��� = True Then
            '1����ѯҩƷ�����Ϣ
            gstrSQL = "Select �ϴ�����,�ϴβ��� From ҩƷ��� Where �ⷿid=[1] And ҩƷid=[2] and nvl(����,0) = [3] "
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ж�Ч��", Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_����)))
            
            '2����ⷿ�Ƿ����
            bln���ⷿ = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
            bln������ = (Val(.TextMatrix(intRow, mBillCol.C_�ⷿ����)) = 1)
            bln���÷��� = (Split(.TextMatrix(intRow, mBillCol.C_���Ч��), "||")(2) = 1)
            If ((bln���ⷿ And bln������) Or (Not bln���ⷿ And bln���÷���)) Then '��ⷿ����
                If (IsNull(rsData!�ϴ�����) Or rsData.EOF) Then '���ⷿ�޿�������Ϊ��
                    .ColData(mBillCol.c_����) = 4
                Else
                    .ColData(mBillCol.c_����) = 5
                End If
                If (IsNull(rsData!�ϴβ���) Or rsData.EOF) Then
                    .ColData(mBillCol.C_����) = 1
                Else
                    .ColData(mBillCol.C_����) = 5
                End If
            End If
        End If
        
    End With
End Sub

Private Sub msh����_DblClick()
    msh����_KeyDown vbKeyReturn, 0
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    
    With mshBill
    
        If KeyCode = vbKeyEscape Then
            msh����.Visible = False
            .SetFocus
        End If
        
        If .Col = mBillCol.C_���� Then
            If CheckQualifications(mlngModule, 1, msh����.TextMatrix(msh����.Row, 2)) = False Then
                Exit Sub
            End If
            
            If KeyCode = vbKeyReturn Then
                .TextMatrix(.Row, .Col) = msh����.TextMatrix(msh����.Row, 2)
                msh����.Visible = False
                
                gstrSQL = "select ��׼�ĺ� from ҩƷ�����̶��� where ��������=[1] and ҩƷid=[2]"
                Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, "msh����_KeyDown", .TextMatrix(.Row, .Col), .TextMatrix(.Row, 0))
                If rsProvider.RecordCount > 0 Then
                    .TextMatrix(.Row, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsProvider!��׼�ĺ�), "", rsProvider!��׼�ĺ�)
                Else
                    .TextMatrix(.Row, mBillCol.C_��׼�ĺ�) = ""
                End If
                
                .Col = mBillCol.c_����
                .SetFocus
            End If
        End If
    End With
End Sub

Private Sub msh����_LostFocus()
    If msh����.Visible Then
        msh����.Visible = False
    End If
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And stbThis.Tag <> "PY" Then
        Logogram stbThis, 0
        stbThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And stbThis.Tag <> "WB" Then
        Logogram stbThis, 1
        stbThis.Tag = Panel.Key
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
    Dim bln���ⷿ As Boolean, bln����ⷿ As Boolean
    Dim bln������ As Boolean, bln���÷��� As Boolean
    ValidData = False
    If cboEnterStock.ListCount = 0 Then
        cboEnterStock.SetFocus
        Exit Function
    End If
    If cboStock.ListCount = 0 Then
        cboStock.SetFocus
        Exit Function
    End If
    
    bln���ⷿ = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    bln����ⷿ = CheckStockProperty(cboStock.ItemData(cboStock.ListIndex))

    
    ValidData = False
    
    Dim intLop As Integer
    
    If txtNO.Locked = False Then
        If Trim(txtNO.Text) = "" Then
            ShowMsgBox "���ݺŲ���Ϊ��"
            Exit Function
        End If
        
        If InStr(1, txtNO.Text, "'") <> 0 Then
            ShowMsgBox "���ݺ��в��ܺ��зǷ��ַ�"
            Exit Function
        End If
        
        If LenB(StrConv(txtNO.Text, vbFromUnicode)) > txtNO.MaxLength Then
            ShowMsgBox "���ݺų���,���������" & CInt(txtNO.MaxLength / 2) & "�����֣���ò�Ҫ���֣���" & txtNO.MaxLength & "���ַ�!"
            txtNO.SetFocus
            Exit Function
        End If
    End If
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If cboEnterStock.ListCount = 0 Then
                MsgBox "��������������Ĳ��ţ�[������������]�е���������", vbInformation, gstrSysName
                Exit Function
            End If
            If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                MsgBox "����ⷿ���Ƴ��ⷿ��ͬ�ˣ�������ѡ��", vbInformation, gstrSysName
                If cboEnterStock.Enabled Then cboEnterStock.SetFocus
                Exit Function
            End If
            
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .Rows - 1
                If Trim(.TextMatrix(intLop, mBillCol.C_����)) <> "" Then
                    If Val(Trim(.TextMatrix(intLop, mBillCol.C_��д����))) = 0 Then
                        MsgBox "��" & intLop & "�����ĵ�����Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_��д����
                        Exit Function
                    End If
                    

                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mBillCol.c_����))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "��" & intLop & "�����ĵ����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.c_����
                        Exit Function
                    End If
                    
                    '˵����ֻ�������ⷿ�����ж�
                    '   1�����ⷿ��ҩ�����������������������Ϣ
                    '   2�����ҩ����ҩ������������������������Ϣ
                    bln������ = (Val(mshBill.TextMatrix(intLop, mBillCol.C_�ⷿ����)) = 1)
                    bln���÷��� = (Split(mshBill.TextMatrix(intLop, mBillCol.C_���Ч��), "||")(2) = 1)
                    
                    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                        If mbln�ƿ���ȷ���� = True Then '�������γ��⣬���ж����ź�Ч���Ƿ���д
                            If ((bln���ⷿ And bln������) Or (Not bln���ⷿ And bln���÷���)) Then
                                If Split(.TextMatrix(intLop, mBillCol.C_���Ч��), "||")(0) <> 0 Then
'                            If .TextMatrix(intLop, mBillCol.C_�ⷿ����) <> "0" And Split(.TextMatrix(intLop, mBillCol.C_���Ч��), "||")(0) <> 0 Then
                                    If .TextMatrix(intLop, mBillCol.c_����) = "" Or .TextMatrix(intLop, mBillCol.C_Ч��) = "" Then
                                        MsgBox "��" & intLop & "�е�������Ч������,����������ż�ʧЧ���������뵥���У�", vbInformation, gstrSysName
                                        mshBill.SetFocus
                                        .Row = intLop
                                        .MsfObj.TopRow = intLop
                                        If .TextMatrix(intLop, mBillCol.c_����) = "" Then
                                            .Col = mBillCol.c_����
                                        Else
                                            .Col = mBillCol.C_Ч��
                                        End If
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                            
                            '����/�޸�ʱ��������������ֹ����
                        If Not CompareUsableQuantity(intLop, Val(Trim(.TextMatrix(intLop, mBillCol.C_��д����))), True) Then
                            .SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            .Col = mBillCol.C_��д����
                            Exit Function
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_��д����)) > 9999999999# Then
                        MsgBox "��" & intLop & "�����ĵ���д�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_��д����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_ʵ������)) > 9999999999# Then
                        MsgBox "��" & intLop & "�����ĵ�ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mBillCol.C_ʵ������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mBillCol.C_�ɹ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_��д����) = 4, mBillCol.C_��д����, mBillCol.C_ʵ������)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconintcol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "�����ĵ��ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mBillCol.C_��д����) = 4, mBillCol.C_��д����, mBillCol.C_ʵ������)
                        Exit Function
                    End If
                    
                    If mbln�����������Ų��ؿ��� = True Then
                        If ((bln���ⷿ And bln������) Or (Not bln���ⷿ And bln���÷���)) And (.TextMatrix(intLop, mBillCol.c_����) = "" Or .TextMatrix(intLop, mBillCol.C_����) = "") And .TextMatrix(intLop, 0) <> "" Then
                            MsgBox "��" & intLop & "�У����ⷿ�Ƿ�����������¼�����źͲ��أ�", vbInformation, gstrSysName
                            .SetFocus
                            .Row = intLop
                            .MsfObj.TopRow = intLop
                            If .TextMatrix(intLop, mBillCol.c_����) = "" Then
                                .Col = mBillCol.c_����
                            Else
                                .Col = mBillCol.C_����
                            End If
                            Exit Function
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

'Private Function ReValidData() As Boolean
'    Dim intLop As Integer
'    Dim rsStock As New Recordset
'
'    With mshBill
'        If .TextMatrix(1, 0) <> "" Then         '�����з�����
'            For intLop = 1 To .Rows - 1
'                If Trim(.TextMatrix(intLop, mBillCol.C_����)) <> "" Then
'                    gstrSQL = "select * from ҩƷ��� where ҩƷid=" & .TextMatrix(intLop, 0)
'                End If
'            Next
'        Else
'            Exit Function
'        End If
'    End With
'End Function

Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lng��� As Long
    
    Dim lng�ⷿID As Long
    Dim lng��ⷿID As Long
    Dim lng����ID As Long
    Dim str���� As String
    Dim lng���� As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim str��д���� As Double
    Dim dbl�ɹ��� As Double
    Dim dblʵ������ As Double
    Dim dbl�ɱ���� As Double
    Dim dbl���ۼ� As Double
    Dim dbl���۽�� As Double
    Dim dbl��� As Double
    Dim strժҪ As String
    Dim str������ As String
    Dim str�������� As String
    Dim str����� As String
    Dim datAssessDate As String
    Dim str���Ч�� As String
    Dim n As Long
    
    Dim arrSQL As Variant
    Dim intRow As Integer
    
    arrSQL = Array()
    SaveCard = False
    
    '���õ����Ƿ��ڽ���༭����󣬱���������Ա�޸�
    If mint�༭״̬ = 2 Or blnǿ�Ʊ��� Then          '�޸�
        mstrTime_End = GetBillInfo(19, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Function
        End If
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    With mshBill
        chrNo = Trim(txtNO)
        lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 11 Then  ' Or mbln��������
            If chrNo <> "" Then
                If CheckNOExists(72, chrNo) Then Exit Function
            End If
        
            If chrNo = "" Then chrNo = sys.GetNextNo(72, lng�ⷿID)
            If IsNull(chrNo) Then Exit Function
        End If
        txtNO.Tag = chrNo
        
        lng��ⷿID = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        strժҪ = Trim(txtժҪ.Text)
        If Txt������ <> "" Then
            str������ = Txt������
        Else
            str������ = UserInfo.�û���
        End If
        If Txt�������� <> "" Then
            str�������� = Txt��������.Caption
        Else
            str�������� = Format(sys.Currentdate, "yyyy-mm-dd HH:MM:SS")
        End If
        str����� = Txt�����
        On Error GoTo ErrHandle
        
        If mint�༭״̬ = 2 Or blnǿ�Ʊ��� Then        '�޸�
            If Not mbln���쵥 Then
                gstrSQL = "zl_�����ƿ�_Delete('" & mstr���ݺ� & "')"
            Else
                gstrSQL = "zl_��������_Delete('" & mstr���ݺ� & "')"
            End If
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "0;" & vbCrLf & gstrSQL
        End If
            
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                lng����ID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mBillCol.C_����)
                str���� = .TextMatrix(intRow, mBillCol.c_����)
                lng���� = Val(.TextMatrix(intRow, mBillCol.c_����))
                strЧ�� = IIf(.TextMatrix(intRow, mBillCol.C_Ч��) = "", "", .TextMatrix(intRow, mBillCol.C_Ч��))
                str��д���� = Round(Val(.TextMatrix(intRow, mBillCol.C_��д����)) * Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                dblʵ������ = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), g_С��λ��.obj_���С��.����С��)
                
                If Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����)) <> 0 Then
                    
                    If Val(Format(Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����)) / Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), mFMT.FM_����)) = Val(.TextMatrix(intRow, mBillCol.C_��д����)) Then
                        If str��д���� = dblʵ������ Then
                            str��д���� = Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����))
                            dblʵ������ = str��д����
                        ElseIf str��д���� < dblʵ������ Then
                            str��д���� = Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����))
                        End If
                    End If
                End If
                                
                dbl�ɹ��� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ���)) / .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_���С��.�ɱ���С��)
                dbl�ɱ���� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ɹ����)), g_С��λ��.obj_���С��.���С��)
                dbl���ۼ� = Round(Val(.TextMatrix(intRow, mBillCol.C_�ۼ�)) / Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), g_С��λ��.obj_���С��.���ۼ�С��)
                dbl���۽�� = Round(Val(.TextMatrix(intRow, mconintcol�ۼ۽��)), g_С��λ��.obj_���С��.���С��)
                dbl��� = Round(Val(.TextMatrix(intRow, mconintcol���)), g_С��λ��.obj_���С��.���С��)
                str���Ч�� = IIf(.TextMatrix(intRow, mBillCol.C_���ʧЧ��) = "", "", .TextMatrix(intRow, mBillCol.C_���ʧЧ��))
                lng��� = 2 * intRow - 1
                
                'zl_�����ƿ�_INSERT( /*NO_IN*/, /*���_IN*/, /*�ⷿID_IN*/,
                '/*�Է�����ID_IN*/, /*����ID_IN*/, /*����_IN*/, /*��д����_IN*/,ʵ������/,
                '/*�ɱ���_IN*/, /*�ɱ����_IN*/, /*���ۼ�_IN*/, /*���۽��_IN*/,
                '/*���_IN*/, /*������_IN*/, /*����_IN*/, /*����_IN*/, /*Ч��_IN*/,/���Ч��_IN/
                '/*ժҪ_IN*/��������_IN );
                
                If Not mbln���쵥 Or blnǿ�Ʊ��� Then
                    gstrSQL = "zl_�����ƿ�_INSERT('" & chrNo & "'," & lng��� & "," & lng�ⷿID & "," & _
                         lng��ⷿID & "," & lng����ID & "," & lng���� & "," & str��д���� & "," & dblʵ������ & "," & _
                         dbl�ɹ��� & "," & dbl�ɱ���� & "," & dbl���ۼ� & "," & dbl���۽�� & "," & _
                         dbl��� & ",'" & str������ & "','" & str���� & "','" & _
                         str���� & "'," & _
                        IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & "," & _
                        IIf(str���Ч�� = "", "Null", "to_date('" & Format(str���Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ",'" & _
                        strժҪ & "',to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS')," & _
                        IIf(mstr�˲��� = "", "null", "'" & mstr�˲��� & "'") & "," & _
                        IIf(mstr�˲����� = "", "Null", "to_date('" & Format(mstr�˲�����, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')") & ")"
                Else
                    gstrSQL = "zl_��������_INSERT('" & _
                        chrNo & "'," & _
                        lng��� & "," & _
                        lng�ⷿID & "," & _
                        lng��ⷿID & "," & _
                        lng����ID & "," & _
                        lng���� & "," & _
                        str��д���� & "," & _
                        dblʵ������ & "," & _
                        dbl�ɹ��� & "," & _
                        dbl�ɱ���� & "," & _
                        dbl���ۼ� & "," & _
                        dbl���۽�� & "," & _
                        dbl��� & ",'" & _
                        str������ & "','" & _
                        str���� & "','" & _
                        str���� & "'," & _
                        IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & "," & _
                        IIf(str���Ч�� = "", "Null", "to_date('" & Format(str���Ч��, "yyyy-MM-dd") & "','yyyy-mm-dd')") & ",'" & _
                        strժҪ & "',to_date('" & str�������� & "','yyyy-mm-dd HH24:MI:SS')" & "," & _
                        IIf(mstr�˲��� = "", "null,", "'" & mstr�˲��� & "',") & _
                        IIf(mstr�˲����� = "", "Null", "to_date('" & Format(mstr�˲�����, "yyyy-MM-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')") & ")"
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = CStr(lng����ID) & ";" & vbCrLf & gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSQL, mstrCaption, False, Not blnǿ�Ʊ���) Then Exit Function
'        If Not ��鵥��(19, txtNO.Tag) Then
'            If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
'            Exit Function
'        End If
        If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    Dim int�д� As Integer
    Dim intԭ��¼״̬ As Integer
    Dim strNo As String
    Dim str��� As Integer
    Dim lng����ID As Long
    Dim dbl�������� As Double
    Dim str������ As String
    Dim str��������  As String
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim int����� As Integer, lng�ⷿID As Long, lng���� As Long
    Dim n As Long
    
    SaveStrike = False
    
    With mshBill
        strNo = Trim(txtNO.Tag)
        lng�ⷿID = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        int����� = Get������(lng�ⷿID)
    
        '����������������С����
        For intRow = 1 To .Rows - 1
            If Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mBillCol.C_��д����)), Val(.TextMatrix(intRow, mBillCol.C_ʵ������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
                If int����� <> 0 Then
                    dbl�������� = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_ɢװС��.����С��)
                    If Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) = Val(.TextMatrix(intRow, mBillCol.C_��д����)) Then
                        dbl�������� = Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����))
                    End If
                    lng���� = ȡ��������(19, strNo, Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_���)) + 1)
                    If Check��������(lng�ⷿID, Val(.TextMatrix(intRow, 0)), lng����, dbl��������, int�����, IIf(mint������ʽ = 2, 1, 0)) = False Then Exit Function
                End If
                
            End If
        Next
        
        str������ = UserInfo.�û���
        str�������� = Format(sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        intԭ��¼״̬ = mint��¼״̬
        
        On Error GoTo ErrHandle
        gcnOracle.BeginTrans
        
        int�д� = 0
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
'        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) <> 0 Then
                int�д� = int�д� + 1
                
                lng����ID = .TextMatrix(intRow, 0)
                dbl�������� = Round(Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) * .TextMatrix(intRow, mBillCol.C_����ϵ��), g_С��λ��.obj_ɢװС��.����С��)
  
                If Val(.TextMatrix(intRow, mBillCol.C_ʵ������)) = Val(.TextMatrix(intRow, mBillCol.C_��д����)) Then
                    dbl�������� = Val(.TextMatrix(intRow, mBillCol.c_ԭʼ����))
                End If
                dbl�������� = IIf(mint�༭״̬ = 6 And mint������ʽ = 2, -1, 1) * dbl��������
                           
                str��� = .TextMatrix(intRow, mBillCol.c_���)
                
                'ZL_�����ƿ�_STRIKE(/*int�д�*/,/*intԭ��¼״̬*/,/*strNO*/,/*str���*/, /*lng����ID*/,
                '/*dbl��������*/,/*str������*/, /*str��������*/);
                gstrSQL = "ZL_�����ƿ�_STRIKE(" & int�д� & "," & intԭ��¼״̬ & ",'" & strNo & "'," & str��� & "," & lng����ID & "," & dbl�������� & ",'" _
                    & str������ & "',to_date('" & Format(str��������, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS') ," & mint������ʽ & ")"
                zlDatabase.ExecuteProcedure gstrSQL, mstrCaption
            End If
            recSort.MoveNext
        Next
        gcnOracle.CommitTrans
        
        If int�д� = 0 Then
            MsgBox "û��ѡ��һ�в��������������ܳ��������飡", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .Rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mBillCol.C_�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconintcol�ۼ۽��))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & Format(curTotal, mFMT.FM_���)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���)
    lblDifference.Caption = "��ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
End Sub

Private Sub ��ʾ�����()
    Dim rsUseCount As New Recordset
    Dim strNote As String
    
    On Error GoTo ErrHandle
    With mshBill
        If .TextMatrix(.Row, mBillCol.C_����) = "" Then
            stbThis.Panels(2).Text = ""
            Exit Sub
        End If
        If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
        
        If mint�༭״̬ <> 10 Then
            If mbln���쵥 Then '���쵥
                If mbln��ȷ���� Then
                    gstrSQL = "" & _
                        "   Select ��������/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as  �������� " & _
                        "   From ҩƷ��� " & _
                        "   Where �ⷿid=[1]" & _
                        "          and ҩƷid=[2]" & _
                        "           and ����=1 and " & _
                        "          nvl(����,0)=[3]"
                Else
                    gstrSQL = "" & _
                        "   Select Sum(��������)/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as  �������� " & _
                        "   From ҩƷ��� " & _
                        "   Where �ⷿid=[1]" & _
                        "          and ҩƷid=[2]" & _
                        "           and ����=1  "
                End If
            Else '�ƿⵥ
                If mbln�ƿ���ȷ���� Then
                    gstrSQL = "" & _
                        "   Select ��������/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as  �������� " & _
                        "   From ҩƷ��� " & _
                        "   Where �ⷿid=[1]" & _
                        "          and ҩƷid=[2]" & _
                        "           and ����=1 and " & _
                        "          nvl(����,0)=[3]"
                Else
                    gstrSQL = "" & _
                        "   Select Sum(��������)/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as  �������� " & _
                        "   From ҩƷ��� " & _
                        "   Where �ⷿid=[1]" & _
                        "          and ҩƷid=[2]" & _
                        "           and ����=1  "
                End If
                    
            End If
        
                
                Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mBillCol.c_����)))
                
                If rsUseCount.EOF Then
                    .TextMatrix(.Row, mBillCol.C_��������) = 0
                Else
                    .TextMatrix(.Row, mBillCol.C_��������) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
                End If
                rsUseCount.Close
                
                stbThis.Panels(2).Text = "�����ĵ�ǰ�����Ϊ[" & Format(.TextMatrix(.Row, mBillCol.C_��������), mFMT.FM_����) & "]" & .TextMatrix(.Row, mBillCol.c_��λ)
        Else
            '���ڷ���ʱ����ʾ��ҩƷ�����пⷿ�Ŀ�棬�Ա��ڿⷿ��Ա����ʵ�ʵķ�������
            gstrSQL = "" & _
            "   Select B.���� AS �ⷿ,Nvl(A.��������,0)/" & .TextMatrix(.Row, mBillCol.C_����ϵ��) & " as �������� " & _
            "   From ҩƷ��� A,���ű� B" & _
            "    Where A.�ⷿID=B.ID And A.ҩƷid=[1]" & _
            "           And A.����=1 "
            
            Set rsUseCount = zlDatabase.OpenSQLRecord(gstrSQL, "��ʾ�����", Val(.TextMatrix(.Row, 0)))
            With rsUseCount
                Do While Not .EOF
                    strNote = strNote & "," & !�ⷿ & ":" & Format(zlStr.NVL(!��������, 0), mFMT.FM_����) & mshBill.TextMatrix(mshBill.Row, mBillCol.c_��λ)
                    .MoveNext
                Loop
            End With
            stbThis.Panels(2).Text = Mid(strNote, 2)
        End If
    End With
    Exit Sub
ErrHandle:
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
    Dim int��װϵ�� As Integer
    Dim lngҩƷID As Long
    Dim blnInput As Boolean
    
    '��ʼ׼��
    intNO = 68
    lng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtIn.Text) = "" Then Exit Sub
    
    If Len(txtIn.Text) < 8 Then
        txtIn.Text = zlCommFun.GetFullNO(txtIn.Text, intNO, lng�ⷿID)
    End If
    
    '��ҪҪ������е�������
    For IntCheck = 1 To mshBill.Rows - 1
        If mshBill.TextMatrix(IntCheck, 0) <> "" Then
            Exit For
        End If
    Next
    If IntCheck <> mshBill.Rows Then
        If MsgBox("��ҪҪ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '����ҩƷ��λ�ı�
        mshBill.ClearBill
    End If
    
    gstrSQL = "select �շ�ϸĿid,ִ�п���id from �շ�ִ�п���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ�洢�ⷿ")
    
    '��ȡ�õ��ݲ���ձ��ֻ������ȡ�������ݣ��ҷ��˻�����
    gstrSQL = "Select a.ҩƷid As ����id, '[' || c.���� || ']' As ����, '[' || c.���� || ']' || Nvl(f.����, c.����) As ҩƷ����, c.���� As ͨ����, f.���� As ��Ʒ��," & vbNewLine & _
                "       c.���, a.����, c.���㵥λ As ���۵�λ, 1 As ����ϵ��, b.��װ��λ, b.����ϵ��, Nvl(a.����, 0) As ����, Nvl(c.�Ƿ���, 0) As ʱ��," & vbNewLine & _
                "       Nvl(b.�ⷿ����, 0) As �ⷿ����, Nvl(b.���÷���, 0) As ���÷���, b.���Ч��, a.����, a.Ч��, a.���Ч��, b.���Ч��, b.ָ�������, a.ʵ������, d.��������," & vbNewLine & _
                "       d.ʵ�ʽ��, d.ʵ�ʲ��, e.�ּ�, a.��׼�ĺ�, Nvl(d.ƽ���ɱ���, 0) As ƽ���ɱ���, a.��ҩ��λid" & vbNewLine & _
                "From ҩƷ�շ���¼ A, �������� B, �շ���ĿĿ¼ C, ҩƷ��� D, �շѼ�Ŀ E, �շ���Ŀ���� F" & vbNewLine & _
                "Where a.ҩƷid = b.����id And b.����id = c.Id And b.����id = d.ҩƷid(+) And b.����id = f.�շ�ϸĿid(+) And f.����(+) = 3 And f.����(+) = 1 And" & vbNewLine & _
                "      b.����id = e.�շ�ϸĿid(+) And Sysdate >= e.ִ������(+) And Sysdate <= Nvl(e.��ֹ����(+), Sysdate) And d.�ⷿid(+) = [2] And" & vbNewLine & _
                "      d.����(+) = 1 And Nvl(a.����, 0) = Nvl(d.����, 0) And a.���� = 15 And a.��¼״̬ = 1 And Nvl(a.��ҩ��ʽ, 0) = 0 And" & vbNewLine & _
                "      a.������� Is Not Null And a.No = [1] And a.�ⷿid + 0 = [2]" & vbNewLine & _
                GetPriceClassString("E") & "Order By a.���"

    Set rsBill = zlDatabase.OpenSQLRecord(gstrSQL, mstrCaption & "[��ȡ�⹺��ⵥ]", txtIn.Text, Me.cboStock.ItemData(Me.cboStock.ListIndex))
             
    If rsBill.RecordCount = 0 Then
        MsgBox "û���ҵ����⹺��ⵥ�ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsBill
        intRow = 1
        Do While Not .EOF
            lngҩƷID = !����ID
            rsTemp.Filter = " �շ�ϸĿid=" & lngҩƷID & " and ִ�п���id=" & lng�ⷿID
            If rsTemp.RecordCount = 0 Then
                MsgBox "����[" & !ҩƷ���� & "]δ��" & cboStock.Text & "�����ô洢���ԣ��������ƿ⣡"
                blnInput = True
            End If
            rsTemp.Filter = ""
            rsTemp.Filter = " �շ�ϸĿid=" & lngҩƷID & " and ִ�п���id=" & cboEnterStock.ItemData(cboEnterStock.ListIndex)
            If rsTemp.RecordCount = 0 Then
                MsgBox "����[" & !ҩƷ���� & "]δ��" & cboEnterStock.Text & "�����ô洢���ԣ��������ƿ⣡"
                blnInput = True
            End If
            
            If blnInput = False Then
                '����ƻ����൱�ڶ��ǰ������ƿ⣬��Ҫ��װ������ǰ���ȼ����
                If !ʵ������ > !�������� Then
                    Select Case mint�����
                    Case 1
                        If MsgBox(!ҩƷ���� & "��治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            blnInput = True
                        End If
                    Case 2
                        MsgBox !ҩƷ���� & "��治�㣬�������ƿ⣡", vbInformation, gstrSysName
                        blnInput = True
                    End Select
                End If
            End If
            
            'װ������(SetColValue)
            If blnInput = False Then
                int��װϵ�� = IIf(mintUnit = 0, 1, !����ϵ��)
                If Not SetColValue(intRow, !����ID, "[" & !���� & "]" & !ͨ����, _
                   NVL(!���), NVL(!����), IIf(mintUnit = 0, !���۵�λ, !��װ��λ), _
                    NVL(!�ּ�, 0), NVL(!����), NVL(!Ч��), IIf(IsNull(!���Ч��), "", Format(!���Ч��, "yyyy-MM-dd")), _
                    NVL(!���Ч��, 0), !�ⷿ����, NVL(!��������, 0), NVL(!ʵ�ʽ��, 0), NVL(!ʵ�ʲ��, 0), _
                    IIf(IsNull(!ָ�������), "0", !ָ�������), int��װϵ��, NVL(!����, 0), !ʱ��, _
                    !���÷���, IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)) Then
                    mshBill.ClearBill
                    Exit Sub
                End If

                '��д�������ɹ��ۡ��ۼ۵���
                mshBill.TextMatrix(intRow, mBillCol.C_�к�) = intRow
                mshBill.TextMatrix(intRow, mBillCol.C_��д����) = Format(!ʵ������ / int��װϵ��, mFMT.FM_����)
                mshBill.TextMatrix(intRow, mBillCol.C_ʵ������) = Format(!ʵ������ / int��װϵ��, mFMT.FM_����)
                mshBill.TextMatrix(intRow, mBillCol.C_�ɹ���) = Format(!ƽ���ɱ��� * int��װϵ��, mFMT.FM_�ɱ���)
                mshBill.TextMatrix(intRow, mBillCol.C_�ɹ����) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_�ɹ���)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_ʵ������)), mFMT.FM_���)
                mshBill.TextMatrix(intRow, mconintcol�ۼ۽��) = Format(Val(mshBill.TextMatrix(intRow, mBillCol.C_�ۼ�)) * Val(mshBill.TextMatrix(intRow, mBillCol.C_ʵ������)), mFMT.FM_���)
                mshBill.TextMatrix(intRow, mconintcol���) = Format(Val(mshBill.TextMatrix(intRow, mconintcol�ۼ۽��)) - mshBill.TextMatrix(intRow, mBillCol.C_�ɹ����), mFMT.FM_���)

                intRow = intRow + 1
                mshBill.Rows = mshBill.Rows + 1
            End If
            blnInput = False
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
    
    OS.OpenIme (True)
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
    OS.OpenIme False
End Sub

'������������бȽ�
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double, Optional ByVal blnSave As Boolean = False) As Boolean
    Dim dblUsableQuantity As Double      'ʵ��������Ӧ���������
    Dim numUsedCount As Double
    Dim vardrug As Variant
    Dim dbltotal As Double              'ĳ�������������������
    Dim intLop As Integer
    Dim rsCheck As ADODB.Recordset
    Dim strSaveCheck As String
    
    'mint�����: 0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    
    CompareUsableQuantity = False
    If Not mbln�ƿ���ȷ���� Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        If Not blnSave Then
            dblUsableQuantity = Format(.TextMatrix(intRow, mBillCol.C_��������), mFMT.FM_����)
        Else
            '�������޸ı���ʱ����ȡ���ݿ��еĿ�����������Ҫ��ֹ���������ͬʱ��Կ�������ȡֵ��Ӱ��
            gstrSQL = "Select Nvl(��������, 0) �������� From ҩƷ��� Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And Nvl(����, 0) = [3] "
            Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "CompareUsableQuantity", Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mBillCol.c_����)))
            
            If rsCheck.EOF Then
                dblUsableQuantity = 0
            Else
                dblUsableQuantity = Val(Format(rsCheck!�������� / Val(.TextMatrix(intRow, mBillCol.C_����ϵ��)), mFMT.FM_����))
                
                If dblUsableQuantity <> Val(Format(.TextMatrix(intRow, mBillCol.C_��������), mFMT.FM_����)) Then
                    .TextMatrix(intRow, mBillCol.C_��������) = dblUsableQuantity
                End If
                
                strSaveCheck = "����������������Ա�ռ����"
            End If
        End If
        
        If mint����� = 0 Then
            '0-�����
        ElseIf mint����� = 1 Then
            '1-��飬��������
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    If MsgBox("��" & intRow & "��[" & .TextMatrix(intRow, C_����) & "]��������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "��" & strSaveCheck & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If gSystem_Para.para_������¿��ÿ�� = False Then
                    '���û��Ԥ��������������������ԭʼ����
                    numUsedCount = 0
                End If
                
                If dbl��д���� > dblUsableQuantity + numUsedCount Then
                    If MsgBox("��" & intRow & "��[" & .TextMatrix(intRow, C_����) & "]��������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + numUsedCount & "��" & strSaveCheck & "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint����� = 2 Then
            '2-��飬�����ֹ
            If mint�༭״̬ = 1 Then
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mBillCol.c_����) = "", "0", .TextMatrix(intLop, mBillCol.c_����)) = "0" Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mBillCol.C_��д����)
                        End If
                    End If
                Next
                
                
                If dbl��д���� + dbltotal > dblUsableQuantity Then
                    MsgBox "��" & intRow & "��[" & .TextMatrix(intRow, C_����) & "]��������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity - dbltotal & "��" & strSaveCheck & "�������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 3 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mBillCol.c_����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbltotal = 0
                For intLop = 1 To .Rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mBillCol.c_����) = "", "0", .TextMatrix(intLop, mBillCol.c_����)) = "0" Then
                            dbltotal = dbltotal + Val(.TextMatrix(intLop, mBillCol.C_ʵ������))
                        End If
                    End If
                Next
                
                If gSystem_Para.para_������¿��ÿ�� = False Then
                    '���û��Ԥ��������������������ԭʼ����
                    numUsedCount = 0
                End If
                
                If dbl��д���� + dbltotal > dblUsableQuantity + numUsedCount Then
                    MsgBox "��" & intRow & "��[" & .TextMatrix(intRow, C_����) & "]��������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + numUsedCount - dbltotal & "��" & strSaveCheck & "�������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'��ӡ����
Private Sub printbill()
    Dim strNo As String
    strNo = txtNO.Tag
    FrmBillPrint.ShowMe Me, glngSys, "zl1_bill_1716", mint��¼״̬, mintUnit, 1716, "���ĵ�����", strNo
End Sub

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    
    zlDatabase.OpenRecordset rsBatchNolen, gstrSQL, "ȡ�ֶγ���"
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AutoExpend(Optional blnCheck As Boolean = False) As Boolean
    Dim lng�ⷿID As Long, lng����ID As Long, lng����ID_Last As Long, lng���� As Long
    Dim bln�ⷿ As Boolean, bln���� As Boolean, blnʱ�� As Boolean, blnAddRow As Boolean
    Dim dbl��д���� As Double, dbl�������� As Double, dbl���� As Double, dbl����ϵ�� As Double
    Dim dbl�ּ� As Currency, dbl�ּ�_ʱ�� As Double, dbl�ɱ��� As Double
    Dim lngCol As Long, lngCols As Long, lngRow As Long
    Dim intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim dblʵ������ As Double
        
    '�����ļ�¼�����Զ��ֽ⣬��������������
    On Error GoTo ErrHand
    Screen.MousePointer = 11
    lngRow = 1: lngCols = mshBill.Cols - 1
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln�ⷿ = CheckStockProperty(lng�ⷿID)
    
    Do While True
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl�������� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_��д����))
        dbl��д���� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������))
        dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��))
        lng���� = Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
        If lng����ID = 0 Then Exit Do
        
        '��ȡ�����Ķ��ڳ���ⷿ�Ƿ������ʱ�۵�����
        If lng����ID <> lng����ID_Last Then
            lng����ID_Last = lng����ID
            gstrSQL = " Select Nvl(A.�ⷿ����,0) �ⷿ����,Nvl(A.���÷���,0) ���÷���," & _
                      " Nvl(B.�Ƿ���,0) ʱ��,Nvl(P.�ּ�,0) �ּ�,Nvl(A.�ɱ���,0) �ɱ���" & _
                      " From �������� A,�շ���ĿĿ¼ B,�շѼ�Ŀ P" & _
                      " Where A.����ID = B.ID And B.ID=P.�շ�ϸĿID And A.����ID =[1] " & _
                      " And Sysdate between P.ִ������ And Nvl(P.��ֹ����,Sysdate)" & _
                      GetPriceClassString("P")
                      
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��϶��ڳ���ⷿ�Ƿ������ʱ�۵�����", lng����ID)
                      
            blnʱ�� = (rsTemp!ʱ�� = 1)
            dbl�ּ� = rsTemp!�ּ� * dbl����ϵ��
            dbl�ɱ��� = rsTemp!�ɱ��� * dbl����ϵ��
            bln���� = IIf(bln�ⷿ, (rsTemp!�ⷿ���� = 1), (rsTemp!���÷��� = 1))
        End If
        
        '����������Ƿ������ģ�������Ϊ�㣬��˵����Ҫ�Զ��ֽ�
        blnAddRow = False
        If bln���� And lng���� = 0 Then
            If blnCheck Then
                If dbl��д���� > Val(mshBill.TextMatrix(lngRow, mBillCol.C_��������)) Then
                    MsgBox "��" & lngRow & "�е����������λ�ʱ�����ģ��������ĵ�ǰ��治�㣬���ܼ�����", vbInformation, gstrSysName
                    Screen.MousePointer = 0: Exit Function
                End If
            End If
            gstrSQL = " Select Nvl(��������,0)/" & dbl����ϵ�� & " As ��������,Nvl(ʵ������,0)/" & dbl����ϵ�� & " As ʵ������," & _
                      " Nvl(ʵ�ʽ��,0) ʵ�ʽ��,Nvl(ʵ�ʲ��,0) ʵ�ʲ��,ƽ���ɱ���,nvl(���ۼ�,0) * " & dbl����ϵ�� & " as ���ۼ�," & _
                      " Nvl(����,0) ����,�ϴ����� ����,to_char(Ч��,'yyyy-MM-dd') Ч��,�ϴβ��� ����,��׼�ĺ�" & _
                      " From ҩƷ��� Where nvl(��������,0)<>0   and �ⷿID=[1] And ҩƷID=[2]  And ����=1 "
            If gSystem_Para.P156_�����㷨 = 0 Then '���λ���Ч�������ȳ���
                gstrSQL = gstrSQL & " Order by Nvl(����, 0)"
            Else
                gstrSQL = gstrSQL & " Order by Ч��,Nvl(����, 0)"
            End If
            
            Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������ָ���������п���¼", lng�ⷿID, lng����ID)
                      
            intCount = 0
            With rsCheck
                Do While Not .EOF
                    '����д��¼
                    intCount = intCount + 1
                    blnAddRow = False
                    If .AbsolutePosition <> 1 Then
                        Call InsertRow(lngRow)
                        For lngCol = 0 To lngCols
                            mshBill.TextMatrix(lngRow, lngCol) = mshBill.TextMatrix(lngRow - 1, lngCol)
                        Next
                        mshBill.RowData(lngRow) = mshBill.RowData(lngRow - 1)
                    End If
   
                    If intCount = 1 Then
                        dblʵ������ = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������))
                    End If
                    '��д���������Ϣ
                    mshBill.TextMatrix(lngRow, mBillCol.C_�к�) = lngRow
                    mshBill.TextMatrix(lngRow, mBillCol.c_���) = (lngRow - 1) * 2 + 1
                    mshBill.TextMatrix(lngRow, mBillCol.c_����) = rsCheck!����
                    mshBill.TextMatrix(lngRow, mBillCol.c_����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mBillCol.C_Ч��) = IIf(IsNull(rsCheck!Ч��), "", rsCheck!Ч��)
                    mshBill.TextMatrix(lngRow, mBillCol.C_����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mBillCol.C_��׼�ĺ�) = IIf(IsNull(rsCheck!��׼�ĺ�), "", rsCheck!��׼�ĺ�)
                    
                    '���¼���۸������Ϣ
                    If blnʱ�� = True Then
                        If bln���� = True Then
                            dbl�ּ�_ʱ�� = rsCheck!���ۼ�
                        Else
                            If rsCheck!ʵ������ > 0 Then
                                dbl�ּ�_ʱ�� = rsCheck!ʵ�ʽ�� / rsCheck!ʵ������
                            Else
                                dbl�ּ�_ʱ�� = dbl�ּ�
                            End If
                        End If
                    End If
                    
                    If dbl��д���� <= rsCheck!�������� Then
                        dbl���� = dbl��д����
                    Else
                        dbl���� = rsCheck!��������
                    End If
                    If dbl���� > dbl��д���� Then dbl���� = dbl��д����
                    
                    If dblʵ������ <> mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������) Then
                        mshBill.TextMatrix(lngRow, mBillCol.c_ԭʼ����) = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������)) * Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��))
                    End If
                    
                    mshBill.TextMatrix(lngRow, mBillCol.C_��д����) = Format(dbl����, mFMT.FM_����)
                    mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������) = Format(dbl����, mFMT.FM_����)
                    
                    If Trim(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������)) = "" Then mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������) = 0
                    mshBill.TextMatrix(lngRow, mBillCol.C_ʵ�ʲ��) = Format(rsCheck!ʵ�ʲ��, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_ʵ�ʽ��) = Format(rsCheck!ʵ�ʽ��, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_��������) = Format(rsCheck!��������, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�) = Format(IIf(blnʱ��, dbl�ּ�_ʱ��, dbl�ּ�), mFMT.FM_���ۼ�)
                    mshBill.TextMatrix(lngRow, mconintcol�ۼ۽��) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ�)) * dbl����, mFMT.FM_���)
                    
'                    If rsCheck!ʵ�ʽ�� > 0 Then
'                        mshBill.TextMatrix(lngRow, mBillCol.C_���) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��)) * rsCheck!ʵ�ʲ�� / rsCheck!ʵ�ʽ��, mFMT.FM_���)
'                    Else
'                        mshBill.TextMatrix(lngRow, mBillCol.C_���) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��)) * Val(mshBill.TextMatrix(lngRow, mBillCol.C_ָ�������)) / 100, mFMT.FM_���)
'                    End If
'                    mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ۼ۽��)) - Val(mshBill.TextMatrix(lngRow, mBillCol.C_���)), mFMT.FM_���)
                    
'                    If dbl���� <> 0 Then
'                        mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����)) / dbl����, mFMT.FM_�ɱ���)
'                    Else
'                        mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
'                    End If
                    '�����µķ�ʽ����ɱ��� �ɱ���=ҩƷ���.ƽ���ɱ���
                    mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���) = Format(Get�ɱ���(lng����ID, lng�ⷿID, Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))) * dbl����ϵ��, mFMT.FM_�ɱ���)
                    mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = Format(Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ���)) * dbl����, mFMT.FM_���)
                    mshBill.TextMatrix(lngRow, mconintcol���) = Format(Val(mshBill.TextMatrix(lngRow, mconintcol�ۼ۽��)) - Val(mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����)), mFMT.FM_���)
                    
                    dbl��д���� = dbl��д���� - dbl����
                    dbl�������� = dbl�������� - dbl����
                    If dbl��д���� = 0 Then Exit Do
                    lngRow = lngRow + 1
                    blnAddRow = True
                    .MoveNext
                Loop
                If dbl�������� <> 0 And rsCheck.RecordCount <> 0 Then
                    If blnAddRow Then
                        mshBill.TextMatrix(lngRow - 1, mBillCol.C_��д����) = Format(dbl�������� + dbl����, mFMT.FM_����)
                    Else
                        mshBill.TextMatrix(lngRow, mBillCol.C_��д����) = Format(dbl�������� + dbl����, mFMT.FM_����)
                    End If
                End If
            End With
            
            '�������¼Ϊ�㣬��˵��δ���зֽ⣬��Ҫ������������ʵ��������Ϊ��
            If dbl��д���� <> 0 And rsCheck.RecordCount = 0 Then
                mshBill.TextMatrix(lngRow, mBillCol.C_�к�) = lngRow
                mshBill.TextMatrix(lngRow, mBillCol.c_���) = (lngRow - 1) * 2 + 1
                mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������) = ""
                mshBill.TextMatrix(lngRow, mconintcol�ۼ۽��) = ""
                mshBill.TextMatrix(lngRow, mBillCol.C_�ɹ����) = ""
                mshBill.TextMatrix(lngRow, mconintcol���) = ""
            End If
        Else
            mshBill.TextMatrix(lngRow, mBillCol.C_�к�) = lngRow
            mshBill.TextMatrix(lngRow, mBillCol.c_���) = (lngRow - 1) * 2 + 1
        End If
        If blnAddRow = False Then lngRow = lngRow + 1
    Loop
    
    AutoExpend = True
    Screen.MousePointer = 0
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub InsertRow(ByVal lngRow As Long)
    Dim lngReserve As Long, lngRows As Long
    Dim lngCol As Long, lngCols As Long
    lngReserve = lngRow
    lngRows = mshBill.Rows - 1
    lngCols = mshBill.Cols - 1
    mshBill.Rows = mshBill.Rows + 1
    
    '����ǰ�м�������ȫ������
    For lngRow = lngRows To lngReserve Step -1
        For lngCol = 0 To lngCols
            mshBill.TextMatrix(lngRow + 1, lngCol) = mshBill.TextMatrix(lngRow, lngCol)
        Next
        mshBill.RowData(lngRow + 1) = mshBill.RowData(lngRow)
        'У���к�
        mshBill.TextMatrix(lngRow + 1, mBillCol.C_�к�) = lngRow + 1
    Next
End Sub

Private Sub ShowColor(Optional ByVal lngCurRow As Long = 0)
    '�ڲ��Ļ����ʱ������治��ļ�¼�԰���ɫ��ʾ����
    Dim lngSelect_Row  As Long, lngSelect_Col As Long
    Dim lng�ⷿID As Long, lng����ID As Long, lng����ID_Last As Long, lng���� As Long
    Dim bln�ⷿ As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim dbl��д���� As Double, dbl���� As Double, dbl����ϵ�� As Double
    Dim dbl�ּ� As Currency, dbl�ּ�_ʱ�� As Currency
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    
    Dim lngRow As Long
    On Error GoTo ErrHand
    
    mshBill.Redraw = False
    lngSelect_Row = mshBill.Row: lngSelect_Col = mshBill.Col
    lngRow = IIf(lngCurRow > 0, lngCurRow, 1)
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln�ⷿ = CheckStockProperty(lng�ⷿID)
    
    Do While True
        If lngRow > mshBill.Rows - 1 Then Exit Do
        mshBill.Row = lngRow: mshBill.Col = mBillCol.C_����
        mshBill.MsfObj.CellForeColor = &H0&
    
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl��д���� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_��д����))
        dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��))
        lng���� = Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
        If lng����ID = 0 Then Exit Do
        
        '��ȡ�ò��϶��ڳ���ⷿ�Ƿ������ʱ�۵�����
        If lng����ID <> lng����ID_Last Then
            lng����ID_Last = lng����ID
            gstrSQL = "" & _
                "   Select Nvl(A.�ⷿ����,0) �ⷿ����,Nvl(A.���÷���,0) ���÷���,Nvl(B.�Ƿ���,0) ʱ��,Nvl(P.�ּ�,0) �ּ� " & _
                "   From �������� A,�շ���ĿĿ¼ B,�շѼ�Ŀ P" & _
                "   Where A.����ID = B.ID And B.ID=P.�շ�ϸĿID And A.����ID = [1]" & _
                "           And Sysdate between P.ִ������ And Nvl(P.��ֹ����,Sysdate)" & _
                GetPriceClassString("P")
                
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ò��϶��ڳ���ⷿ�Ƿ������ʱ�۵�����", lng����ID)
            
            dbl�ּ� = rsTemp!�ּ�
            blnʱ�� = (rsTemp!ʱ�� = 1)
            bln���� = IIf(bln�ⷿ, (rsTemp!�ⷿ���� = 1), (rsTemp!���÷��� = 1))
        End If
        
        '���������������������������������Ԫ����ɫ
        If bln���� And lng���� <> 0 Then
            '����������Ƿ������ģ���ָ������
            gstrSQL = "" & _
                "   Select Nvl(��������,0)/" & dbl����ϵ�� & " As ��������,Nvl(ʵ������,0)/" & dbl����ϵ�� & " As ʵ������," & _
                "           Nvl(ʵ�ʽ��,0) ʵ�ʽ��,Nvl(ʵ�ʲ��,0) ʵ�ʲ��" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿID=[1] And ҩƷID=[2] And ����=1 And Nvl(����,0)=[3]"
        Else
            'δָ�����λ򲻷��������ģ�ֱ�ӽ�����ⷿ���������п���¼�ۼ�
            gstrSQL = "" & _
                "   Select ҩƷid ����ID,Sum(Nvl(��������,0))/" & dbl����ϵ�� & " As ��������,Sum(Nvl(ʵ������,0))/" & dbl����ϵ�� & " As ʵ������," & _
                "           Sum(Nvl(ʵ�ʽ��,0)) ʵ�ʽ��,Sum(Nvl(ʵ�ʲ��,0)) ʵ�ʲ��" & _
                "   From ҩƷ��� Where �ⷿID=[1] And ҩƷID=[2] And ����=1 " & _
                "   Group by ҩƷID"
        End If
        
        Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������ָ���������п���¼", lng�ⷿID, lng����ID, lng����)
        
        If rsCheck.EOF Then
            mshBill.MsfObj.CellForeColor = &H400040
        Else
            If rsCheck!�������� < dbl��д���� Then
                mshBill.MsfObj.CellForeColor = &H400040
            End If
        End If
        If lngCurRow > 0 Then Exit Do
        lngRow = lngRow + 1
    Loop
    
    mshBill.Row = lngSelect_Row: mshBill.Col = lngSelect_Col
    mshBill.Redraw = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshBill.Redraw = True
End Sub
Private Function CheckStockProperty(ByVal lng�ⷿID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '���ָ���ⷿ�ǿⷿ�����ϲ��Ż����Ƽ���(����Ŀⷿ�϶��ǿⷿ�����ϲ��Ż��Ƽ����е�һ��)
    On Error GoTo ErrHandle
    gstrSQL = " Select ����ID From ��������˵�� " & _
              " Where (�������� like '���ϲ���' Or �������� like '%�Ƽ���') And ����id=[1]"
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��ǲ��ǿⷿ���Ƽ���", lng�ⷿID)
              
    If rsCheck.EOF Then
        CheckStockProperty = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckStock() As Boolean
    Dim dbl����ϵ�� As Double, dbl�������� As Double, dbl��д���� As Double
    Dim lngRow As Long, lngRows As Long, int����� As Integer
    Dim lng����ID As Long, lng�ⷿID As Long, lng���� As Long
    Dim bln�ⷿ As Boolean, bln���� As Boolean
    Dim str����ID As String, strMsg As String
    Dim rsProperty As New ADODB.Recordset           '���Ϲ��
    Dim rsCheck As New ADODB.Recordset              '���Ͽ��
    Dim bln�¿�� As Boolean
    
    
    '��鵥���и����ϵĿ��
    'mint�����:0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    '������ʱ�۲��ϲ��ܴ���'
    
    On Error GoTo ErrHandle
    bln�¿�� = Val(zlDatabase.GetPara(95, glngSys, 0)) = 1
    
    lngRows = mshBill.Rows - 1
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln�ⷿ = CheckStockProperty(lng�ⷿID)
    
    For lngRow = 1 To lngRows
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng����ID <> 0 Then
            If InStr(1, str����ID & ",", "," & lng����ID & ",") = 0 Then str����ID = str����ID & "," & lng����ID
        End If
    Next
    
    If str����ID = "" Then
        CheckStock = True
        Exit Function
    Else
        str����ID = Mid(str����ID, 2)
    End If
    
    '��ȡ�����������в��ϵ�����
    gstrSQL = " Select A.����ID,'['||B.����||']'||B.���� ͨ����,A.�ⷿ����,A.���÷���,B.�Ƿ���" & _
              " From �������� A,�շ���ĿĿ¼ B,Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C" & _
              " Where A.����ID=B.ID And A.����ID =C.Column_Value "
    Set rsProperty = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����������в��ϵ�����", str����ID)

    '��ȡ�����������в��ϵĵ�ǰ��棨û�п��Ĳ��ϸü�¼����Ҳ�����м�¼��
    gstrSQL = " Select A.ҩƷid ����ID,Nvl(A.����,0) As ����," & _
              " SUM(NVL(��������,0)) As ��������,SUM(NVL(ʵ������,0)) As ʵ������" & _
              " From ҩƷ��� A,�շ���ĿĿ¼ B,�������� C,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) D " & _
              " Where A.�ⷿID=[1] And A.ҩƷID=B.ID And B.ID=C.����ID And A.����=1 " & _
              "         And A.ҩƷID=D.Column_Value" & _
              " Group by A.ҩƷID,Nvl(A.����,0)"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����������в��ϵĵ�ǰ���", lng�ⷿID, str����ID)
              
    '���ÿ������
    For lngRow = 1 To lngRows
        lng����ID = Val(mshBill.TextMatrix(lngRow, 0))
        If lng����ID <> 0 Then
            lng���� = Val(mshBill.TextMatrix(lngRow, mBillCol.c_����))
            dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_����ϵ��))
            dbl��д���� = Val(mshBill.TextMatrix(lngRow, mBillCol.C_ʵ������))
            
            dbl�������� = 0
            '���Ҹò��ϵĿ���¼
            rsCheck.Filter = "����ID=" & lng����ID & " And ����=" & lng����
            If rsCheck.RecordCount <> 0 Then
                If mint�༭״̬ = 10 Then '����ʱӦ���ÿ��������ж�
                    dbl�������� = zlStr.NVL(rsCheck!��������, 0) / dbl����ϵ��
                ElseIf mint�༭״̬ = 3 Then  '���ʱӦ����ʵ�������ж�
                    dbl�������� = zlStr.NVL(rsCheck!ʵ������, 0) / dbl����ϵ��
                End If
            End If
            
            '������Ŀ�����������
            If Not (dbl�������� >= dbl��д����) Then
                int����� = mint�����
                '����ò�����ʱ�ۻ��������治�㲻������⣬�൱�ڽ�ֹ����
                rsProperty.Filter = "����ID=" & lng����ID
                bln���� = (IIf(bln�ⷿ, (rsProperty!�ⷿ���� = 1), (rsProperty!���÷��� = 1)) Or (rsProperty!�Ƿ��� = 1))
                strMsg = ""
                If bln���� Then
                    int����� = 2
                    '��������β��ϣ�������С�ڵ����㣬˵��δִ�зֽ⹦��
                    If lng���� <= 0 And IIf(bln�ⷿ, (rsProperty!�ⷿ���� = 1), (rsProperty!���÷��� = 1)) Then
                        strMsg = "������ִ�зֽ⹦����ȷ���β��ϵĳ������Σ�"
                    End If
                End If
                
                If bln�¿�� = True And (mint�༭״̬ = 10 Or (mint�༭״̬ = 3 And mint�ƿ⴦������ = 0)) Then
                Else
                    '���������̽�����ʾ���ֹ
                    Select Case int�����
                    Case 1  '����ʾ
                        If MsgBox(rsProperty!ͨ���� & "�Ŀ��ÿ�治�㣬�Ƿ������" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Function
                        End If
                    Case 2
                        MsgBox rsProperty!ͨ���� & "�Ŀ��ÿ�治�㣡" & strMsg, vbInformation, gstrSysName
                        Exit Function
                    End Select
                End If
            End If
        End If
    Next
    rsCheck.Filter = 0
    rsCheck.Close
    rsProperty.Filter = 0
    rsProperty.Close
    CheckStock = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function CheckSend() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��鵱ǰ�����Ƿ��ѷ���
    On Error GoTo ErrHand
    
    gstrSQL = "Select ��ҩ���� From ҩƷ�շ���¼ " & _
              "Where ����=19 And NO=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��鵱ǰ�����Ƿ��ѷ���", Me.txtNO.Tag)
              
    If (zlStr.NVL(rsTemp!��ҩ����) = "") Then
        MsgBox "�õ����ѱ���������Աȡ�����ͣ���������գ�", vbInformation, gstrSysName
        Exit Function
    End If
    CheckSend = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.Rows < 2 Then Exit Sub
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
        
        For n = 1 To mshBill.Rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(mshBill.TextMatrix(n, mBillCol.c_���)) = 0, n, Val(mshBill.TextMatrix(n, mBillCol.c_���)))
                !ҩƷid = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mBillCol.c_����))
                
                .Update
            End If
        Next
        
    End With
End Sub
